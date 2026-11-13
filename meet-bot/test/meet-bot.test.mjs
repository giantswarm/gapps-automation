import assert from 'node:assert/strict';
import {test} from 'node:test';
import {createApiServer} from '../src/api.mjs';
import {RequestVerifier, sign} from '../src/auth.mjs';
import {splitChatMessage} from '../src/meet.mjs';
import {parseSessionSpec, SessionManager} from '../src/sessions.mjs';

const SECRET = 'x'.repeat(40);
const quiet = {info: () => undefined, warn: () => undefined, error: () => undefined};
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));


function signedHeaders(method, path, body, timestamp = Math.floor(Date.now() / 1000)) {
    return {
        'x-meetbot-timestamp': '' + timestamp,
        'x-meetbot-signature': 'sha256=' + sign(SECRET, '' + timestamp, method, path, body)
    };
}


/** Records everything the bot would do in Meet. */
class FakeDriver {
    constructor({failJoin = false} = {}) {
        this.failJoin = failJoin;
        this.said = [];
        this.joined = [];
        this.left = 0;
        this.active = true;
    }

    async join(session) {
        if (this.failJoin) {
            throw new Error('bot account is signed out');
        }
        this.joined.push(session.meetingUri);
        return {
            say: async text => void this.said.push(text),
            isActive: async () => this.active,
            leave: async () => void ++this.left
        };
    }
}


test('verifier accepts signed requests once and rejects tampering, skew and replays', () => {
    const verifier = new RequestVerifier(SECRET);
    const body = '{"a":1}';
    const headers = signedHeaders('PUT', '/v1/sessions/e1', body);

    assert.equal(verifier.verify(headers, 'PUT', '/v1/sessions/e1', body), null);
    assert.equal(verifier.verify(headers, 'PUT', '/v1/sessions/e1', body), 'replayed request');
    assert.equal(verifier.verify(signedHeaders('PUT', '/v1/sessions/e1', body), 'PUT', '/v1/sessions/e2', body), 'invalid signature');
    assert.equal(verifier.verify(signedHeaders('PUT', '/v1/sessions/e1', body), 'PUT', '/v1/sessions/e1', '{"a":2}'), 'invalid signature');
    assert.equal(verifier.verify(signedHeaders('GET', '/v1/sessions', '', Math.floor(Date.now() / 1000) - 600), 'GET', '/v1/sessions', ''),
        'timestamp outside of the accepted window');
    assert.equal(verifier.verify({}, 'GET', '/v1/sessions', ''), 'missing signature headers');
    assert.throws(() => new RequestVerifier('short'));
});


test('session specs only accept Meet links and sane times', () => {
    const now = Date.parse('2026-10-29T10:00:00Z');
    const spec = parseSessionSpec('abc_20261029T100000Z', {
        meetingUri: 'https://meet.google.com/abc-defg-hij?authuser=0',
        joinAt: '2026-10-29T10:05:00Z'
    }, now);
    assert.equal(spec.meetingUri, 'https://meet.google.com/abc-defg-hij');
    assert.equal(spec.leaveAt, Date.parse('2026-10-29T10:15:00Z'));

    for (const meetingUri of ['https://evil.example/abc-defg-hij', 'http://meet.google.com/abc-defg-hij',
        'https://meet.google.com.evil.example/abc-defg-hij', 'https://meet.google.com/lookup/x']) {
        assert.throws(() => parseSessionSpec('e1', {meetingUri}, now), /meetingUri/);
    }
    assert.throws(() => parseSessionSpec('../x', {meetingUri: 'https://meet.google.com/abc-defg-hij'}, now), /session id/);
    assert.throws(() => parseSessionSpec('e1', {meetingUri: 'https://meet.google.com/abc-defg-hij', joinAt: '2026-11-02T10:00:00Z'}, now), /joinAt/);
    assert.throws(() => parseSessionSpec('e1', {meetingUri: 'https://meet.google.com/abc-defg-hij', leaveAt: '2026-10-29T09:00:00Z'}, now), /leaveAt/);
});


test('session joins, greets, relays messages, says farewell and leaves', async () => {
    const driver = new FakeDriver();
    const sessions = new SessionManager({driver, pollIntervalMs: 10, log: quiet});
    const spec = parseSessionSpec('e1', {
        meetingUri: 'https://meet.google.com/abc-defg-hij',
        leaveAt: new Date(Date.now() + 300).toISOString(),
        greeting: 'hello',
        farewell: 'bye'
    });

    assert.equal(sessions.upsert(spec).created, true);
    assert.equal(sessions.upsert(spec).created, false);   // idempotent
    await sleep(50);
    assert.equal(sessions.get('e1').state, 'in-call');

    sessions.update('e1', {say: 'someone objected', farewell: 'tally: 2/3'});
    await sleep(50);
    assert.deepEqual(driver.said, ['hello', 'someone objected']);

    await sleep(300);
    assert.equal(sessions.get('e1').state, 'left');
    assert.deepEqual(driver.said, ['hello', 'someone objected', 'tally: 2/3']);
    assert.equal(driver.left, 1);
    assert.deepEqual(driver.joined, ['https://meet.google.com/abc-defg-hij']);

    // re-sending the spec after the meeting must not make the bot join again
    assert.equal(sessions.upsert(spec).session.state, 'left');
    await sleep(30);
    assert.equal(driver.joined.length, 1);
});


test('session leaves early without farewell when cancelled, and when the call ends', async () => {
    const driver = new FakeDriver();
    const sessions = new SessionManager({driver, pollIntervalMs: 10, log: quiet});
    const leaveAt = new Date(Date.now() + 60000).toISOString();
    sessions.upsert(parseSessionSpec('e1', {meetingUri: 'https://meet.google.com/abc-defg-hij', leaveAt, farewell: 'bye'}));
    sessions.upsert(parseSessionSpec('e2', {meetingUri: 'https://meet.google.com/abc-defg-hik', leaveAt, farewell: 'bye'}));
    await sleep(30);

    sessions.cancel('e1');
    await sleep(30);
    assert.equal(sessions.get('e1').state, 'left');
    assert.deepEqual(driver.said, []);

    driver.active = false;
    await sleep(30);
    assert.equal(sessions.get('e2').state, 'left');
    assert.deepEqual(driver.said, []);
    assert.equal(driver.left, 2);
});


test('scheduled sessions can be rescheduled and cancelled, failures are reported', async () => {
    const sessions = new SessionManager({driver: new FakeDriver({failJoin: true}), log: quiet});
    const meetingUri = 'https://meet.google.com/abc-defg-hij';

    sessions.upsert(parseSessionSpec('later', {meetingUri, joinAt: new Date(Date.now() + 60000).toISOString()}));
    const moved = sessions.upsert(parseSessionSpec('later', {meetingUri, joinAt: new Date(Date.now() + 120000).toISOString()}));
    assert.equal(moved.session.state, 'scheduled');
    assert.equal(sessions.cancel('later').state, 'cancelled');

    sessions.upsert(parseSessionSpec('now', {meetingUri}));
    await sleep(30);
    assert.equal(sessions.get('now').state, 'failed');
    assert.match(sessions.get('now').error, /signed out/);

    // failed sessions are retried when the spec is sent again
    assert.equal(sessions.upsert(parseSessionSpec('now', {meetingUri})).created, true);
    await sessions.shutdown();
});


test('chat messages are split at line boundaries below the Meet limit', () => {
    assert.deepEqual(splitChatMessage('a\nb', 10), ['a\nb']);
    assert.deepEqual(splitChatMessage('aaaa\nbbbb\ncccc', 9), ['aaaa\nbbbb', 'cccc']);
    assert.deepEqual(splitChatMessage('one two three', 8), ['one two', 'three']);
    for (const chunk of splitChatMessage(('word '.repeat(40) + '\n').repeat(10))) {
        assert.ok(chunk.length <= 500);
    }
});


test('API authenticates and routes requests', async () => {
    const sessions = new SessionManager({driver: new FakeDriver(), log: quiet});
    const server = createApiServer({sessions, verifier: new RequestVerifier(SECRET), log: quiet}).listen(0);
    await new Promise(resolve => server.once('listening', resolve));
    const base = `http://127.0.0.1:${server.address().port}`;

    const call = (method, path, payload, headers) => {
        const body = payload === undefined ? '' : JSON.stringify(payload);
        return fetch(base + path, {
            method,
            body: body || undefined,
            headers: headers || signedHeaders(method, path, body)
        });
    };

    try {
        assert.equal((await fetch(base + '/healthz')).status, 200);
        assert.equal((await call('GET', '/v1/sessions', undefined, {})).status, 401);

        const joinAt = new Date(Date.now() + 60000).toISOString();
        const put = await call('PUT', '/v1/sessions/e1', {meetingUri: 'https://meet.google.com/abc-defg-hij', joinAt, greeting: 'hi'});
        assert.equal(put.status, 201);
        assert.equal((await put.json()).state, 'scheduled');

        assert.equal((await call('PUT', '/v1/sessions/e2', {meetingUri: 'https://example.com/'})).status, 400);
        assert.equal((await call('PATCH', '/v1/sessions/e1', {say: 'hello again'})).status, 200);
        assert.equal((await call('PATCH', '/v1/sessions/missing', {say: 'x'})).status, 404);

        const list = await (await call('GET', '/v1/sessions')).json();
        assert.deepEqual(list.sessions.map(s => [s.id, s.pendingMessages]), [['e1', 2]]);

        assert.equal((await (await call('DELETE', '/v1/sessions/e1')).json()).state, 'cancelled');
    } finally {
        server.close();
        await sessions.shutdown();
    }
});
