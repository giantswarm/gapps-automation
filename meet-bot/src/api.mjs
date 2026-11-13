/**
 * HTTP control API (remote control for the Apps Script in meeting-consent-poll/).
 *
 *   GET    /healthz             liveness, unauthenticated
 *   GET    /v1/sessions         list sessions
 *   PUT    /v1/sessions/{id}    create or reschedule (idempotent): {meetingUri, title, joinAt, leaveAt, greeting, farewell}
 *   PATCH  /v1/sessions/{id}    {say: "post now (or right after joining)", farewell: "post before leaving"}
 *   DELETE /v1/sessions/{id}    cancel, or leave the meeting now
 *
 * All /v1 requests must be signed, see auth.mjs.
 */
import {createServer} from 'node:http';
import {parseSessionSpec, parseSessionUpdate} from './sessions.mjs';

const MAX_BODY_BYTES = 64 * 1024;
const SESSION_PATH = /^\/v1\/sessions\/([^/]+)$/;


export function createApiServer({sessions, verifier, log = console}) {
    return createServer(async (req, res) => {
        try {
            const path = new URL(req.url, 'http://localhost').pathname;

            if (req.method === 'GET' && path === '/healthz') {
                return send_(res, 200, {ok: true});
            }
            if (!path.startsWith('/v1/')) {
                return send_(res, 404, {error: 'not found'});
            }

            const body = await readBody_(req);
            const reason = verifier.verify(req.headers, req.method, path, body);
            if (reason) {
                log.warn(`rejected ${req.method} ${path}: ${reason}`);
                return send_(res, 401, {error: 'unauthorized'});
            }

            if (path === '/v1/sessions') {
                return req.method === 'GET'
                    ? send_(res, 200, {sessions: sessions.list()})
                    : send_(res, 405, {error: 'method not allowed'});
            }

            const match = SESSION_PATH.exec(path);
            if (!match) {
                return send_(res, 404, {error: 'not found'});
            }
            const id = decodeURIComponent(match[1]);

            switch (req.method) {
                case 'GET': {
                    const session = sessions.get(id);
                    return session ? send_(res, 200, session) : send_(res, 404, {error: 'no such session'});
                }
                case 'PUT': {
                    const result = sessions.upsert(parseSessionSpec(id, parseJson_(body)));
                    return send_(res, result.created ? 201 : 200, result.session);
                }
                case 'PATCH': {
                    const session = sessions.update(id, parseSessionUpdate(parseJson_(body)));
                    return session ? send_(res, 200, session) : send_(res, 404, {error: 'no such session'});
                }
                case 'DELETE': {
                    const session = sessions.cancel(id);
                    return session ? send_(res, 200, session) : send_(res, 404, {error: 'no such session'});
                }
                default:
                    return send_(res, 405, {error: 'method not allowed'});
            }
        } catch (e) {
            if (e.status) {
                return send_(res, e.status, {error: e.message});
            }
            log.error(`request failed: ${e.stack || e}`);
            return send_(res, 500, {error: 'internal error'});
        }
    });
}


class ClientError extends Error {
    constructor(status, message) {
        super(message);
        this.status = status;
    }
}


function readBody_(req) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        let size = 0;
        req.on('data', chunk => {
            size += chunk.length;
            if (size > MAX_BODY_BYTES) {
                reject(new ClientError(413, 'body too large'));
                req.destroy();
                return;
            }
            chunks.push(chunk);
        });
        req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
        req.on('error', reject);
    });
}


function parseJson_(body) {
    try {
        return JSON.parse(body || '{}');
    } catch (e) {
        throw new ClientError(400, 'body must be valid JSON');
    }
}


function send_(res, status, payload) {
    res.writeHead(status, {'Content-Type': 'application/json', 'Cache-Control': 'no-store'});
    res.end(JSON.stringify(payload));
}
