/**
 * Meeting sessions: one per calendar event, scheduled, joined, talked into and left by a driver.
 *
 * The driver (see meet.mjs) is injected so the scheduling logic can be tested without a browser:
 *   driver.join(session) -> call
 *   call.say(text), call.isActive() -> boolean, call.leave()
 */

/** Only plain Google Meet meeting codes are accepted, the bot never browses anywhere else. */
const MEETING_URI_PATTERN = /^https:\/\/meet\.google\.com\/([a-z]{3}-[a-z]{4}-[a-z]{3})(?:[/?#].*)?$/;

/** Calendar event IDs (recurring instances contain '_'). */
const SESSION_ID_PATTERN = /^[A-Za-z0-9_-]{1,256}$/;

const MAX_TEXT_LENGTH = 4000;
const MAX_SESSION_MS = 4 * 60 * 60 * 1000;
const MAX_SCHEDULE_AHEAD_MS = 24 * 60 * 60 * 1000;
const DEFAULT_STAY_MS = 10 * 60 * 1000;

/** Invalid client input (HTTP 400). */
export class ValidationError extends Error {
    constructor(message) {
        super(message);
        this.status = 400;
    }
}

/** States a session never leaves. */
const FINAL_STATES = new Set(['left', 'failed', 'cancelled']);


/** Validate and normalize a session spec (PUT body), throws an Error with a client facing message. */
export function parseSessionSpec(id, body, now = Date.now()) {
    if (!SESSION_ID_PATTERN.test(id || '')) {
        throw new ValidationError('invalid session id');
    }
    if (!body || typeof body !== 'object') {
        throw new ValidationError('body must be a JSON object');
    }

    const match = MEETING_URI_PATTERN.exec(body.meetingUri || '');
    if (!match) {
        throw new ValidationError('meetingUri must be a https://meet.google.com/xxx-xxxx-xxx link');
    }

    const joinAt = body.joinAt ? Date.parse(body.joinAt) : now;
    if (Number.isNaN(joinAt) || joinAt > now + MAX_SCHEDULE_AHEAD_MS) {
        throw new ValidationError('joinAt must be an ISO timestamp within the next 24 hours');
    }
    const leaveAt = body.leaveAt ? Date.parse(body.leaveAt) : Math.max(joinAt, now) + DEFAULT_STAY_MS;
    if (Number.isNaN(leaveAt) || leaveAt <= now || leaveAt <= joinAt || leaveAt - Math.max(joinAt, now) > MAX_SESSION_MS) {
        throw new ValidationError('leaveAt must be an ISO timestamp after joinAt and now, at most 4 hours after joining');
    }

    return {
        id: id,
        meetingUri: 'https://meet.google.com/' + match[1],
        title: parseText_(body.title, 'title') || '',
        joinAt: joinAt,
        leaveAt: leaveAt,
        greeting: parseText_(body.greeting, 'greeting'),
        farewell: parseText_(body.farewell, 'farewell')
    };
}


/** Validate a session update (PATCH body). */
export function parseSessionUpdate(body) {
    if (!body || typeof body !== 'object') {
        throw new ValidationError('body must be a JSON object');
    }
    const update = {say: parseText_(body.say, 'say'), farewell: parseText_(body.farewell, 'farewell')};
    if (update.say === null && update.farewell === null) {
        throw new ValidationError('nothing to update, expected "say" and/or "farewell"');
    }
    return update;
}


function parseText_(value, name) {
    if (value === undefined || value === null || value === '') {
        return null;
    }
    if (typeof value !== 'string' || value.length > MAX_TEXT_LENGTH) {
        throw new ValidationError(name + ' must be a string of at most ' + MAX_TEXT_LENGTH + ' characters');
    }
    return value;
}


export class SessionManager {

    constructor({driver, maxActive = 4, pollIntervalMs = 5000, retentionMs = 48 * 60 * 60 * 1000, log = console, now = () => Date.now()}) {
        this.driver = driver;
        this.maxActive = maxActive;
        this.pollIntervalMs = pollIntervalMs;
        this.retentionMs = retentionMs;
        this.log = log;
        this.now = now;
        this.sessions = new Map();
    }


    /** Create a session, or update a still scheduled one (e.g. the meeting moved). Idempotent.
     *
     * Sessions that already joined or ended are returned unchanged, so callers may re-send the same
     * spec on every run without making the bot join twice. Failed sessions are retried.
     */
    upsert(spec) {
        this.cleanup_();

        const existing = this.sessions.get(spec.id);
        if (existing && existing.state !== 'failed') {
            if (existing.state === 'scheduled') {
                clearTimeout(existing.timer);
                Object.assign(existing, spec, {queue: spec.greeting ? [spec.greeting] : []});
                this.schedule_(existing);
            }
            return {session: view_(existing), created: false};
        }

        const session = Object.assign({}, spec, {
            state: 'scheduled',
            queue: spec.greeting ? [spec.greeting] : [],
            createdAt: this.now(),
            joinedAt: null,
            endedAt: null,
            error: null,
            leaveRequested: false,
            timer: null,
            wake: null
        });
        this.sessions.set(spec.id, session);
        this.schedule_(session);
        this.log.info(`session ${session.id}: scheduled for ${new Date(session.joinAt).toISOString()} (${session.meetingUri})`);
        return {session: view_(session), created: true};
    }


    /** Queue a chat message and/or replace the farewell message. */
    update(id, update) {
        const session = this.sessions.get(id);
        if (!session) {
            return null;
        }
        if (!FINAL_STATES.has(session.state)) {
            if (update.say) {
                session.queue.push(update.say);
            }
            if (update.farewell) {
                session.farewell = update.farewell;
            }
            this.wake_(session);
        }
        return view_(session);
    }


    /** Cancel a scheduled session or make the bot leave the meeting (without the farewell message). */
    cancel(id) {
        const session = this.sessions.get(id);
        if (!session) {
            return null;
        }
        if (session.state === 'scheduled') {
            clearTimeout(session.timer);
            this.end_(session, 'cancelled');
        } else if (!FINAL_STATES.has(session.state)) {
            session.leaveRequested = true;
            this.wake_(session);
        }
        return view_(session);
    }


    get(id) {
        const session = this.sessions.get(id);
        return session ? view_(session) : null;
    }


    list() {
        this.cleanup_();
        return [...this.sessions.values()].map(view_);
    }


    /** Leave all meetings (used on SIGTERM). */
    async shutdown() {
        const running = [];
        for (const session of this.sessions.values()) {
            clearTimeout(session.timer);
            if (session.state === 'scheduled') {
                this.end_(session, 'cancelled');
            } else if (!FINAL_STATES.has(session.state)) {
                session.leaveRequested = true;
                this.wake_(session);
                running.push(session.done);
            }
        }
        await Promise.allSettled(running);
    }


    schedule_(session) {
        const delay = Math.max(0, session.joinAt - this.now());
        session.timer = setTimeout(() => {
            session.timer = null;
            session.done = this.run_(session);
        }, delay);
    }


    async run_(session) {
        const active = [...this.sessions.values()].filter(s => s.state === 'joining' || s.state === 'in-call').length;
        if (active >= this.maxActive) {
            this.end_(session, 'failed', new Error(`too many active sessions (${active})`));
            return;
        }
        if (this.now() >= session.leaveAt) {
            this.end_(session, 'failed', new Error('leaveAt passed before joining'));
            return;
        }

        session.state = 'joining';
        let call = null;
        try {
            call = await this.driver.join(session);
            session.state = 'in-call';
            session.joinedAt = this.now();
            this.log.info(`session ${session.id}: joined ${session.meetingUri}`);

            while (!session.leaveRequested && this.now() < session.leaveAt) {
                await this.flush_(session, call);
                if (!await call.isActive()) {
                    this.log.info(`session ${session.id}: call ended or bot was removed`);
                    break;
                }
                await this.sleep_(session, Math.min(this.pollIntervalMs, session.leaveAt - this.now()));
            }

            if (!session.leaveRequested && await call.isActive()) {
                await this.flush_(session, call);
                if (session.farewell) {
                    await call.say(session.farewell);
                }
            }
            this.end_(session, 'left');
        } catch (e) {
            this.end_(session, 'failed', e);
        } finally {
            if (call) {
                await call.leave().catch(e => this.log.warn(`session ${session.id}: failed to leave cleanly: ${e.message}`));
            }
        }
    }


    async flush_(session, call) {
        while (session.queue.length > 0) {
            await call.say(session.queue[0]);
            session.queue.shift();
        }
    }


    sleep_(session, ms) {
        return new Promise(resolve => {
            const timer = setTimeout(done, Math.max(0, ms));
            function done() {
                clearTimeout(timer);
                session.wake = null;
                resolve();
            }
            session.wake = done;
        });
    }


    wake_(session) {
        if (session.wake) {
            session.wake();
        }
    }


    end_(session, state, error) {
        session.state = state;
        session.endedAt = this.now();
        session.queue = [];
        if (error) {
            session.error = error.message || '' + error;
            this.log.error(`session ${session.id}: ${state}: ${session.error}`);
        } else {
            this.log.info(`session ${session.id}: ${state}`);
        }
    }


    cleanup_() {
        const now = this.now();
        for (const [id, session] of this.sessions) {
            if (FINAL_STATES.has(session.state) && now - session.endedAt > this.retentionMs) {
                this.sessions.delete(id);
            }
        }
    }
}


/** The public representation of a session (no message contents). */
function view_(session) {
    return {
        id: session.id,
        meetingUri: session.meetingUri,
        title: session.title,
        state: session.state,
        joinAt: new Date(session.joinAt).toISOString(),
        leaveAt: new Date(session.leaveAt).toISOString(),
        joinedAt: session.joinedAt ? new Date(session.joinedAt).toISOString() : null,
        endedAt: session.endedAt ? new Date(session.endedAt).toISOString() : null,
        pendingMessages: session.queue.length,
        error: session.error
    };
}
