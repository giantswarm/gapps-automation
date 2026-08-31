/** Returns OAuth2 authenticated calendar scoped impersonation service for the specified user (by primary email).
 *
 * Requires service account credentials as exported from Google Cloud Admin Console.
 */
class GmailClientV1 extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance. */
    static async withImpersonatingService(serviceAccountCredentials, primaryEmail) {
        const service = await UrlFetchJsonClient.createImpersonatingService('GmailClientV1-' + primaryEmail,
            serviceAccountCredentials,
            primaryEmail,
            'https://mail.google.com/');

        return new GmailClientV1(service);
    }


    /**
     * Delete message (can't be undone, be careful!).
     *
     * @param {null|string} userId The mailbox to delete from ('me' if false).
     * @param {string} messageId The ID of the message to delete permanently.
     */
    async deleteUserMessage(userId, messageId) {
        await this.fetch(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/${messageId}`, {
            method: 'delete'
        });
    }


    /**
     * Bulk delete messages (can't be undone, be careful!).
     *
     * @param {null|string} userId The mailbox to delete from ('me' if false).
     * @param {Array<string>} messageIds The ID strings of the messages to be deleted permanently.
     */
    async deleteUserMessages(userId, messageIds) {
        await this.postJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/batchDelete`, {
            ids: messageIds
        });
    }


    /**
     * Retrieve message metadata (for debugging/development).
     *
     * Do not use on user mailboxes.
     *
     * @param {null|string} userId The mailbox to delete from ('me' if false).
     * @param {string} messageId The ID strings of the messages to be deleted permanently.
     *
     * @return Message metadata, including headers
     */
    async getUserMessageMetadata(userId, messageId) {
        const query = GmailClientV1.buildQuery({
            format: 'metadata'
        });
        return await this.getJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/${messageId}${query}`);
    }


    /**
     * Retrieve a single message, including headers and body parts.
     *
     * @param {null|string} userId The mailbox to read from ('me' if false).
     * @param {string} messageId The ID of the message to retrieve.
     * @param {null|string} format One of: minimal, full, raw, metadata ('full' if false).
     *
     * @return {Object} The message resource.
     */
    async getUserMessage(userId, messageId, format) {
        const query = GmailClientV1.buildQuery({
            format: format || 'full'
        });
        return await this.getJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/${messageId}${query}`);
    }


    /**
     * Send a message on behalf of the specified mailbox.
     *
     * The From address must be the mailbox itself or one of its verified "send as" aliases,
     * otherwise Gmail silently rewrites it.
     *
     * @param {null|string} userId The sending mailbox ('me' if false).
     * @param {string} rawMessage The complete RFC 5322 message, base64url (web-safe) encoded.
     * @param {null|string} threadId Optional ID of the thread to attach the message to.
     *
     * @return {Object} The sent message resource.
     */
    async sendUserMessage(userId, rawMessage, threadId) {
        return await this.postJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/send`, {
            raw: rawMessage,
            threadId: threadId || undefined
        });
    }


    /**
     * Add and/or remove labels of a single message.
     *
     * @param {null|string} userId The mailbox owning the message ('me' if false).
     * @param {string} messageId The ID of the message to modify.
     * @param {null|Array<string>} addLabelIds IDs of the labels to add.
     * @param {null|Array<string>} removeLabelIds IDs of the labels to remove.
     *
     * @return {Object} The modified message resource.
     */
    async modifyUserMessage(userId, messageId, addLabelIds, removeLabelIds) {
        return await this.postJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/${messageId}/modify`, {
            addLabelIds: addLabelIds || [],
            removeLabelIds: removeLabelIds || []
        });
    }


    /** List all labels of the specified mailbox.
     *
     * @param {null|string} userId The mailbox to list labels of ('me' if false).
     *
     * @return {Array<Object>} The list of label resources.
     */
    async listUserLabels(userId) {
        const response = await this.getJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/labels`);
        return (response || {}).labels || [];
    }


    /** Create a new user label.
     *
     * @param {null|string} userId The mailbox to create the label in ('me' if false).
     * @param {string} name The label name ('/' separated for nested labels).
     *
     * @return {Object} The created label resource.
     */
    async createUserLabel(userId, name) {
        return await this.postJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/labels`, {
            name: name,
            labelListVisibility: 'labelShow',
            messageListVisibility: 'show'
        });
    }


    /** Get the ID of a user label, creating the label if it doesn't exist yet.
     *
     * Label names are unique ignoring case, hence existing labels are matched ignoring case.
     *
     * @param {null|string} userId The mailbox owning the label ('me' if false).
     * @param {string} name The label name ('/' separated for nested labels).
     *
     * @return {string} The label ID.
     */
    async ensureUserLabel(userId, name) {
        const findLabel = labels => labels.find(label => (label.name || '').toLowerCase() === name.toLowerCase());

        const existing = findLabel(await this.listUserLabels(userId));
        if (existing) {
            return existing.id;
        }

        try {
            return (await this.createUserLabel(userId, name)).id;
        } catch (e) {
            // may have been created concurrently (Gmail answers 409 in that case)
            const created = findLabel(await this.listUserLabels(userId));
            if (!created) {
                throw e;
            }
            return created.id;
        }
    }


    /** Get the value of the first header with the specified name (compared ignoring case).
     *
     * @param {null|Array<Object>} headers The headers array of a message or message part.
     * @param {string} name The header name to look for.
     *
     * @return {null|string} The header value or null if there is no such header.
     */
    static getHeader(headers, name) {
        const wantedName = (name || '').toLowerCase();
        for (const header of headers || []) {
            if ((header.name || '').toLowerCase() === wantedName) {
                return header.value || '';
            }
        }

        return null;
    }


    /** Decode the (base64url encoded) body data of a single message part.
     *
     * The charset announced by the part's Content-Type header is honored, defaulting to UTF-8.
     *
     * @param {null|Object} part A message part (or a message payload).
     *
     * @return {string} The decoded part body (empty string if the part has no inline data).
     */
    static decodePartData(part) {
        const data = ((part || {}).body || {}).data;
        if (!data) {
            return '';
        }

        // the Gmail API omits base64 padding
        const padded = data + '='.repeat((4 - (data.length % 4)) % 4);
        const blob = Utilities.newBlob(Utilities.base64DecodeWebSafe(padded));

        const charset = (GmailClientV1.getHeader((part || {}).headers, 'Content-Type') || '')
            .match(/charset\s*=\s*"?([\w-]+)"?/i);
        if (charset) {
            try {
                return blob.getDataAsString(charset[1]);
            } catch (e) {
                // fall through to the UTF-8 default below
            }
        }

        return blob.getDataAsString();
    }


    /** Get the human readable body text of a message.
     *
     * Prefers text/plain parts over text/html ones and never descends into attached messages.
     *
     * @param {null|Object} payload The payload of a message (as returned with format 'full').
     *
     * @return {{text: string, isHtml: boolean}} The first body part found (empty text if there is none).
     */
    static getMessageBody(payload) {
        let html = null;

        const findText = part => {
            if (!part || part.mimeType === 'message/rfc822') {
                return null;  // attached messages are not part of this message's body
            }

            if (part.mimeType === 'text/plain' && ((part.body || {}).data)) {
                return GmailClientV1.decodePartData(part);
            }

            if (part.mimeType === 'text/html' && ((part.body || {}).data) && html === null) {
                html = GmailClientV1.decodePartData(part);
            }

            for (const child of part.parts || []) {
                const text = findText(child);
                if (text) {
                    return text;
                }
            }

            return null;
        };

        const text = findText(payload);

        return text ? {text: text, isHtml: false} : {text: html || '', isHtml: !!html};
    }


    /** Depth first search for the first message part accepted by the predicate.
     *
     * @param {null|Object} part A message part (or a message payload).
     * @param {function} predicate Callback deciding whether a part matches.
     *
     * @return {null|Object} The first matching part or null.
     */
    static findPart(part, predicate) {
        if (!part) {
            return null;
        }

        if (predicate(part)) {
            return part;
        }

        for (const child of part.parts || []) {
            const match = GmailClientV1.findPart(child, predicate);
            if (match) {
                return match;
            }
        }

        return null;
    }


    /** List messages in the specified mailbox (or 'me') and accumulate those that messageFilter returns.
     *
     * If operating on large mailboxes, one may use the messageFilter to perform the actual operation and just return [],
     * to avoid accumulating a huge amount of data.
     *
     * @param {null|string} userId The mailbox to list ('me' if false).
     * @param {null|string} searchExpression The search expression (see https://support.google.com/mail/answer/7190?hl=en ).
     * @param {null|function} messageFilter Optional callback that receives each page of the message list and returns an array of those messages filtered.
     * @param {undefined|boolean} includeSpamTrash Whether to also list messages in spam and trash (defaults to true).
     *
     * @return {Array<Object>} List of messages that passed the messageFilter.
     */
    async listUserMessages(userId, searchExpression, messageFilter, includeSpamTrash) {
        const messages = [];
        const filter = messageFilter ? messageFilter : async (messages) => messages;
        const params = {
            q: searchExpression,
            includeSpamTrash: includeSpamTrash === undefined ? true : !!includeSpamTrash,
            pageToken: undefined
        };
        do {
            const list = await this.getJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages${GmailClientV1.buildQuery(params)}`);

            // the Gmail API omits the messages field on empty pages
            messages.push(...(await filter(list.messages || []) || []));

            params.pageToken = list.nextPageToken;
        }
        while (params.pageToken);

        return messages;
    }
}
