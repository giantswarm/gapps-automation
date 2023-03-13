/** Returns OAuth2 authenticated calendar scoped impersonation service for the specified user (by primary email).
 *
 * Requires service account credentials as exported from Google Cloud Admin Console.
 */
class GmailClientV1 extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance. */
    static withImpersonatingService(serviceAccountCredentials, primaryEmail) {
        const service = UrlFetchJsonClient.createImpersonatingService('GmailClientV1-' + primaryEmail,
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
    deleteUserMessage(userId, messageId) {
        this.fetch(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/${messageId}`, {
            method: 'delete'
        });
    }


    /**
     * Bulk delete messages (can't be undone, be careful!).
     *
     * @param {null|string} userId The mailbox to delete from ('me' if false).
     * @param {Array<string>} messageIds The ID strings of the messages to be deleted permanently.
     */
    deleteUserMessages(userId, messageIds) {
        this.postJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/batchDelete`, {
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
    getUserMessageMetadata(userId, messageId) {
        const query = GmailClientV1.buildQuery({
            format: 'metadata'
        });
        return this.getJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages/${messageId}${query}`);
    }


    /** List messages in the specified mailbox (or 'me') and accumulate those that messageFilter returns.
     *
     * If operating on large mailboxes, one may use the messageFilter to perform the actual operation and just return [],
     * to avoid accumulating a huge amount of data.
     *
     * @param {null|string} userId The mailbox to list ('me' if false).
     * @param {null|string} searchExpression The search expression (see https://support.google.com/mail/answer/7190?hl=en ).
     * @param {null|function} messageFilter Optional callback that receives each page of the message list and returns an array of those messages filtered.
     *
     * @return {Array<Object>} List of messages that passed the messageFilter.
     */
    listUserMessages(userId, searchExpression, messageFilter) {
        const messages = [];
        const filter = messageFilter ? messageFilter : (messages) => messages;
        const params = {
            q: searchExpression,
            includeSpamTrash: true,
            pageToken: undefined
        };
        do {
            const list = this.getJson(`https://gmail.googleapis.com/gmail/v1/users/${userId || 'me'}/messages${GmailClientV1.buildQuery(params)}`);

            messages.push(...filter(list.messages));

            params.pageToken = list.nextPageToken;
        }
        while (params.pageToken);

        return messages;
    }
}
