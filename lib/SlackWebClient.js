/** A simple Slack WebAPI client. */
const SLACK_API_URL = 'https://slack.com/api/';

class SlackWebClient {

    constructor(botToken) {
        this._botToken = botToken;
    }

    /** Get all conversations the app is a member of. */
    async getUserConversations() {
        let channels = [];
        let nextCursor = null;
        do {
            const options = {
                method: 'post',
                contentType: 'application/x-www-form-urlencoded',
                payload: {
                    token: this._botToken,
                    limit: 100,
                    cursor: nextCursor,
                    types: 'public_channel,private_channel',
                    exclude_archived: true
                }
            };
            const response = await UrlFetchApp.fetch(`${SLACK_API_URL}users.conversations`, options);
            const channelResponse = JSON.parse(response.getContentText());
            if (!channelResponse.ok) {
                throw new Error('Failed to enumerate Slack channels');
            }

            channels = channels.concat(channelResponse.channels);

            nextCursor = channelResponse.response_metadata.next_cursor;
        }
        while (nextCursor);

        return channels;
    }

    /** Post a message to a single conversation. */
    async postMessage(conversationId, text) {
        const options = {
            method: 'post',
            contentType: 'application/json; charset=utf-8',
            headers: {
                Authorization: `Bearer ${this._botToken}`
            },
            payload: JSON.stringify({
                channel: conversationId,
                text: text
            })
        };
        const response = await UrlFetchApp.fetch(`${SLACK_API_URL}chat.postMessage`, options);
        const postMessageResponse = JSON.parse(response.getContentText());
        if (!postMessageResponse.ok) {
            throw new Error('Failed to post message to Slack channel ' + conversationId);
        }
    }

    /** Broadcast message to all conversations the app is a member of. */
    async broadcastMessage(text) {
        const channels = await this.getUserConversations();
        for (const channel of channels) {
            await this.postMessage(channel.id, text);
        }
    }
}
