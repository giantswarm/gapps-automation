/** A simple Slack WebAPI client. */
const SLACK_API_URL = 'https://slack.com/api/';

class SlackWebClient {

    constructor(botToken) {
        this._botToken = botToken;
    }

    /** Get all channels the app is a member of. */
    async getUserChannels() {
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

    /** Post a message to a single channel. */
    async postMessage(channelId, text, blocks) {
        const payload = {
            channel: channelId,
            text: text
        };

        if (blocks) {
            payload.blocks = blocks;
        }

        const options = {
            method: 'post',
            contentType: 'application/json; charset=utf-8',
            headers: {
                Authorization: `Bearer ${this._botToken}`
            },
            payload: JSON.stringify(payload)
        };
        const response = await UrlFetchApp.fetch(`${SLACK_API_URL}chat.postMessage`, options);
        const postMessageResponse = JSON.parse(response.getContentText());
        if (!postMessageResponse.ok) {
            throw new Error('Failed to post message to Slack channel ' + channelId + ': ' + postMessageResponse.error);
        }
        return postMessageResponse;
    }

    /** Broadcast message to all channels the app has been added to. */
    async broadcastMessage(text, blocks) {
        const channels = await this.getUserChannels();
        for (const channel of channels) {
            await this.postMessage(channel.id, text, blocks);
        }
    }

    /** Look up a user by their email address. */
    async lookupByEmail(email) {
        const options = {
            method: 'get',
            contentType: 'application/x-www-form-urlencoded',
            headers: {
                Authorization: `Bearer ${this._botToken}`
            }
        };
        const response = await UrlFetchApp.fetch(`${SLACK_API_URL}users.lookupByEmail?email=${encodeURIComponent(email)}`, options);
        const lookupResponse = JSON.parse(response.getContentText());
        if (!lookupResponse.ok) {
            return null; // User not found or other error
        }
        return lookupResponse.user;
    }
}
