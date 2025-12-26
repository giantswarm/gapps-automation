/** Returns OAuth2 authenticated drive scoped impersonation service for the specified user (by primary email).
 *
 * Requires service account credentials as exported from Google Cloud Admin Console.
 */
class DriveClientV1 extends UrlFetchJsonClient {

    constructor(service) {
        super(service);
    }


    /** Helper to create an impersonating instance. */
    static async withImpersonatingService(serviceAccountCredentials, primaryEmail) {
        const service = await UrlFetchJsonClient.createImpersonatingService('DriveClient-' + primaryEmail,
            serviceAccountCredentials,
            primaryEmail,
            'https://www.googleapis.com/auth/drive');

        return new DriveClientV1(service);
    }


    /** Share a file with others.
     *
     * @param fileId The ID of the file to share
     * @param type The type of audience (e.g., 'user', 'group', 'domain', 'anyone')
     * @param value The actual email or domain in question
     * @param role The role to grant (e.g., 'reader', 'commenter', 'writer')
     * @param withLink Whether the link is required to access
     * @return {Promise<*>} The created permission
     */
    async shareWith(fileId, type, value, role = 'reader', withLink = true) {
        const permission = {
            type: type,
            role: role,
            value: value,
            withLink: withLink
        };

        if (type === 'domain') {
            permission.domain = value;
        }

        return await this.postJson(
            `https://www.googleapis.com/drive/v3/files/${fileId}/permissions`,
            permission
        );
    }


    /** Get file metadata.
     *
     * @param fileId The ID of the file
     * @param fields Optional fields to retrieve (e.g., 'id,name,mimeType,permissions')
     * @return {Promise<*>} The file metadata
     */
    async getFile(fileId, fields) {
        const query = fields ? DriveClientV1.buildQuery({ fields }) : '';
        return await this.getJson(`https://www.googleapis.com/drive/v3/files/${fileId}${query}`);
    }

    /** Extract file ID from a Google Drive URL.
     *
     * @param url The Google Drive URL
     * @return {string|null} The file ID or null if not found
     */
    static extractFileId(url) {
        if (!url) return null;

        // Match patterns like:
        // https://docs.google.com/document/d/{fileId}/edit
        // https://drive.google.com/file/d/{fileId}/view
        // https://docs.google.com/file/d/{fileId}/view
        const patterns = [
            /\/d\/([a-zA-Z0-9_-]+)/,
            /id=([a-zA-Z0-9_-]+)/
        ];

        for (const pattern of patterns) {
            const match = url.match(pattern);
            if (match) {
                return match[1];
            }
        }

        return null;
    }
}


