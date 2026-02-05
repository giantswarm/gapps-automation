/** A simple Google Gemini REST API client for document summarization. */
class GeminiRestClient {

    constructor(apiKey) {
        this._apiKey = apiKey;
    }

    /** Summarize a Google Doc by its content.
     *
     * @param docContent The text content of the document
     * @param prompt The prompt to use for summarization
     * @return {Promise<string>} The summarized content
     */
    async summarizeText(docContent, prompt) {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-pro-preview:generateContent?key=${this._apiKey}`;
        
        const requestBody = {
            contents: [{
                parts: [{
                    text: `${prompt}\n\nDocument content:\n${docContent}`
                }]
            }],
            generationConfig: {
                temperature: 0.7,
                topK: 40,
                topP: 0.95,
                maxOutputTokens: 2048
            }
        };

        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(requestBody),
            muteHttpExceptions: true
        };

        const response = await UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode < 200 || responseCode >= 300) {
            throw new Error(`Gemini API request failed with code ${responseCode}: ${response.getContentText()}`);
        }

        const result = JSON.parse(response.getContentText());
        
        if (!result.candidates || result.candidates.length === 0) {
            throw new Error('No response from Gemini API');
        }

        return result.candidates[0].content.parts[0].text;
    }

    /** Read and summarize a Google Doc.
     *
     * @param docId The ID of the Google Doc
     * @param prompt The prompt to use for summarization
     * @param driveClient A DriveClientV1 instance with access to the document
     * @return {Promise<string>} The summarized content
     */
    async summarizeGoogleDoc(docId, prompt, driveClient) {
        // Resolve file ID in case the file was moved (shortcut redirect)
        const resolvedDocId = await driveClient.resolveFileId(docId);

        // Export the Google Doc as plain text
        const exportUrl = `https://www.googleapis.com/drive/v3/files/${resolvedDocId}/export?mimeType=text/plain&supportsAllDrives=true`;
        const docContent = (await driveClient.fetch(exportUrl)).getContentText();

        return await this.summarizeText(docContent, prompt);
    }
}

