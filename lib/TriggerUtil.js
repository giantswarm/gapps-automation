/** Apps Script trigger related helper functions.
 *
 * IMPORTANT: Note Apps Script resource scoping:
 *   https://developers.google.com/apps-script/guides/libraries#resource_scoping
 */


/** The permitted trigger delays according to:
 *
 * https://developers.google.com/apps-script/reference/script/clock-trigger-builder#everyMinutes(Integer)
 */
const TRIGGER_DELAY_MINUTES = [1, 5, 10, 15, 30];

class TriggerUtil {

    /** Sanitize delay minutes input.
     *
     * ClockTriggerBuilder supported only a limited number of values, see:
     * https://developers.google.com/apps-script/reference/script/clock-trigger-builder#everyMinutes(Integer)
     *
     * @param {number} delayMinutes The intended trigger delay in minutes.
     *
     * @return {number} sanitized trigger delay (within supported range).
     */
    static sanitizeTriggerDelayMinutes(delayMinutes) {
        return TRIGGER_DELAY_MINUTES.reduceRight((v, prev) =>
            typeof +delayMinutes === 'number' && v <= +delayMinutes ? v : prev);
    }


    /** Uninstall time based execution trigger for this script.
     *
     * @param handler The name of the handler function whose associated triggers shall be removed.
     */
    static uninstall(handler) {
        // Remove pre-existing triggers
        const triggers = ScriptApp.getProjectTriggers();
        for (const trigger of triggers) {
            if (trigger.getHandlerFunction() === handler) {
                ScriptApp.deleteTrigger(trigger);
                Logger.log("Uninstalled time based trigger for %s", TRIGGER_HANDLER_FUNCTION);
            }
        }
    }


    /** Setup for periodic execution and do some checks.
     *
     * @param handler The name of the handler function to be called on trigger event.
     * @param delayMinutes The delay in minutes: 1, 5, 10, 15 or 30
     */
    static install(handler, delayMinutes) {
        TriggerUtil.uninstall();

        const delay = TriggerUtil.sanitizeDelayMinutes(delayMinutes);
        Logger.log("Installing time based trigger (every %s minutes)", delayMinutes);

        ScriptApp.newTrigger(TRIGGER_HANDLER_FUNCTION)
            .timeBased()
            .everyMinutes(delay)
            .create();

        Logger.log("Installed time based trigger for %s every %s minutes", TRIGGER_HANDLER_FUNCTION, delay);
    }


    /** Helper function to configure the required script properties.
     *
     * USAGE EXAMPLE:
     *   clasp run 'setProperties' --params '[{"PersonioDump.SHEET_ID": "SOME_PERSONIO_URL|CLIENT_ID|CLIENT_SECRET"}, false]'
     *
     * Warning: passing argument true for parameter deleteAllOthers will also cause the schema to be reset!
     */
    static setProperties(properties, deleteAllOthers) {
        PropertiesService.getScriptProperties().setProperties(properties, deleteAllOthers);
    }
}
