// export all relevant classes and variables/constants
module.exports = {
    ...module.exports,
    UrlFetchJsonClient: UrlFetchJsonClient,
    SlackWebClient: SlackWebClient,
    CalendarClient: CalendarClient,
    CalendarListClient: CalendarListClient,
    GmailClientV1: GmailClientV1,
    // OAuth2,  // exported directly via CommonJS module OAuth2.gs
    PeopleTime: PeopleTime,
    PersonioAuthV1: PersonioAuthV1,
    PersonioClientV1: PersonioClientV1,
    SheetUtil: SheetUtil,
    TriggerUtil: TriggerUtil,
    Util: Util
};

// merge exports
const ourExports = module.exports;
module = nativeModule;

for (const exp in ourExports) {
    module.exports[exp] = ourExports[exp];
}
