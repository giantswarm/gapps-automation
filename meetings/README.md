## Meetings

Reports and functionality around scheduled meetings.

### Configuration

Configuring the synchronization is possible using ScriptProperties.

The following configuration properties are available:

| Mandatory | Property Key                       | Value Example or Default                        |
|-----------|------------------------------------|-------------------------------------------------|
| **yes**   | Meetings.personioToken             | `CLIENT_ID&#124;CLIENT_SECRET`                  |
| **yes**   | Meetings.serviceAccountCredentials | `{SERVICE_ACCOUNT_CREDENTIALS_FILE_CONTENT...}` |
| no        | Meetings.allowedDomains            | `giantswarm.io,giantswarm.com`                  |
| no        | Meetings.emailWhiteList            | `jonas@giantswarm.io,marcel@giantswarm.io`      |
| no        | Meetings.lookaheadDays             | `180`                                           |
| no        | Meetings.lookbackDays              | `30`                                            |
