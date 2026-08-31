## Domain Harvester

Harvests external email domains out of Slack workspace user profiles and mails a report of the
domains not seen before, so they can be reviewed and added to a Gmail **address list** (for example
`urgent-group-allowed-sender-domains`, which backs the allowlist for a shared support escalation
mailbox alias): see [giantswarm/giantswarm#37417](https://github.com/giantswarm/giantswarm/issues/37417).

### Why this doesn't add the domains itself

Address lists (Admin console → Apps → Google Workspace → Gmail → Advanced settings →
[Manage address lists](https://admin.google.com/ac/apps/gmail/manageaddresslist)) have **no
management API** — Google confirms this is a known gap
([Issue Tracker #273613691](https://issuetracker.google.com/issues/273613691)). This script can
therefore only *discover* candidate domains and mail them for a human to paste into the console's
**Bulk add addresses** dialog (which accepts a comma or space separated list of domains).

### How it works

1. A time based trigger runs `harvestDomains()` periodically
2. The script lists all Slack workspace users (`users.list`, requires a bot token with `users:read`
   and `users:read.email`) and extracts the domain of each user's profile email address
3. Bot/deleted users, internal domains, well known public/free mail providers (Gmail, Outlook, GMX, ...)
   and any explicitly excluded domains are filtered out
4. The remaining domains are compared against the set already reported in a previous run (tracked in
   Script Properties); if there are new ones, a plain text report mail is sent listing:
    * the newly discovered domains, ready to paste into "Bulk add addresses"
    * the full current set of external domains, for an occasional full re-sync
5. Reported domains are remembered, so the same domain is never reported twice — even if it wasn't
   actually added to the address list (over-inclusion in the *report* is fine, since a human reviews it
   before anything reaches the actual allowlist)

No message is sent if there are no new domains since the last run.

### Prerequisites

* A Slack app/bot token with the `users:read` and `users:read.email` scopes, installed in the workspace.
* A regular Workspace user account (for example `automation@giantswarm.io`) the report is sent from,
  reachable via a service account with domain-wide delegation for the scope `https://mail.google.com/`.

### Configuration

Configure via Script Properties (use `setProperties` or the Apps Script UI):

| Mandatory | Property Key                              | Value Example                                   |
|-----------|--------------------------------------------|--------------------------------------------------|
| **yes**   | `DomainHarvester.serviceAccountCredentials`| `{...}` (single line service account JSON)        |
| **yes**   | `DomainHarvester.slackBotToken`            | `xoxb-...`                                        |
| **yes**   | `DomainHarvester.mailbox`                  | `automation@giantswarm.io`                        |
| **yes**   | `DomainHarvester.reportTo`                 | `it@giantswarm.io,security@giantswarm.io`         |
| **yes**   | `DomainHarvester.internalDomains`          | `giantswarm.io,giantswarm.com`                    |
| no        | `DomainHarvester.fromName`                 | `Domain Harvester`                                |
| no        | `DomainHarvester.excludedDomains`          | `mailinator.com,example-vendor.com`               |
| no        | `DomainHarvester.addressListName`          | `urgent-group-allowed-sender-domains` (default)   |
| no        | `DomainHarvester.adminConsoleUrl`          | admin console address list page (has a default)   |
| no        | `DomainHarvester.dryRun`                   | `false`                                           |

### Deployment

Follow the general deployment steps in the [root README](../README.md), then:

```sh
# create the Apps Script project (once)
cd domain-harvester/
clasp create --title "domain-harvester" --type standalone
cd ..

# assemble the library and push
SCRIPT_ID={SCRIPT_ID} make domain-harvester/

# configure
clasp run 'setProperties' --params '[{"DomainHarvester.slackBotToken": "xoxb-...", "DomainHarvester.mailbox": "automation@giantswarm.io", "DomainHarvester.reportTo": "it@giantswarm.io", "DomainHarvester.internalDomains": "giantswarm.io,giantswarm.com"}, false]'
clasp run 'setProperties' --params "[{\"DomainHarvester.serviceAccountCredentials\": $(cat credentials.json | tr -d '\n ' | jq -Rs .)}, false]"
```

### Testing before going live

```sh
# logs the report that would be sent, without sending mail or updating state
clasp run 'previewHarvestDomains'
```

Once satisfied, install the trigger:

```sh
clasp run 'install' --params '[30]'   # run every 30 minutes
clasp run 'uninstall'                 # remove the trigger again
```

If a domain was deliberately excluded from the address list but should be reconsidered (or a report
was missed), forget it so it's included in the next report again:

```sh
clasp run 'forgetDomains' --params '[["example.com"]]'
```
