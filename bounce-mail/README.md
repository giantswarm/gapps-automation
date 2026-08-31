## Bounce Mail

Auto-replies to job applications that employees **forward** to a shared Google Group (for example
`jobs@giantswarm.io`).

Mail sent to the group directly by external senders is already answered by the group's own
auto-responder. That responder does not fire for mail forwarded by group members (employees), which
is the gap this script closes:
see [giantswarm/giantswarm#36069](https://github.com/giantswarm/giantswarm/issues/36069).

### How it works

1. A time based trigger runs `bounceMail()` every few minutes
2. The script impersonates the configured **mailbox** (a regular Workspace user account that receives
   the group's messages) and searches it for messages that are addressed to the group and not yet
   labelled as handled
3. For each message the *original* sender and subject are extracted from the forwarded content:
    * from the attached message headers, if the mail was forwarded as attachment (`message/rfc822`)
    * else from the forwarded header block in the body (Gmail, Outlook and Thunderbird flavors,
      several languages)
    * else from the `X-Forwarded-For` header, if the mail was forwarded automatically by a Gmail filter
4. Address and subject are **sanitized** (control characters removed, address validated, length limited)
   before they are used in the outgoing message
5. The configured auto-reply is sent to the original sender and the forwarded message is labelled as
   handled, so it is never answered twice

Messages that are not forwards, that were sent by an external address, that carry auto-response
headers or whose original sender is internal, a role account (`noreply@`, `postmaster@`, ...) or the
group itself are skipped and only labelled. All decisions are written to the execution log.

### Prerequisites

The Gmail API cannot read a Google Group, so a regular Workspace user account must receive the
group's messages. Either:

* add the account (for example `automation@giantswarm.io`) as a **member** of the group, or
* assign the group address to the account as an **alias**

To let the auto-reply come from the group address instead of the bot account, add the group address
as a verified *send as* alias of that account (Gmail settings → Accounts) and configure it as
`BounceMail.fromAddress`. Without a verified alias, Gmail rewrites the `From` header.

The service account needs domain-wide delegation for the scope `https://mail.google.com/`.

### Configuration

Configure via Script Properties (use `setProperties` or the Apps Script UI):

| Mandatory | Property Key                          | Value Example                                  |
|-----------|---------------------------------------|------------------------------------------------|
| **yes**   | `BounceMail.serviceAccountCredentials`| `{...}` (single line service account JSON)      |
| **yes**   | `BounceMail.mailbox`                  | `automation@giantswarm.io`                      |
| **yes**   | `BounceMail.internalDomains`          | `giantswarm.io,giantswarm.com`                  |
| no        | `BounceMail.groupEmail`               | `jobs@giantswarm.io`                            |
| no        | `BounceMail.fromAddress`              | `jobs@giantswarm.io` (defaults to the mailbox)  |
| no        | `BounceMail.fromName`                 | `Giant Swarm`                                   |
| no        | `BounceMail.replyTo`                  | `jobs@giantswarm.io`                            |
| no        | `BounceMail.replySubject`             | `Re: ${subject}`                                |
| no        | `BounceMail.replyText`                | the auto-reply body (built-in default wording)  |
| no        | `BounceMail.searchExpression`         | `newer_than:2d -in:chats -in:trash -in:spam`    |
| no        | `BounceMail.handledLabel`             | `BounceMail/handled`                            |
| no        | `BounceMail.requireInternalForwarder` | `true`                                          |
| no        | `BounceMail.maxRepliesPerRun`         | `25`                                            |
| no        | `BounceMail.replyCooldownHours`       | `24`                                            |
| no        | `BounceMail.dryRun`                   | `false`                                         |

`BounceMail.replySubject` and `BounceMail.replyText` support the placeholders `${senderName}`,
`${senderEmail}` and `${subject}`, substituted with the sanitized values of the original message.

`BounceMail.replyCooldownHours` keeps the same applicant from being auto-replied to repeatedly (set
to `0` to reply to every single forward). `BounceMail.maxRepliesPerRun` bounds the damage should a
mail loop ever occur despite the loop protection.

### Deployment

Follow the general deployment steps in the [root README](../README.md), then:

```sh
# create the Apps Script project (once)
cd bounce-mail/
clasp create --title "bounce-mail" --type standalone
cd ..

# assemble the library and push
SCRIPT_ID={SCRIPT_ID} make bounce-mail/

# configure
clasp run 'setProperties' --params '[{"BounceMail.mailbox": "automation@giantswarm.io", "BounceMail.groupEmail": "jobs@giantswarm.io", "BounceMail.internalDomains": "giantswarm.io,giantswarm.com", "BounceMail.fromAddress": "jobs@giantswarm.io", "BounceMail.fromName": "Giant Swarm"}, false]'
clasp run 'setProperties' --params "[{\"BounceMail.serviceAccountCredentials\": $(cat credentials.json | tr -d '\n ' | jq -Rs .)}, false]"
```

### Testing before going live

Forward a test application to the group from an employee account, then run:

```sh
# logs the auto-reply that would be sent, without sending anything or labelling messages
clasp run 'previewBounceMail'
```

Once the extracted sender and subject look right, install the trigger:

```sh
clasp run 'install' --params '[5]'   # run every 5 minutes
clasp run 'uninstall'                # remove the trigger again
```
