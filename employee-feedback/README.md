## Employee Feedback

Generates personalised feedback forms on demand. When a coordinator submits the intake Google Form, the script fetches role-specific question templates from GitHub, merges them, creates a new Google Form, and emails the edit/share links to the respondent.

### How it works

1. Coordinator fills in the **intake form** (Name, Roles, optional Custom Questions)
2. `onFormSubmit` fires and fetches Markdown templates from `giantswarm/feedback-templates` for each selected role
3. Templates are merged (deduplicating questions across roles)
4. A new Google Form is created in `automation@giantswarm.io`'s Drive
5. An email is sent to the coordinator with the edit URL and the shareable URL

### Configuration

Configure via Script Properties (use `setProperties` or the Apps Script UI):

| Mandatory | Property Key                  | Value Example                                       |
|-----------|-------------------------------|-----------------------------------------------------|
| **yes**   | `EmployeeFeedback.formId`     | `1FAIpQLSe...` (ID from the intake form URL)        |
| no        | `EmployeeFeedback.folderId`   | `1a2B3c...` (Drive folder ID to store output forms) |

The form ID is the long string in the Google Form URL:
`https://docs.google.com/forms/d/**<FORM_ID>**/edit`

### Deployment

Follow the general deployment steps in the [root README](../README.md), then:

#### 1. Create the Apps Script project as `automation@giantswarm.io`

```sh
# Log in to clasp as automation@giantswarm.io
clasp login

# From the employee-feedback directory, create a new standalone project
cd employee-feedback/
clasp create --title "employee-feedback" --type standalone
cd ..
```

This creates a `.clasp.json` (gitignored) containing the new Script ID.

#### 2. Push the code

```sh
make employee-feedback/
```

#### 3. Set script properties

```sh
cd employee-feedback/
clasp run 'setProperties' --params '[{"EmployeeFeedback.formId": "YOUR_INTAKE_FORM_ID"}, false]'
```

Or set them manually in the Apps Script editor: **Project Settings → Script Properties**.

#### 4. Give `automation@giantswarm.io` access to the intake form

The account running the trigger must have at least **Editor** access to the intake form so it can read responses.

Share the intake form with `automation@giantswarm.io` as an Editor.

#### 5. Install the trigger

```sh
cd employee-feedback/
clasp run 'install'
```

This registers an `onFormSubmit` installable trigger on the intake form, running as `automation@giantswarm.io`.

#### 6. Remove the old trigger from your personal account

In your personal Google Account:
- Open the intake form → **⋮ → Script editor**
- Delete the container-bound script, **or**
- Go to **Triggers** and delete the `onFormSubmit` trigger

Alternatively, remove it via the [Apps Script dashboard](https://script.google.com/home).

### Dependencies

- GitHub repository `giantswarm/feedback-templates` must be public and contain Markdown role templates at the root level (e.g. `generic.md`, `engineer.md`, `manager.md`)
- The intake form must collect **Respondent Email** (Form settings → Collect email addresses)
- The intake form must have fields named exactly: **Name**, **Roles**, **Custom Questions**
