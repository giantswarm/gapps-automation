![Library Test](https://github.com/giantswarm/gapps-automation/actions/workflows/build_and_test.yaml/badge.svg)

# Apps Script Automation Projects

Apps Script based automation for a company that uses GApps with Personio.

Uses Google's `clasp` tool for deployment.


## Dependencies

### Local Dependencies

* `node`  
    The `node` binary allows running JavaScript locally (used to run `clasp` and for tests).  
    **Install:** `sudo apt install nodejs`.  
    **Documentation:** https://nodejs.org/en

* `clasp`  
    Automates Apps Script deployment tasks.  
    **Install:** `sudo npm install -g clasp`.  
    **Documentation:** https://developers.google.com/apps-script/guides/clasp  

* `clasp-env`  
    Links subdirectories to clasp projects.  
    **Install:** `sudo npm install -g clasp-env`.  
    **Documentation:** https://medium.com/geekculture/if-you-use-clasp-with-google-apps-script-you-need-this-utility-right-now-de61fd4e67c8  


### Cloud Services for Deployment

* Advanced Sheets Service  
  Enable via Google Cloud Console or `clasp apis enable sheets`
* Google Calendar API  
  Enable via Google Cloud Console or `clasp apis enable calendar`
* Personio API v1  
  Create an API token with the necessary fields on Personio


## Deployment

1. Enable Google Apps Script API  
   https://script.google.com/u/1/home/usersettings  
   * Some scripts need additional configuration (service accounts, scopes, ...). Refer to the function comments for that.
2. Login globally using clasp  
   ```sh
   clasp login
   ```
   The login info is stored in `~/.clasprc.json`.
3. Change into the relevant project sub-directory
   ```sh
   cd personio-to-sheets
   ```
4. Login "locally" inside the script project directory:
   ```sh
   clasp login --creds <OAUTH2_GOOGLE_CLOUD_PROJECT_CLIENT_SECRET_FILE>
   ```
5. Upload from working copy to a new or existing Apps Script project  
   The `personio-to-sheets` sub project is used as an example.
   * Link and push to existing Apps Script project:
      ```
      clasp list | sed 's/https\:\/\/script\.google\.com\/d\///g'
      SCRIPT_ID={SCRIPT_ID_COPIED_FROM_ABOVE} make personio-to-sheets/
      ```
   * Link and push to new project (with or without parent sheets/docs document):
      ```
      cd personio-to-sheets/
      clasp create --help  # see documentation for parameter '--type'
      clasp create --title "$(basename "$(pwd)")" --type standalone --parentId OPTIONAL_PARENT_DOCUMENT_ID
      cd ..
      SCRIPT_ID={SCRIPT_ID_COPIED_FROM_ABOVE} make personio-to-sheets/
      ```
6. Configure properties for the scripts using the builtin helper function:
   ```sh
   clasp run 'setProperties' --params '[{"KEY": "VALUE"}, false]'
   ```
7. Install automated script functions:
   ```sh
   # install a script (argument "5" could be the delay in minutes)
   clasp run 'install' --params '[5]'  
   ```
8. Uninstall automated script function:
    ```sh
    clasp run 'uninstall' # remove trigger(s) for the relevant project
    ```

### Using Makefile

A Makefile is provided to ease development and deployment on CI.

#### Examples

* Build node library in `lib-output` and run local tests:

  Running the tests requires the `node` binary to be installed. The installed node must implement `fetch()` (>= v16).

  ```make clean && make test```

* Assemble and push all projects (must be locally logged in, see above):

  ```make```

* Assemble and push the logged in project `options-sheets`:

  ```make options-sheets/```

* Clean all generated files (for example `$project/lib` or `lib-output`):

  ```make clean```

## Usage

Apps Scripts projects have differing forms of deployments.

Many are deployed as API Endpoint to be called via REST API or triggers.

Other possibilities include SheetUtil Macros and GApps Addons.

### GCloud

Optionally install `gcloud` (CLI) to manage execution, settings and logging of the automation:

https://cloud.google.com/sdk/docs/cheatsheet

### Library

According to Google docs Apps Script Libraries should be used sparingly to avoid decreasing performance.

We maintain a shared directory based library, called `lib`, which includes all our shared code for in-house App Script
work.

This directory is copied into each dependent project via `Makefile`.

### Hints


#### Attaching to a Sheets/Docs/Forms Container

The Google Sheets/Forms UI makes it hard to attach a script project when using multiple Google accounts.

To attach a project, get the document ID and use the following clasp command line in the project subdirectory:  
`clasp create --parentId {SHEET_ID} --rootDir .`

The logged in clasp user account must have `Editor` access to the document.

#### Building just the Lib

For testing or use in other projects (possibly targetting another runtime) building just the library part of this project may be required.

To build just `lib-output/lib.js` run the following:
```
# the directory lib-output will be automatically created if it doesn't exist
make lib
```
