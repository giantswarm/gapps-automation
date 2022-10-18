
# Apps Script Automation Project

Apps Script based automation uses Google's `clasp` tool for automating the build/release/deploy workflow.

Extensive documentation here: https://developers.google.com/apps-script/guides/clasp

## Deployment


 1. Enable Google Apps Script API
    https://script.google.com/u/1/home/usersettings
 2. Login using clasp manually or via credentials:
    ```
    clasp login [--creds <file>]
    ```
    The login info is stored in `~/.clasprc.json`.
 3. Change into the relevant project sub directory
 4. Upload local files to drive:
    ```
    clasp push
    ```
 5. Create a new version of the project and output the version number:
    ```
    clasp version MY_VERSION_DESCRIPTION
    ```
 6. Deploy this new version:
    ```
    clasp deploy VERSION_NUMBER MY_DEPLOYMENT_DESCRIPTION
    ```

## Usage

Apps Scripts projects have differing forms of deployments.

Many are deployed as API Endpoint to be called via REST API or triggers.

Other possibilities include Sheets Macros and GApps Addons.

