
# Apps Script Automation Project

Apps Script based automation uses Google's `clasp` tool for automating the build/release/deploy workflow.

Extensive documentation here: https://developers.google.com/apps-script/guides/clasp


## Dependencies

 * Advanced Sheets Service
   * Can be enabled in Google Cloud Console or via `clasp apis enable sheets`
 * Personio API v1

## Deployment

 1. Enable Google Apps Script API
    https://script.google.com/u/1/home/usersettings
 2. Login using clasp manually or via credentials:
    ```sh
    clasp login [--creds <file>]
    ```
    The login info is stored in `~/.clasprc.json`.
 3. Change into the relevant project sub directory
 4. Upload local files to drive:
    ```sh
    clasp push
    ```
 5. Create a new version of the project and output the version number:
    ```sh
    clasp version MY_VERSION_DESCRIPTION
    ```
 6. Deploy this new version:
    ```sh
    clasp deploy VERSION_NUMBER MY_DEPLOYMENT_DESCRIPTION
    ```
 7. Configure properties for the scripts using the builtin helper function:
    ```sh
    clasp run 'setProperties' --params '[{"KEY": "VALUE"}, true]'
    ```

## Usage

Apps Scripts projects have differing forms of deployments.

Many are deployed as API Endpoint to be called via REST API or triggers.

Other possibilities include Sheets Macros and GApps Addons.

### Default GCP Project for Execution

* Name: `gapps-automation`
* ID: `gapps-automation-2022`
* Number: `216505013251`

### GCloud

Install `gcloud` to manage execution, settings and logging of the automation:

https://cloud.google.com/sdk/docs/cheatsheet
