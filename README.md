
# Apps Script Automation Projects

Apps Script based automation, managed via Google's `clasp` tool (build/release/deploy).

Extensive documentation here: https://developers.google.com/apps-script/guides/clasp

## Dependencies

 * Advanced Sheets Service
   * Can be enabled in Google Cloud Console or via `clasp apis enable sheets`
 * Personio API v1

## Deployment

 1. Enable Google Apps Script API
    https://script.google.com/u/1/home/usersettings
 2. Login globally using clasp
    ```sh
    clasp login
    ```
    The login info is stored in `~/.clasprc.json`.
 3. Change into the relevant project sub directory
    ```sh
    cd personio-to-sheets
    ```
 4. Login "locally" inside the script project directory:
    ```sh
    clasp login --creds <OAUTH2_GOOGLE_CLOUD_PROJECT_CLIENT_SECRET_FILE>
    ```
 5. Upload local files to drive:
    ```sh
    clasp push
    ```
 6. Deploy this new version:
    ```sh
    clasp deploy
    ```
 7. Configure properties for the scripts using the builtin helper function:
    ```sh
    clasp run 'setProperties' --params '[{"KEY": "VALUE"}, true]'
    ```
 8. Install automated script functions:
    ```sh
    # install a script (argument "5" could be the delay in minutes)
    clasp run 'install' --params '[5]'  
    ```
 9. Uninstall automated script function:
    ```sh
    clasp run 'uninstall' # remove trigger(s) for the relevant project
    ```

## Usage

Apps Scripts projects have differing forms of deployments.

Many are deployed as API Endpoint to be called via REST API or triggers.

Other possibilities include Sheets Macros and GApps Addons.

### GCloud

Install `gcloud` (CLI) to manage execution, settings and logging of the automation:

https://cloud.google.com/sdk/docs/cheatsheet

### Default Giant Swarm GCP Project for Execution

* Name: `gapps-automation`
* ID: `gapps-automation-2022`
* Number: `216505013251`
