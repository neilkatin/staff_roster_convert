Staff Roster Convert
====================


Overview
--------

This application converts the automatic staffing reports into a single xlsx spreadsheet.

It takes all the sheets and add tables and freezes the titles.

It also adds extra sheets for the roster:

* 'Active' which does not have checked out responders
* one sheet for all the major groups (MC, LOG, etc)
* a sheet for responders with two days left
* a sheet for responders with 'negative' days off
* a sheet for responders without supervisors
* a sheet for all the other periodic reports
  * air travel roster
  * upcoming shifts
  * arrival roster
  * open staff requests

If you have any complaints or suggestions for improvements: send them to
[dr-report-automation@redcross.org](mailto:dr-report-automation@redcross.org)

One Time Setup
---------------

This repo needs a .env file to hold the security credentials for the project.  It should have the following contents:

* CLIENT_ID - the azure client id that is authorized to access the MS365 environment microsoft graph API
* CLIENT_SECRET - the azure client secret for this application

These credentials needs to be approved by Red Cross IT.

The first time the application runs you need to paste a magic URL into your browser.
The browser will redirect to a new URL that you need to copy/paste into the application.
You only have a few seconds after the redirect to copy the URL.

The app will save its local credentials in the config.TOKEN_FILENAME file, and use the contents of that file to periodically refresh its OAuth credentials.
