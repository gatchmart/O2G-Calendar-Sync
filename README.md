# Outlook-Google-Calendar-Sync
An Outlook addin to sync google calendars with outlook calendars.

Thank you for looking into using this add-in and helping me troubleshoot it. In order to use the add-in you'll need Visual Studio 2017 to
build it and a Google API developer account. Both are free and and easier to get access to. 

The initial setup of the plug is pretty straight forward. You'll need to get a Google API Credentials, please follow the directions for
Step 1 here (https://developers.google.com/google-apps/calendar/quickstart/dotnet). Once you have the credentials you need to add them to
the project resources as a file named client_secret. To do this ensure the name of the file is client_secrets.json, right click on the
project name and select Properties. From there go to Resources, click the little down arrow to the right of the "Add Resource" button and
select "Add Existing File...". Select the client_secret.json file and click okay. Now the project should work properly.

If you have any issues running the plugin please create an issue for me to troubleshoot and fix. Make sure you provide details of how to
recreate the issue. If you have the call stack paste a copy of it in the issue as well.
