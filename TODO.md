# TODO

- Switch frontend to SharePoint Framework
- Have backend store timestamp of last run in property bag / config list for JS to display on frontpage
- Add filtering dropdowns on top of overview table to enable better working with a large table
- Make unit tests more robust (don't rely on null guid which isn't unique)
- Display loader icons on overview and guide to prevent flickering of controls
- Bug in SPO tenant API sometimes causes number of external users in site collection to be off
- When deleting external user, make sure user is really gone. Deleting from Site Users isn't enough (hidden list? Azure AD? So user doesn't appear in import)
- How to restore Nuget packages from command-line?
- Add GuidId column to Sentmail
- Investigate how the REST API for external sharing works
- Build-in telemetry into the frontend code determine usage frequency and patterns
- Create config list / config.json file and deploy to document library to hold app wise settings
- If site collection with sharings no longer exist, adjust End date of sharings such that users don't receive warning and expiration mails for non-existing site collection in the future
- Investigate switching to JSLink for better use of existing UI