# TODO

- Switch to PnP Widget Wrangler (or custom implement relevant parts)
- Have backend store timestamp of last run in property bag / config list for JS to display on frontpage
- Add filtering dropdowns on top of overview table to enable better working with a large table
- Make unit tests more robust (don't rely on null guid which isn't unique)
- Display loader icons on overview and guide to prevent flickering of controls
- Bug in SPO tenant API sometimes causes number of external users in site collection to be off
- When deleting external user, make sure user is really gone. Deleting from Site Users isn't enough (hidden list? Azure AD? So user doesn't appear in import)
- Add CleanUp command to remove sharings which no longer exist/or has never existing in site collections
  Currently, the backend throws a "Sequence contains no matching element" exception in GenerateExpirationWarnings.cs:20.
  The exception is caused by InvitedAs being different from AcceptedAs. The Expriration command should've removed those
  but the feature may have been overlooked.
- How to restore Nuget packages from command-line?
- Add GuidId column to Sentmail
- Investigate how the REST API for external sharing works
- Build-in telemetry into the frontend code determine usage frequency and patterns
- Create config list / config.json file and deploy to document library to hold app wise settings