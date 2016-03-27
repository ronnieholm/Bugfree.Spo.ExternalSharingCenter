# TODO

- Switch to PnP Widget Wrangler
- Have backend store timestamp of last run in property bag for JS to display on frontpage
- Add filtering dropdown on top over overview table for better working with large table
- Make unit tests more robust
- Display loader icons on overview and guide to prevent flickering controls
- Bug in SPO tenant API sometimes causes number of external users in site collection to be off
- When deleting external user, make sure he's really gone. Deleting from Site Users isn't enough (hidden list? So not appear in import)
- Add CleanUp command to remove sharings which no longer exist/or has never existing in site collections.
  Currently, the backend throws a "Sequence contains no matching element" exception in GenerateExpirationWarnings.cs:20.
- How to restore Nuget packages from command-line?