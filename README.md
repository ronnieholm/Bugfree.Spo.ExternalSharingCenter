# Bugfree.Spo.ExternalSharingCenter

With SharePoint Online, getting a tenant-wide overview of which site
collections are shared with which external users (anyone with a
Microsoft account) is tricky. Without such overview, and governance in
place for managing external users, one risks inadvertently sharing
information with too many for too long.

The External Sharing Center (from here on simply ESC) adds a layer of
management and governance on top of out of the box features. It
provides the overview and supports creating and maintaining sharings
with a set start and end date. As the end date gets close, ESC reminds
the invited that a sharing is about to expire (end date can then be
adjusted). Once expired, ESC revokes user access from the site
collection and sends an access revocation mail to the invited.

ESC consists of a frontend and a backend part:

  - A frontend SharePoint web to be installed into a site collection
    of your choice. The web holds pages to display an overview of
    sharings by site collection and for easy adding and editing of
    sharing using a guided approach (see end-user documentation).

  - A background job which monitors site collections and sends out
    expiration warnings and access revocation mails once user is
    removed. The job may be setup to continuously import new users and
    sharings. Otherwise, the job threats sharings not already in ESC
    as unknown and removes those. Which approach to pick depends on
    whether sharings recorded on ESC or site collections are
    considered master.

Other notable features of ESC include:

  - Aggregated expiration warning and access revocation mails grouped
    by inviter to minimize the number of mails an inviter
    receives. Each mail contains a table of sharings and provides a
    link to ESC for editing.

## Compiling

The repository is self-contained and may be compiled by opening the
solution in Visual Studio 2015. Compiling triggers NuGet package
restore and transpiles TypeScript used in implementing the frontend.

## Installing

### Backend

The CLI is used for partial one-time setup of the frontend, importing
existing sharings, expire user access, and generating and sending
mails. The WebJob serves a similar purpose running in Azure, but only
generating mails and expire user access.

Before running the CLI or WebJob, AppSettings inside
Bugfree.Spo.ExternalSharingCenter.Cli/App.config must be defined as
per below documentation. Running CLI in different modes is
accomplished by editing Program.cs file:

  - **ExternalSharingCenterUrl**. Running CLI in setup mode assumes
    the hosting web exists and is based on an English (1033) team site
    template. CLI then sets up pages, lists, quick launch, and so
    forth.

  - **TenantAdministratorUsername**, **TenantAdministratorPassword**.
    Iterating site collections and their external users require
    tenant-level access using this account and password.

  - **ExpirationWarningDays**. When a sharing expires in this number
    of days, the backend sends a warning mail to the inviter or one in
    the Owners groups of the site collection in case the inviter
    cannot be located.

  - **ExpirationWarningMailsMinimumDaysBetween**. To avoid repeatedly
    sending expiration mails whenever the backend runs, this value
    indicates the minimum number of days that should elapse between
    sending expirations mails for external users in a single site
    collection.

  - **MailFrom**. Mail is send using SharePoint's Utility.SendEmail
    routine, i.e., no SMTP server setup is required. SendEmail
    requires a site collection context and thus only supports mailing
    users in that site collection. The from address must be a valid
    mail address as verified by SharePoint Online.

For the WebJob, two connection strings are required. These are
standard WebJob settings and doesn't have anything to do with ESC:

  - **AzureWebJobsDashboard**, **AzureWebJobsStorage**. The WebJob API
    stores logs and other runtime information inside an Azure Storage
    Account. Ideally, the account would be named after the WebJob, but
    the name must be between 3 and 24 lower-case characters. Thus,
    bugfreespoextsharingcntr is a decent candidate.

    The connection string is available under the Storage Account,
    Keys, and then Primary Connection String. As with any connection
    string, for security reasons Azure ignores the ones in the config
    file, and requires those to be setup in the Azure Portal under
    Application Settings.

On first-time publication of Bugfree.Spo.ExternalSharingCenter.WebJob
from inside Visual Studio, a schedule must be setup. Running the
WebJob on a recurring schedule at 1am, every night is a sensible
default given that regular access expires at midnight.

### Frontend

A few manual setup steps are required to finalize ESC setup:

  1. On Start.aspx and
     Site%20collection%20external%20user%20guide.aspx within
     SitePages, put the page in edit mode and edit the properties of
     the Loader web part. Under Appearance, change Chrome Type from
     Default to None and exit page editing.

  2. Within SiteAssets, delete the existing Notebook file.
  
  3. Within SiteAssets, create a folder named App and upload the
     following files from Bugfree.Spo.ExternalSharingCenter.Core/App
     to it:

        Apps.js
        Loader.js
        Loader.html
        SiteCollectionExternalUserGuide.html
        SiteCollectionExternalUsersOverview.html

Depending on web security group setup, Contributor access may need to
be explicitly granted for users (JavaScript runs on the user's behalf)
to add and edit items on the External users and Site collection
external users lists.

## Limitations

Only sharing an entire site collection by authenticated users is
supported. Sharing at a finer level of granularity, such as a list or
list item, makes an overview even trickier. Instead of these
fine-grained sharing options, consider creating a dedicated sharing
site collection besides the regular ones.

## Design notes

Within ESC, we try hard not to replicate build-in SharePoint dialogs
for external sharing. Since we can't easily hook into SharePoint's
invite logic, we provide import logic for either one-time or
continuous synchronization of SharePoint actual with ESC. In either
case, users and sharing are added to ESC with configurable default
values, such as some number of days until the sharing expires.

SharePoint Online does provide an API (through the ShareObject class)
for inviting external users. But calling it requires running on behalf
of the invited by user and thus requires a provider hosted app or app
part. Still, we don't want to replicate the out of the box user
interface in such an app.

Given the time it takes for the backend to iterate site collections
and sharing, the frontend can't provide up-to-date sharing information
except for what it already stores. That's why the ESC frontend
currently doesn't show display name of user, invited as, invited by,
and when created information. Such information would have to be stored
inside ESC (in a list of property bag entry) and maintained by the
backend.

## Additional resources

[The Definitive Guide to Office 365 External Sharing](http://en.share-gate.com/blog/ultimate-guide-deal-with-office-365-external-sharing)

[Manage external sharing for your SharePoint Online environment](https://support.office.com/en-us/article/Manage-external-sharing-for-your-SharePoint-Online-environment-c8a462eb-0723-4b0b-8d0a-70feafe4be85)

[Controlling tenant and site collection setting with CSOM](https://github.com/OfficeDev/PnP/tree/master/Samples/Core.ExternalSharing#controlling-tenant-and-site-collection-setting-with-csom)

## Supported platforms

SharePoint Online.