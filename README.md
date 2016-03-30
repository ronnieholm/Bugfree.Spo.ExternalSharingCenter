# Bugfree.Spo.ExternalSharingCenter

With SharePoint Online, getting a tenant-wide overview of which site
collections are shared with which external users (anyone with a
Microsoft account) is tricky. Without such overview, and governance in
place for managing external users, one risks inadvertently sharing
information with too many for too long.

The External Sharing Center (from here on ESC) adds a layer of
management and governance on top of the out of the box features. It
provides the overview and support for creating and maintaining
sharings with a set start and end date. As the end date nears, ESC
reminds the inviter that a sharing is about to expire (with the option
to extend it). Once expired, ESC revokes user access from the site
collection and sends an access revocation mail to the inviter.

ESC consists of a frontend and a backend part:

  - A frontend SharePoint web to be installed into any site
    collection. The web holds pages to display an overview of sharings
    by site collection and for easy adding and editing of sharing
    using a guided approach (see end-user documentation).

  - A background job which monitors site collections and sends out
    expiration warnings, revokes user access, and sends out access
    revocation mails. The job can be setup to continuously import new
    users and sharings. Otherwise, it threats sharings not already in
    ESC as unknown and removes those. The approach to use depends on
    whether sharings recorded in ESC or site collections are
    considered master.

Other notable features of ESC include:

  - Aggregated expiration warning and access revocation mails grouped
    by inviter to minimize number of mails sent. Each mail contains a
    table of sharings and provides a link to ESC for editing.

## Compiling

The repository is self-contained and may be compiled by opening the
solution in Visual Studio 2015. Compiling triggers NuGet package
restore and transpiles the TypeScript which powers the frontend.

## Installing

### Backend

The CLI is used for partial one-time setup of the frontend, importing
existing sharings, expirering user access, and generating and sending
mails. The WebJob serves a similar purpose running in Azure, but only
generating mails and expires user access.

Before running the CLI (or WebJob), AppSettings inside
Bugfree.Spo.ExternalSharingCenter.Cli/App.config must be defined as
per below description. Running the CLI in different modes is
accomplished by editing Program.cs file:

  - **ExternalSharingCenterUrl**. Running CLI in setup mode assumes
    the hosting web exists and is based on an English (1033) team site
    template. CLI sets up pages, lists, quick launch, and so forth. In
    regular mode, CLI processes users and sharings stored in lists on
    the web.

  - **TenantAdministratorUsername**, **TenantAdministratorPassword**.
    Iterating site collections and their external users requires
    tenant-level access using this account and password.

  - **ExpirationWarningDays**. When a sharing expires in this number
    of days, the backend sends a warning mail to the inviter or one in
    the Owners groups of the site collection in case the inviter
    cannot be located.

  - **ExpirationWarningMailsMinimumDaysBetween**. To avoid repeatedly
    sending expiration mails whenever the backend runs, this value
    indicates the minimum number of days between sending warning mails
    for external users in a single site collection.

  - **MailFrom**. Mail is sent using SharePoint's Utility.SendEmail
    routine, i.e., no SMTP server setup is required. SendEmail
    requires a site collection context and thus only supports mailing
    users in a specific site collection. The from address must be a
    valid mail address as verified by SharePoint Online.

For the WebJob, two connection strings are required. These are
standard WebJob settings and not related to ESC, per se:

  - **AzureWebJobsDashboard**, **AzureWebJobsStorage**. The WebJob API
    stores logs and other runtime information inside an Azure Storage
    Account. Ideally, the account would be named after the WebJob, but
    the name must be between 3 and 24 lower-case characters. Thus,
    bugfreespoextsharingcntr would be a decent candidate (this is the
    placeholder used inside the config files).

    The connection string is available under Storage Account, Keys,
    and then Primary Connection String. As with any config file
    connection string, Azure ignores it for security reasons and
    requires those be setup in the Azure Portal under Application
    Settings.

On first-time publication of Bugfree.Spo.ExternalSharingCenter.WebJob
from inside Visual Studio, a schedule must be setup. Running the
WebJob on a recurring schedule at 1am, every night is a sensible
default given that user access expires at midnight.

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

Depending on the web's security settings, Contributor access may need
to be granted to users creating and editing sharings (JavaScript runs
on the user's behalf). Specifically, External users and Site
collection external users lists must be editable.

## Limitations

Only sharing an entire site collection by authenticated users is
supported by ESC. Sharing at a finer level of granularity, such as
list or list item, makes gaining an overview even trickier. Instead of
these finer-grained sharing options, consider creating a dedicated
sharing site collection besides the regular ones.

## Design notes

Within ESC, we try hard not to replicate build-in SharePoint dialogs
for external sharing. Since we cannot easily hook into SharePoint's
invite logic, instead we provide import logic for either one-time or
continuous synchronization of SharePoint actual with ESC. In either
case, users and sharings are added to ESC with configurable default
values, such as some number of days until the sharing expires.

SharePoint Online provides an API (through the
[ShareObject](https://blogs.msdn.microsoft.com/vesku/2015/10/02/external-sharing-api-for-sharepoint-and-onedrive-for-business)
class in CSOM and similar in REST) by which invitations can be
programmatically created. In principle, it would be possible for ESC's
JavaScript to call this REST API on the user's behalf. This would
relieve the user from having to create a matching sharing in both the
site collection and ESC.

Given the time it takes for the backend to iterate site collections
and their sharing, the frontend can't provide up-to-date sharing
information except for what it already stores. That's why the ESC
frontend doesn't show display name of user, invited as, invited by,
and when the sharing is created. Such information would have to be
stored inside ESC (in a list or property bag entry) and maintained by
the backend.

## Additional resources

[The Definitive Guide to Office 365 External Sharing](http://en.share-gate.com/blog/ultimate-guide-deal-with-office-365-external-sharing)

[Manage external sharing for your SharePoint Online environment](https://support.office.com/en-us/article/Manage-external-sharing-for-your-SharePoint-Online-environment-c8a462eb-0723-4b0b-8d0a-70feafe4be85)

[Controlling tenant and site collection setting with CSOM](https://github.com/OfficeDev/PnP/tree/master/Samples/Core.ExternalSharing#controlling-tenant-and-site-collection-setting-with-csom)

[Understanding External Users in SharePoint Online](https://blogs.technet.microsoft.com/lystavlen/2013/04/14/understanding-external-users-in-sharepoint-online)

[Share sites or documents with people outside your organization](https://support.office.com/en-us/article/Share-sites-or-documents-with-people-outside-your-organization-80e49744-e30f-44db-8d51-16661b1d4232?CTT=5&origin=HA104036809&CorrelationId=df67710b-baa6-466d-958e-783cc8f0f8fd&ui=en-US&rs=en-US&ad=US)

[Sign up for a Microsoft account](http://windows.microsoft.com/en-us/windows-10/sign-up-for-a-microsoft-account)

## Supported platforms

SharePoint Online.