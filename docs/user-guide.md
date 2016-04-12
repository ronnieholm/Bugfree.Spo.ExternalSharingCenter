# External Sharing Center user guide

This user guide outlines the three stages involved in managing
sharings with the External Sharing Center (ESC). First, from within
the SharePoint admin portal, a tenant administrator must enable
sharing of the site collection. Then, within the site collection, an
existing user invites a Microsoft account, adding it to the
appropriate security group. Lastly, an existing user registers the
Microsoft account with ESC, providing start and end dates for the
access (and possibly other metadata).

## Enable site collection sharing

Enabling external sharing at the tenant level has already been
[thoroughly
documented](http://en.share-gate.com/blog/ultimate-guide-deal-with-office-365-external-sharing#configureExternalSharing). Suffice
it to say that ESC supports only sharing entire collections with
authenticated external users. Within a tenant, it's still possible to
share at the list or item level, but ESC will ignore those site
collections.

## Managing external users within a site collection

Sharing a site collection with an external user has already been
[thoroughly
documented](http://en.share-gate.com/blog/ultimate-guide-deal-with-office-365-external-sharing#shareContent). One
caviat though is that when an external user accepts an invitation, the
user must be logged in with the correct Windows account. To SharePoint
Online, the correct Windows account is any valid account, but a user
may have multiple accounts -- work, private, Office 365 accounts which
are by default Microsoft accounts -- and using the right one is
essential to logging in and receiving mails from SharePoint Online.

In case the address to which the invitation is sent isn't a Microsoft
account, the recipient is free to login with any other Microsoft
account or register the mail address as a Microsoft account. Clicking
the link in the invitation mail provides the recipient with the
necessary information.

## Managing external users within the External Sharing Center

The start page of ESC is shown below. Out of the box, the design is
kept simple, allowing ESC administrators to customize it for local
needs.

![alt text](https://raw.githubusercontent.com/ronnieholm/Bugfree.Spo.ExternalSharingCenter/master/docs/external-sharing-center-overview.png)

The Quick launch on the left links to two page: *Start* is the front
page and *External user guide* is a step-wise guide to create a new
sharing or modify an existing one. Adding or editing a sharing is done
either by clicking *External user guide* or clicking the *Add* or
*Edit* links next to site collections in the table, displaying a
summary of sharings for site collections to which the user has at
least read-access.

Clicking *External user guide*, the guide starts at step one (see
details of steps below), whereas the Add or Edit links skip ahead to
step 3 and 4, respectively, with information about site collection and
user filled in based on the row selected.

The guide consists of a total of five steps (three if we ignore
introduction and summary):

  - Introduction

    The text should be self-explanatory.

    ![alt text](https://raw.githubusercontent.com/ronnieholm/Bugfree.Spo.ExternalSharingCenter/master/docs/site-collection-external-user-guide-step-1-of-5.png)

  - Site collection selection

    The dropdown contains an alphabetically sorted list of titles of
    site collection to which the user has at least read access. It's
    the same list displayed on the start page.

    ![alt text](https://raw.githubusercontent.com/ronnieholm/Bugfree.Spo.ExternalSharingCenter/master/docs/site-collection-external-user-guide-step-2-of-5.png)

  - User selection

    An existing mail address may be selected from the dropdown,
    listing all previously added external users in ESC. If the user
    doesn't already exist, the new mail address must be entered in the
    text field. Once the sharing is created, the new mail address is
    added to ESC and included in the dropdown for future sharings.

    ![alt text](https://raw.githubusercontent.com/ronnieholm/Bugfree.Spo.ExternalSharingCenter/master/docs/site-collection-external-user-guide-step-3-of-5.png)

    Behind the scenes, SharePoint Online doesn't store external users
    inside each site collection. Instead, users are added to Azure
    Active Directory and from there users are shared between site
    collections much the same way organizational users are. Similarly,
    within ESC the user is created only once and user specific
    metadata, such as the mail address, is shared between sharings.

  - Start and end date selection
  
    Clicking inside the date fields causes a date selector to appear,
    alleviating the need to enter dates by hand. The start and end
    date (both inclusive) concern the sharing within the selected site
    collection and not the user. The user may have external access to
    multiple site collection, each with a different start and end
    date.

    ![alt text](https://raw.githubusercontent.com/ronnieholm/Bugfree.Spo.ExternalSharingCenter/master/docs/site-collection-external-user-guide-step-4-of-5.png)

  - Summary

    Once satisfied with the
    sharing, clicking the *Save* button creates the sharing (and
    possibly the user) within ESC. After being redirect to the start page, 
    the new sharing becomes visible.

    ![alt text](https://raw.githubusercontent.com/ronnieholm/Bugfree.Spo.ExternalSharingCenter/master/docs/site-collection-external-user-guide-step-5-of-5.png)

## Mail notifications

Based on the provided end date, ESC reminds inviters when a sharing is
about to end and when it has ended. As a rule of thumb, mails are sent
to whoever invited the user. In case that user no longer exist, insted
one of the users from the site collection's Owner's group becomes the
recipient of the mail.

Below is an example of what mails look like. The wording may differ
based on whether it's a reminder or a user access expiration mail, but
the gist of it remains the same. It consists of a grouping of
notifications for the inviter. The mail may contain multiple entries
for a single site collection in case multiple users require attention
or it may contain multiple site collections if the users are spread
out.

![alt text](https://raw.githubusercontent.com/ronnieholm/Bugfree.Spo.ExternalSharingCenter/master/docs/external-user-mail-template.png)

Because mails are static by nature, they don't contain direct links to
the External user guide. Actual information may have changed since the
mail was dispatched, causing links to become stale.