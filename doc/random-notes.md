# Random notes

## External users in Azure Active Directory vs. SharePoint site (SPODS)

On rare occassions the Accepted By mail isn't unique within a site collection. Turns out that 
accounts of the form "i:0#.f|membership|live.com#foo@bar.com" may exist from a time when external 
users were not registered in Azure Active Directory. They would exist only in the SharePoint side 
of Office 365 (internally called SPODS). 

Later on, accounts of the form Account "i:0#.f|membership|foo@bar.com#ext#@mytenant.onmicrosoft.com" 
may have been added to the same tenant as an external user which would then appear in Azure Active
Directory. The SharePoint code that adds new users doesn't check for an existing SPODS user.

Having two accounts with the same Accepted By mail is a situation that only exists in these specific 
situations, and not something that can happen for new external user invites.

One has to manually remove one of the users as clearly one cannot login as different users with a 
single Live account. Remove is possible through: 
ctx.Site.RootWeb.SiteUsers.RemoveByLoginName("i:0#.f|membership|live.com#foo@bar.com");

## InvitedAs != AcceptedAs

SharePoint allows a user to accept the invitation from a different email address than the one the
user was invited on. In that case, SharePoint records InvitedAs != AcceptedAs.