using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using C = Bugfree.Spo.ExternalSharingCenter.Core.Commands.CreateExternalSharingCenterWeb.SiteCollectionExternalUserColumns;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Commands
{
    public class AddSiteCollectionExternalUser : Command
    {
        public AddSiteCollectionExternalUser(ILogger l) : base(l) { }

        public void Execute(List siteCollectionExternalUsers, SiteCollectionExternalUser sceu)
        {
            Logger.Verbose($"About to execute {nameof(AddSiteCollectionExternalUser)} for site collection Url '{sceu.SiteCollectionUrl}' and external user id '{sceu.ExternalUserId}'");
            var i = siteCollectionExternalUsers.AddItem(new ListItemCreationInformation());
            i[C.SiteCollectionExternalUserId] = sceu.SiteCollectionExternalUserId;
            i[C.SiteCollectionUrl] = sceu.SiteCollectionUrl;
            i[C.ExternalUserId] = sceu.ExternalUserId;
            i[C.Start] = sceu.Start;
            i[C.End] = sceu.End;
            i[C.Comment] = sceu.Comment;
            i.Update();
            siteCollectionExternalUsers.Context.ExecuteQuery();
        }
    }
}
