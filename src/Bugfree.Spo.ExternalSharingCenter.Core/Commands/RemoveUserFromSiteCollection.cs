using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Commands
{
    public class RemoveUserFromSiteCollection : Command
    {
        public RemoveUserFromSiteCollection(ILogger l) : base(l) { }

        public void Execute(ClientContext ctx, int userId)
        {
            Logger.Verbose($"About to execute {nameof(RemoveUserFromSiteCollection)} on '{ctx.Url}' for userId '{userId}");
            ctx.Web.SiteUsers.RemoveById(userId);
            ctx.ExecuteQuery();
        }
    }
}
