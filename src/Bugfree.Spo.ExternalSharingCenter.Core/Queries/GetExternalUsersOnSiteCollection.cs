using System.Linq;
using System.Collections.Generic;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GetExternalUsersOnSiteCollection : Query
    {
        public GetExternalUsersOnSiteCollection(ILogger l) : base(l) { }

        private List<SharePointExternalUser> GetExternalUsersRecursive(Site siteCollection, int startPosition) 
        {
            Logger.Verbose($"Fetching external users for '{siteCollection.Url}' starting at position {startPosition}");
            const int pageSize = 50;
            var ctx = siteCollection.Context;

            var o365Tenant = new Office365Tenant(ctx);
            var users = o365Tenant.GetExternalUsersForSite(siteCollection.Url, startPosition, pageSize, string.Empty, SortOrder.Ascending);
            users.Context.Load(users, e => e.ExternalUserCollection, e => e.TotalUserCount);
            users.Context.ExecuteQuery();

            ctx.Load(users, e => e.ExternalUserCollection);                    
            ctx.ExecuteQuery();

            var myExternalUsers = users.ExternalUserCollection.ToList().Select(u =>
                new SharePointExternalUser
                {
                    UserId = u.UserId,
                    DisplayName = u.DisplayName,
                    AcceptedAs = u.AcceptedAs.ToLower(),
                    InvitedAs = u.InvitedAs.ToLower(),
                    InvitedBy = u.InvitedBy != null ? u.InvitedBy.ToLower() : null,
                    WhenCreated = u.WhenCreated
                });

            Logger.Verbose($"{users.TotalUserCount} total external users to fetch");
            return users.TotalUserCount > startPosition + pageSize
                ? GetExternalUsersRecursive(siteCollection, startPosition + pageSize).Concat(myExternalUsers).ToList()
                : myExternalUsers.ToList();
        }

        public List<SharePointExternalUser> Execute(Site siteCollection)
        {
            Logger.Verbose($"About to execute {nameof(GetExternalUsersOnSiteCollection)}");
            return GetExternalUsersRecursive(siteCollection, 0);
        }
    }
}
