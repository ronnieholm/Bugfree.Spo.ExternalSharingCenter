using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using System.Security;

namespace Bugfree.Spo.ExternalSharingCenter.StressTest 
{
    public struct SharePointExternalUser 
    {
        public int UserId { get; set; }
        public string DisplayName { get; set; }
        public string AcceptedAs { get; set; }
        public string InvitedAs { get; set; }
        public string InvitedBy { get; set; }
        public DateTime WhenCreated { get; set; }
    }

    class Program 
    {
        private static List<SharePointExternalUser> GetExternalUsersRecursive(Site siteCollection, int startPosition) 
        {
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

            return users.TotalUserCount > startPosition + pageSize
                ? GetExternalUsersRecursive(siteCollection, startPosition + pageSize).Concat(myExternalUsers).ToList()
                : myExternalUsers.ToList();
        }

        static void Main(string[] args) 
        {
            var username = "<user>@<tenant>.onmicrosoft.com";
            var password = "<password>";
            var siteCollection = "https://<tenant>.sharepoint.com/sites/<siteCollection>";

            var securePassword = new SecureString();
            password.ToCharArray().ToList().ForEach(securePassword.AppendChar);
            var ctx = new ClientContext(siteCollection) 
            {
                Credentials = new SharePointOnlineCredentials(username, securePassword)
            };

            ctx.Load(ctx.Site, s => s.Url);
            ctx.ExecuteQuery();

            var baseline = GetExternalUsersRecursive(ctx.Site, 0);

            do 
            {
                var current = GetExternalUsersRecursive(ctx.Site, 0);
                var delta = baseline.Except(current).ToList();

                Console.WriteLine(DateTime.Now);
                if (delta.Count() > 0) 
                {
                    delta.ForEach(s => Console.WriteLine(s.DisplayName));
                    break;
                }               
            } while (true);
        }
    }
}
