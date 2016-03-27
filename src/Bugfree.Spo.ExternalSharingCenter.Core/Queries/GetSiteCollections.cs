using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GetSiteCollections : Query
    {
        public GetSiteCollections(ILogger l) : base(l) { }

        public List<SiteCollection> Execute(ClientContext ctx)
        {
            Logger.Verbose($"About to execute {nameof(GetSiteCollections)}");

            var tenantSiteCollections = new GetTenantSiteCollections(Logger).Execute(ctx);
            Logger.Verbose($"{tenantSiteCollections.Count()} total site collections found on tenant");

            var withSharingEnabled = 
                tenantSiteCollections.Where(sc => 
                    sc.SharingCapability == SharingCapabilities.ExternalUserSharingOnly &&
                    (sc.Url.ToString().Contains("/sites/") || sc.Url.ToString().Contains("/teams/"))).ToList();
           
            Logger.Verbose($"{withSharingEnabled.Count()} site collections found with sharing enabled");
            return withSharingEnabled.Select(sc =>
            {
                using (var localCtx = new ClientContext(sc.Url) { Credentials = ctx.Credentials })
                {
                    var site = localCtx.Site;
                    localCtx.Load(site, s => s.Url);
                    localCtx.ExecuteQuery();

                    // locate candidate contact to use in case SharePoint reports InvitedBy as empty. 
                    // This is normal SharePoint behavior when InvitedBy is a site collection administrator.
                    var rootWeb = localCtx.Site.RootWeb;
                    var ownerGroup = rootWeb.AssociatedOwnerGroup;
                    localCtx.Load(rootWeb);
                    localCtx.Load(ownerGroup, g => g.Title, g => g.Users);
                    localCtx.ExecuteQuery();

                    var candidate = ownerGroup.Users.FirstOrDefault(u => u.Email.Contains("@"));
                    if (candidate == null) 
                    {
                        throw new InvalidOperationException($"To support external sharing, add a user with an email address to the '{ownerGroup.Title}' grouop on site collection '{sc.Url}'");
                    }

                    var externalUsers = new GetExternalUsersOnSiteCollection(Logger).Execute(site);
                    return new SiteCollection 
                    {
                        Url = new Uri(sc.Url),
                        Title = sc.Title,
                        FallbackOwnerMail = candidate.Email.ToLower(),
                        ExternalUsers = externalUsers
                    };
                }
            }).ToList();
        }
    }
}
