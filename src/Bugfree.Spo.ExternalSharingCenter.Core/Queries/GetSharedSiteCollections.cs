using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using Bugfree.Spo.Cqrs.Core.Queries;
using Bugfree.Spo.Cqrs.Core.Utilities;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GetSharedSiteCollections : Query
    {
        public GetSharedSiteCollections(ILogger l) : base(l) { }

        public List<SharedSiteCollection> Execute(ClientContext ctx)
        {
            Logger.Verbose($"About to execute {nameof(GetSharedSiteCollections)}");

            var url = ctx.Url;
            var tenantAdminUrl = new AdminUrlInferrer().InferAdminFromTenant(new Uri(url.Replace(new Uri(url).AbsolutePath, "")));

            var tenantSiteCollections = new List<SiteProperties>();
            using (var tenantContext = new ClientContext(tenantAdminUrl) { Credentials = ctx.Credentials }) 
            { 
                tenantSiteCollections = new GetTenantSiteCollections(Logger).Execute(tenantContext);
            }
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
                        Logger.Warning($"To support external sharing, add a user with an email address to the '{ownerGroup.Title}' group on site collection '{sc.Url}'. Otherwise site collection sharings are eventually deleted without warning or experiration mails sent");
                    }

                    var externalUsers = new GetExternalUsersOnSiteCollection(Logger).Execute(site);
                    return new SharedSiteCollection 
                    {
                        Url = new Uri(sc.Url),
                        Title = sc.Title,
                        FallbackOwnerMail = candidate != null ? candidate.Email.ToLower() : null,
                        ExternalUsers = externalUsers
                    };
                }
            }).ToList();
        }
    }
}
