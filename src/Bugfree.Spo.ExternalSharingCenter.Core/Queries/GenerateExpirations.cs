using System;
using System.Linq;
using System.Collections.Generic;
using Bugfree.Spo.Cqrs.Core;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GenerateExpirations : Query
    {
        public GenerateExpirations(ILogger l) : base(l) { }

        public List<Expiration> Execute(Database db, DateTime current)
        {
            Logger.Verbose($"About to execute {nameof(GenerateExpirations)}");
            var expirations = new List<Expiration>();
            foreach (var sc in db.SharedSiteCollections)
            {
                foreach (var eu in sc.ExternalUsers)
                {
                    var externalUserCandidate = db.ExternalUsers.SingleOrDefault(u => u.Mail == eu.InvitedAs);

                    // when external user isn't present in External users list, it indicates the user 
                    // was added without going through the External Sharing Center. Hence we expire 
                    // the user.
                    if (externalUserCandidate == null)
                    {
                        Logger.Verbose($"No record in external users list of user '{eu.InvitedAs}'");
                        expirations.Add(
                            new Expiration
                            {
                                SiteCollection = sc,
                                SharePointExternalUser = eu
                            });
                    } 
                    else
                    {
                        var siteCollectionExternalUserCandidate = 
                            db.SiteCollectionExternalUsers
                                .SingleOrDefault(sceu => 
                                    sceu.SiteCollectionUrl == sc.Url && 
                                    sceu.ExternalUserId == externalUserCandidate.ExternaluserId);

                        // Site collection external users doesn't contain a record of a user for the
                        // site collection question. Hence we expire the user.
                        if (siteCollectionExternalUserCandidate == null)
                        {
                            Logger.Verbose($"No record in site collection external users for '{eu.InvitedAs}'");
                            expirations.Add(
                                new Expiration
                                {
                                    SiteCollection = sc,
                                    SharePointExternalUser = eu
                                });
                        }
                        else
                        {
                            var valid = siteCollectionExternalUserCandidate.Start < current && siteCollectionExternalUserCandidate.End > current;
                            if (!valid)
                            {
                                Logger.Verbose($"Site collection external user contains expired record for {eu.InvitedAs}");
                                expirations.Add(
                                    new Expiration
                                    {
                                        SiteCollection = sc,
                                        SharePointExternalUser = eu,
                                        ExpirationDate = siteCollectionExternalUserCandidate.End
                                    });
                            }
                        }
                    }
                }
            }
            return expirations;
        }
    }
}