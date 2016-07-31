using System;
using System.Linq;
using System.Collections.Generic;
using Bugfree.Spo.Cqrs.Core;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GenerateExpirationWarnings : Query
    {
        public GenerateExpirationWarnings(ILogger l) : base(l) { }

        public List<ExpirationWarning> Execute(Database db, DateTime current, TimeSpan expirationPeriod)
        {
            Logger.Verbose($"About to execute {nameof(GenerateExpirationWarnings)}");

            var warnings = new List<ExpirationWarning>();
            foreach (var sc in db.SharedSiteCollections) 
            {
                foreach (var eu in sc.ExternalUsers) 
                {
                    // from the previous run of GenerateExpiration, it's safe to assume the SiteCollectionExternalUsers
                    // and ExternalUsers contains entries for the sharing. Otherwise, the actual site collection sharing 
                    // would've been removed.
                    var externalUser = db.ExternalUsers.Single(u => u.Mail == eu.InvitedAs);
                    var siteCollectionExternalUser =
                        db.SiteCollectionExternalUsers
                            .Single(r => r.SiteCollectionUrl == sc.Url && 
                                    r.ExternalUserId == externalUser.ExternaluserId);

                    if (siteCollectionExternalUser.End - current < expirationPeriod) 
                    {
                        warnings.Add(new ExpirationWarning 
                        {
                            SiteCollection = sc,
                            SharePointExternalUser = eu,
                            TimeUntilExpiration = siteCollectionExternalUser.End - current,
                            ExpirationDate = siteCollectionExternalUser.End
                        });
                    }
                }
            }
            return warnings;
        }
    }
}
