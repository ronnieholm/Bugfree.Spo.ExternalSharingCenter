using System;
using System.Linq;
using System.Collections.Generic;
using Bugfree.Spo.Cqrs.Core;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GenerateExpirationWarnings : Query
    {
        public GenerateExpirationWarnings(ILogger l) : base(l) { }

        public List<ExpirationWarning> Execute(Database db, DateTime now)
        {
            Logger.Verbose($"About to execute {nameof(GenerateExpirationWarnings)}");
            var candidates = 
                (from sceu in db.SiteCollectionExternalUsers
                 from s in db.ExternalUsers
                 let sc = db.SiteCollections.Single(sc => sc.Url == sceu.SiteCollectionUrl)
                 where s.ExternaluserId == sceu.ExternalUserId
                 select new ExpirationWarning
                 {
                    SiteCollection = sc,
                    // on rare occassions the Accepted By mail isn't unique within a site collection. Turns out that 
                    // accounts of the form "i:0#.f|membership|live.com#foo@bar.com" may exist from a time when external 
                    // users were not registered in Azure Active Directory. They would exist only in the SharePoint side 
                    // of Office 365 (internally called SPODS). 
                    //
                    // Later on accounts of the form Account "i:0#.f|membership|foo@bar.com#ext#@mytenant.onmicrosoft.com" 
                    // may have been added to the same tenant as an external user which would then appear in Azure Active
                    // Directory. The SharePoint code that adds new users doesn't check for an existing SPODS user.
                    //
                    // Having two accounts with the same Accepted By mail is a situation that only exists in these specific 
                    // situations, and not something that can happen for new external user invites.
                    // 
                    // One manually has to remove one of the users as clearly one cannot login as different users with a 
                    // single Live account.
                    SharePointExternalUser = sc.ExternalUsers.Single(e => e.InvitedAs == s.Mail),
                    TimeUntilExpiration = sceu.End - now.ToUniversalTime(),
                    ExpirationDate = sceu.End
                 }).ToList();
            Logger.Verbose($"Generated {candidates.Count()} expiration warning candidates");
            return candidates;
        }
    }
}
