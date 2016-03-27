using System;
using System.Linq;
using System.Collections.Generic;
using Bugfree.Spo.Cqrs.Core;
using E = System.Xml.Linq.XElement;
using A = System.Xml.Linq.XAttribute;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GenerateExpirationWarningMails : Query
    {
        public GenerateExpirationWarningMails(ILogger l) : base(l) { }

        // running this command without the risk of exceptions being thrown assumes GenerateExpirations 
        // was run beforehand to clear out any invalid entries in the Site collection external users. 
        // Say the list contains a sharing for a site collection without sharing enabled, then there'll 
        // be no site collection to lookup based on foreign key inside the Site collection external 
        // sharing list. In case this command cannot guarantee the validity of its result, it should fail 
        // rather than fix-up errornous data.
        public List<Email> Execute(List<ExpirationWarning> expirationWarnings, Uri externalSharingCenterUrl, string expirationMailFrom)
        {
            Logger.Verbose($"About to execute {nameof(GenerateExpirationWarningMails)}");
            return expirationWarnings.GroupBy(e => e.SharePointExternalUser.InvitedBy).Select(g =>
            {
                var warningLines = g.OrderBy(e => e.SiteCollection.Title).Select(e =>
                    new E("tr",
                        new E("td", 
                            new E("a", e.SiteCollection.Title,
                                new A("href", e.SiteCollection.Url))),
                        new E("td", e.SharePointExternalUser.WhenCreated.ToShortDateString()),
                        new E("td", e.ExpirationDate.ToShortDateString()),
                        new E("td", e.SharePointExternalUser.InvitedAs),
                        new E("td", e.SharePointExternalUser.AcceptedAs),
                        new E("td", e.SharePointExternalUser.DisplayName)));

                var body =
                    new E("div",
                        new E("p", "External access is about to expire for users invited by you."),
                        new E("p", "The external users listed below are about to lose access to their respective site collections."),
                        new E("p",
                            new E("span", "Please visit the"),
                            new E("a", "Start",
                                new A("href", $"{externalSharingCenterUrl}/SitePages/Start.aspx")),
                            new E("span", "page within the"),
                            new E("a", "External sharing center",
                                new A("href", $"{externalSharingCenterUrl}")),
                            new E("span", "to extend the sharing period of any user.")),
                        new E("table",
                            new E("tr",
                                new E("th", "Site collection title", new A("style", "text-align: left;")),
                                new E("th", "When created", new A("style", "text-align: left;")),
                                new E("th", "Expiration date", new A("style", "text-align: left;")),
                                new E("th", "Invited as", new A("style", "text-align: left;")),
                                new E("th", "Accepted as", new A("style", "text-align: left;")),
                                new E("th", "Display name", new A("style", "text-align: left;"))),
                            warningLines));

                Logger.Verbose($"Grouped {warningLines.Count()} expiration warning lines for recipient {g.Key}");
                return new Email
                {
                    From = expirationMailFrom,
                    To = g.Key,
                    Subject = "External user expiration warnings: site collection access for users invited by you is about to expire",
                    Body = body,
                    Type = SentMailType.Warning                    
                };
            }).ToList();
        }
    }
}
