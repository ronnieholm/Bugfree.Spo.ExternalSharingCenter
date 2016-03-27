using System;
using System.Linq;
using System.Collections.Generic;
using Bugfree.Spo.Cqrs.Core;
using E = System.Xml.Linq.XElement;
using A = System.Xml.Linq.XAttribute;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GenerateExpirationMails : Query
    {
        public GenerateExpirationMails(ILogger l) : base(l) { }

        public List<Email> Execute(List<Expiration> expirations, Uri externalSharingCenterUrl, string expirationMailFrom)
        {
            Logger.Verbose($"About to execute {nameof(GenerateExpirationMails)}");
            return expirations.GroupBy(e => e.SharePointExternalUser.InvitedBy).Select(g =>
            {
                // initially, we included a link to the Edit action of 
                // SitePages/Site%20collection%20external%20user%20guide.aspx, but removed
                // the column because mail content is static and thus actual content may have 
                // changed since the mail went out, causing the link to be incorrect. Instead 
                // a link to always current SitePages Overview.aspx is included.
                var expirationLines = g.OrderBy(e => e.SiteCollection.Title).Select(e =>
                    new E("tr",
                        new E("td", 
                            new E("a", e.SiteCollection.Title,
                                new A("href", e.SiteCollection.Url))),
                        new E("td", e.SharePointExternalUser.WhenCreated.ToShortDateString()),
                        new E("td", e.ExpirationDate == DateTime.MinValue ? "N/A" : e.ExpirationDate.ToShortDateString()),
                        new E("td", e.SharePointExternalUser.InvitedAs),
                        new E("td", e.SharePointExternalUser.AcceptedAs),
                        new E("td", e.SharePointExternalUser.DisplayName)));

                var body =
                    new E("div",
                        new E("p", "External access has expired for users invited by you."),
                        new E("p", "The external users listed below no longer have access to their respective site collections."),
                        new E("p", 
                            new E("span", "Please visit the"),
                            new E("a", "Start",
                                new A("href", $"{externalSharingCenterUrl}/SitePages/Start.aspx")),
                            new E("span", "page within the"),
                            new E("a", "External sharing center",
                                new A("href", $"{externalSharingCenterUrl}")),
                            new E("span", "to extend the sharing period of any user before re-sharing from inside the site collection in question.")),
                        new E("table",
                            new E("tr",
                                new E("th", "Site collection title", new A("style", "text-align: left;")),
                                new E("th", "When created", new A("style", "text-align: left;")),
                                new E("th", "Expiration date", new A("style", "text-align: left;")),
                                new E("th", "Invited as", new A("style", "text-align: left;")),
                                new E("th", "Accepted as", new A("style", "text-align: left;")),
                                new E("th", "Display name", new A("style", "text-align: left;")),                                
                            expirationLines)));

                Logger.Verbose($"Grouped {expirationLines.Count()} expiration lines for recipient {g.Key}");
                return new Email
                {
                    From = expirationMailFrom,
                    To = g.Key,
                    Subject = "External user expirations: site collection access for users invited by you has expired",
                    Body = body,
                    Type = SentMailType.Expiration
                };
            }).ToList();
        }
    }
}
