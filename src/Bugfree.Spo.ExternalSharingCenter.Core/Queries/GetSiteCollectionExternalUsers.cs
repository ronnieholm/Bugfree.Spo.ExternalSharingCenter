using System;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using C = Bugfree.Spo.ExternalSharingCenter.Core.Commands.CreateExternalSharingCenterWeb.SiteCollectionExternalUserColumns;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GetSiteCollectionExternalUsers : Query
    {
        public GetSiteCollectionExternalUsers(ILogger l) : base(l) { }

        private IEnumerable<ListItem> Query(List l, XElement caml)
        {
            var query = new CamlQuery { ViewXml = caml.ToString() };
            var results = l.GetItems(query);
            l.Context.Load(results);
            l.Context.ExecuteQuery();
            return results;
        }

        private IEnumerable<SiteCollectionExternalUser> Map(IEnumerable<ListItem> items)
        {
            Func<object, int> intMapper = o =>
            {
                var d = (double)o;
                return (int)d;
            };

            Func<object, Guid> guidMapper = o => 
            {
                var s = (string)o;
                return Guid.Parse(s);
            };

            Func<object, Uri> uriMapper = o => {
                var s = (string)o;
                return new Uri(s);
            };

            return items.Select(i =>
                new SiteCollectionExternalUser
                {
                    Id = i.Id,
                    SiteCollectionExternalUserId = guidMapper(i[C.SiteCollectionExternalUserId]),
                    SiteCollectionUrl = uriMapper(i[C.SiteCollectionUrl]),
                    ExternalUserId = guidMapper(i[C.ExternalUserId]),
                    Start = (DateTime)i[C.Start],
                    End = (DateTime)i[C.End],
                    Comment = (string)i[C.Comment],
                    Created = (DateTime)i["Created"]
                }
            );
        }

        public List<SiteCollectionExternalUser> Execute(List l)
        {
            Logger.Verbose($"About to execute {nameof(GetSiteCollectionExternalUsers)}");
            var siteCollectionSharings = Map(Query(l, XElement.Parse(CamlQuery.CreateAllItemsQuery().ViewXml))).ToList();
            Logger.Verbose($"{siteCollectionSharings.Count()} site collection sharings found");
            return siteCollectionSharings;
        }
    }
}
