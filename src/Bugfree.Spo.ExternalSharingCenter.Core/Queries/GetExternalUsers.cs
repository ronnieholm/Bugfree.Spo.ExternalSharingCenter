using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using C = Bugfree.Spo.ExternalSharingCenter.Core.Commands.CreateExternalSharingCenterWeb.ExternalUserColumns;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GetExternalUsers : Query
    {      
        public GetExternalUsers(ILogger l) : base(l) { }

        private IEnumerable<ListItem> Query(List l, XElement caml)
        {
            var query = new CamlQuery { ViewXml = caml.ToString() };
            var results = l.GetItems(query);
            l.Context.Load(results);
            l.Context.ExecuteQuery();
            return results;
        }

        private IEnumerable<ExternalUser> Map(IEnumerable<ListItem> items)
        {
            Func<object, Guid> guidMapper = o => 
            {
                var s = (string)o;
                return Guid.Parse(s);
            };

            return items.Select(i =>
                new ExternalUser
                {
                    Id = i.Id,
                    ExternaluserId = guidMapper(i[C.ExternalUserId]),
                    Mail = ((string)i[C.Mail]).ToLower(),
                    Comment = (string)i[C.Comment]
                }
            );
        }

        public List<ExternalUser> Execute(List l)
        {
            Logger.Verbose($"About to execute {nameof(GetExternalUsers)}");
            var externalUsers = Map(Query(l, XElement.Parse(CamlQuery.CreateAllItemsQuery().ViewXml))).ToList();
            Logger.Verbose($"{externalUsers.Count()} external users found");
            return externalUsers;
        }
    }
}
