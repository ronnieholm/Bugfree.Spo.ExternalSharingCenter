﻿using System;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using C = Bugfree.Spo.ExternalSharingCenter.Core.Commands.CreateExternalSharingCenterWeb.SentMailColumns;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Queries
{
    public class GetSentMails : Query
    {
        public GetSentMails(ILogger l) : base(l) { }

        private IEnumerable<ListItem> Query(List l, XElement caml)
        {
            var query = new CamlQuery { ViewXml = caml.ToString() };
            var results = l.GetItems(query);
            l.Context.Load(results);
            l.Context.ExecuteQuery();
            return results;
        }

        private IEnumerable<SentMail> Map(IEnumerable<ListItem> items)
        {
            return items.Select(i =>
                new SentMail
                {
                    Id = i.Id,
                    From = (string)i[C.From],
                    To = (string)i[C.To],
                    Subject = (string)i[C.Subject],
                    Body = XElement.Parse((string)i[C.Body]),
                    Type = (SentMailType)(int.Parse(i[C.SentMailType].ToString())),
                    Created = (DateTime)i["Created"]
                }
            );
        }

        public List<SentMail> Execute(List l)
        {
            Logger.Verbose($"About to execute {nameof(GetSentMails)}");
            var sentMails = Map(Query(l, XElement.Parse(CamlQuery.CreateAllItemsQuery().ViewXml))).ToList();
            Logger.Verbose($"{sentMails.Count()} sent mails found");
            return sentMails;
        }
    }
}
