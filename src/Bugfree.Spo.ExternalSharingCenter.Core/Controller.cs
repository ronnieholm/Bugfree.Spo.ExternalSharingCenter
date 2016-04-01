/*
    Setting up Bugfree.Spo.ExternalSharingCenter.Core.csproj
    --------------------------------------------------------
 
    The Bugfree.Spo.ExternalSharingCenter.Core project was created as a regular 
    Class library. The project file was then hand-modified to make TypeScript
    work like in a web project, i.e., with compile on save. The following line
    must be added to the project file:

        <Import Project="$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v14.0\TypeScript\Microsoft.TypeScript.targets" />

    The angular.d.ts and jquery.d.ts files were downloaded from the DefinitelyTyped 
    Github repository. At the top of angular.d.ts, the reference to jquery was 
    modified from

        /// <reference path="../jquery/jquery.d.ts" />

    to

        /// <reference path="jquery.d.ts" />

    because we keep both d.ts files in the same folder. Otherwise we get a compiler
    error.
 */

using System;
using System.Linq;
using System.Security;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using Bugfree.Spo.ExternalSharingCenter.Core.Commands;
using Bugfree.Spo.ExternalSharingCenter.Core.Queries;
using C = Bugfree.Spo.ExternalSharingCenter.Core.Commands.CreateExternalSharingCenterWeb;

namespace Bugfree.Spo.ExternalSharingCenter.Core
{
    public class Controller : IDisposable
    {
        Database _db;
        Settings _settings;
        ClientContext _context;
        ILogger _logger;

        public Controller(ILogger logger, Settings settings)
        {
            _logger = logger;
            _settings = settings;
        }

        private void Initialize()
        {
            _context = CreateClientContext(_settings.ExternalSharingCenterUrl);
            var lists = _context.Web.Lists;
            var externalUserList = lists.GetByTitle(C.ExternalUsersTitle);
            var siteCollectionExternalUserList = lists.GetByTitle(C.SiteCollectionExternalUsersTitle);
            _context.Load(externalUserList);
            _context.Load(siteCollectionExternalUserList);
            _context.ExecuteQuery();
            _db = new Database
            {
                ExternalUsers = new GetExternalUsers(_logger).Execute(externalUserList),
                SiteCollectionExternalUsers = new GetSiteCollectionExternalUsers(_logger).Execute(siteCollectionExternalUserList),
                SharedSiteCollections = new GetSharedSiteCollections(_logger).Execute(_context)
            };

            // helps diagnose issues with the slightly buggy tanant API
            var siteCollectionsCount = _db.SharedSiteCollections.Count();
            var externalUsersCount = 0;
            foreach (var sc in _db.SharedSiteCollections) {
                externalUsersCount += sc.ExternalUsers.Count();
                _logger.Verbose($"{sc.ExternalUsers.Count()} {sc.Url}");
            }
            _logger.Verbose($"Found {externalUsersCount} users in total");
        }

        public ClientContext CreateClientContext(Uri url)
        {
            var securePassword = new SecureString();
            _settings.TenantAdministratorPassword.ToCharArray().ToList().ForEach(securePassword.AppendChar);
            return new ClientContext(url)
            {
                Credentials = new SharePointOnlineCredentials(_settings.TenantAdministratorUsername, securePassword)
            };
        }

        public void SetupExternalSharingCenterWeb()
        {
            using (var ctx = CreateClientContext(_settings.ExternalSharingCenterUrl))
            {
                ctx.Load(ctx.Web, x => x.Language, x => x.WebTemplate, x => x.Configuration);
                ctx.ExecuteQuery();
                var w = ctx.Web;
                if (w.Language != 1033 || w.WebTemplate != "STS" || w.Configuration != 0) 
                {
                    throw new NotSupportedException("Only team sites based on English language template are supported");
                }

                new C(_logger).Execute(ctx);
            }
        }

        public void ImportExistingSharings(ImportSettings settings)
        {
            Initialize();
            var lists = _context.Web.Lists;
            var externalUsersList = lists.GetByTitle(C.ExternalUsersTitle);
            var siteCollectionExternalUsersLists = lists.GetByTitle(C.SiteCollectionExternalUsersTitle);
            _context.Load(externalUsersList);
            _context.Load(siteCollectionExternalUsersLists);
            _context.ExecuteQuery();
            new ImportExistingSharings(_logger).Execute(settings, externalUsersList, siteCollectionExternalUsersLists, _db);
        }

        public void SendExpirationWarnings()
        {
            Initialize();
            var warnings = new GenerateExpirationWarnings(_logger).Execute(_db, DateTime.Now);
            warnings.Where(n => n.SharePointExternalUser.InvitedBy == null).ToList().ForEach(n =>
            {
                var sc = _db.SharedSiteCollections.Single(sc1 => sc1.Url == n.SiteCollection.Url);
                n.SharePointExternalUser.InvitedBy = sc.FallbackOwnerMail;
            });

            var withinWarningPeriod = warnings.Where(e => e.TimeUntilExpiration < new TimeSpan(_settings.ExpirationWarningDays, 0, 0, 0)).ToList();
            var warningMails = new GenerateExpirationWarningMails(_logger).Execute(withinWarningPeriod, _settings.ExternalSharingCenterUrl, _settings.MailFrom);
            var sentMails = _context.Web.Lists.GetByTitle(C.SentMailTitle);
            _context.Load(sentMails);
            _context.ExecuteQuery();

            var defaultTimeStamp = new TimeSpan();
            var sentMailArchieve = new GetSentMails(_logger).Execute(sentMails);
            var warningMailsThrottled =
                warningMails.Where(m =>
                {
                    var last = sentMailArchieve
                        .Where(a => m.To == a.To && a.Type == SentMailType.Warning)
                        .Select(a => DateTime.UtcNow - a.Created)
                        .LastOrDefault();
                    return last == defaultTimeStamp
                        ? true
                        : last.Days >= _settings.ExpirationWarningMailsMinimumDaysBetween;
                }).ToList();

            //warningMailsThrottled.ForEach(e =>
            //{
            //    var siteCollectionWithRecipientAsSiteUser = warnings.First(e1 => e1.SharePointExternalUser.InvitedBy == e.To);
            //    using (var warningCtx = CreateClientContext(siteCollectionWithRecipientAsSiteUser.SiteCollection.Url))
            //    {
            //        new SendMail(_logger).Execute(warningCtx, e);
            //        new AddSentMail(_logger).Execute(sentMails, e);
            //    }
            //});
        }

        public void ExpireUsers()
        {
            Initialize();
            var expirations = new GenerateExpirations(_logger).Execute(_db, DateTime.Now);            
            expirations.Where(e => e.SharePointExternalUser.InvitedBy == null).ToList().ForEach(e =>
            {
                var sc = _db.SharedSiteCollections.Single(s => s.Url == e.SiteCollection.Url);
                e.SharePointExternalUser.InvitedBy = sc.FallbackOwnerMail;
            });

            var expirationMails = new GenerateExpirationMails(_logger).Execute(expirations, _settings.ExternalSharingCenterUrl, _settings.MailFrom);
            var sentMails = _context.Web.Lists.GetByTitle(C.SentMailTitle);
            _context.Load(sentMails);
            _context.ExecuteQuery();

            //expirationMails.ForEach(e =>
            //{
            //    var siteCollectionWithRecipientAsSiteUser = expirations.First(e1 => e1.SharePointExternalUser.InvitedBy == e.To);
            //    using (var expirationMailContext = CreateClientContext(siteCollectionWithRecipientAsSiteUser.SiteCollection.Url))
            //    {
            //        new SendMail(_logger).Execute(expirationMailContext, e);
            //        new AddSentMail(_logger).Execute(sentMails, e);
            //    }
            //});

            //expirations.ForEach(e => 
            //{
            //    using (var expirationMailCtx = CreateClientContext(e.SiteCollection.Url)) 
            //    {
            //        new RemoveUserFromSiteCollection(_logger).Execute(expirationMailCtx, e.SharePointExternalUser.UserId);
            //    }
            //});
        }

        public void UpdateSiteCollections() 
        {
            Initialize();

            var lists = _context.Web.Lists;
            var siteCollectionsList = lists.GetByTitle(C.SiteCollectionsTitle);

            var items = siteCollectionsList.GetItems(CamlQuery.CreateAllItemsQuery());
            siteCollectionsList.Context.Load(items);
            siteCollectionsList.Context.ExecuteQuery();
            items.ToList().ForEach(i => 
            {
                i.DeleteObject();
                siteCollectionsList.Update();
            });

            siteCollectionsList.Context.Load(items);
            siteCollectionsList.Context.ExecuteQuery();
            if (items.Count != 0) 
            {
                throw new InvalidOperationException(
                    $"{C.SiteCollectionsTitle} was expected to hold 0 records. Actual count: {items.Count}");
            }

            _db.SharedSiteCollections.ForEach(ssc => 
                new AddSiteCollection(_logger).Execute(siteCollectionsList, ssc));
        }

        public void Dispose()
        {
            if (_context != null)
            {
                _context.Dispose();
            }
        }
    }
}
