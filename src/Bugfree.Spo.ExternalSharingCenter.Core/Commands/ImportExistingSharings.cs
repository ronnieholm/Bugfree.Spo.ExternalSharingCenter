using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using Bugfree.Spo.ExternalSharingCenter.Core.Queries;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Commands
{
    public class ImportSettings
    {
        public string ExternalUserAddedComment { get; set; }
        public string SiteCollectionExternalUserAddedComment { get; set; }
        public DateTime SiteCollectionExternalUserStart { get; set; }
        public DateTime SiteCollectionExternalUserEnd { get; set; }
    }

    public class ImportExistingSharings : Command
    {
        private ImportSettings _settings;

        public ImportExistingSharings(ILogger l) : base(l) { }

        private void ImportExternalUsers(List externalUsers, Database db)
        {
            db.SharedSiteCollections
                .SelectMany(sc => sc.ExternalUsers)
                .Select(eu => eu.InvitedAs)
                .Distinct()
                .Except(db.ExternalUsers.Select(eu => eu.Mail))
                .ToList()
                .ForEach(eu =>
                    new AddExternalUser(Logger).Execute(
                        externalUsers,
                        new ExternalUser
                        {
                            ExternaluserId = Guid.NewGuid(),
                            Mail = eu,
                            Comment = _settings.ExternalUserAddedComment
                        }));
        }

        private void ImportSiteCollectionExternalUsers(List siteCollectionExternalUsers, Database db)
        {
            db.SharedSiteCollections
                .Select(sc =>
                    sc.ExternalUsers.Select(eu =>
                        new
                        {
                            SiteCollectionUrl = sc.Url,
                            ExternalUserId = db.ExternalUsers.Single(eu2 => eu.InvitedAs == eu2.Mail).ExternaluserId
                        }))
                .SelectMany(s => s)
                .Except(db.SiteCollectionExternalUsers.Select(sceu =>
                    new
                    {
                        sceu.SiteCollectionUrl,
                        sceu.ExternalUserId
                    }))
                .ToList()
                .ForEach(sceu =>
                    new AddSiteCollectionExternalUser(Logger).Execute(
                        siteCollectionExternalUsers,
                        new SiteCollectionExternalUser
                        {
                            SiteCollectionExternalUserId = Guid.NewGuid(),
                            SiteCollectionUrl = sceu.SiteCollectionUrl,
                            ExternalUserId = sceu.ExternalUserId,
                            Start = _settings.SiteCollectionExternalUserStart,
                            End = _settings.SiteCollectionExternalUserEnd,
                            Comment = _settings.SiteCollectionExternalUserAddedComment
                        }));
        }

        public void Execute(ImportSettings settings, List externalUsers, List siteCollectionExternalUsers, Database db)
        {
            Logger.Verbose($"About to execute {nameof(ImportExistingSharings)}");

            _settings = settings;
            ImportExternalUsers(externalUsers, db);

            // reload External users as import may have caused new rows to get added.
            // We require the ExternalUserId of new rows in order to add rows to Site
            // collection external users in the next stage.
            db.ExternalUsers = new GetExternalUsers(Logger).Execute(externalUsers);
            ImportSiteCollectionExternalUsers(siteCollectionExternalUsers, db);            
        }
    }
}
