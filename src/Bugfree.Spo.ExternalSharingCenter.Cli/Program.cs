using System;
using Bugfree.Spo.Cqrs.Core;
using Bugfree.Spo.ExternalSharingCenter.Core;

namespace Bugfree.Spo.ExternalSharingCenter.Cli
{
    class Program
    {
        static void Main()
        {
            var logger = new ColoredConsoleLogger();
            var settings = new Settings();
            var controller = new Controller(logger, settings);

            // comment in parts of the code below depending on what operation you want the
            // CLI to carry out.

            // 1. Sets up the External Sharing Center web as per App.config settings
            //controller.SetupExternalSharingCenterWeb();

            // 2. Iterates site collection with sharing enabled and imports external users
            //    and their sharings into External users and Site collection external users
            //    lists. Lack start and end dates for existing sharing, default values are
            //    provided.
            //controller.ImportExistingSharings(
            //    new Core.Commands.ImportSettings {
            //        ExternalUserAddedComment = "Imported by tool",
            //        SiteCollectionExternalUserAddedComment = "Imported by tool",
            //        SiteCollectionExternalUserStart = DateTime.Today,
            //        SiteCollectionExternalUserEnd = DateTime.Today.AddDays(180).AddSeconds(-1)
            //    });

            // 3. These are the day to day operations. Either run through the CLI or as
            //    part of the Azure WebJob.
            //controller.ExpireUsers();
            //controller.SendExpirationWarnings();
        }
    }
}
