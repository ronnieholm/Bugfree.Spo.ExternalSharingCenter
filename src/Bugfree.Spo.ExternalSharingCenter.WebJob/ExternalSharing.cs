using System;
using System.IO;
using Microsoft.Azure.WebJobs;
using Bugfree.Spo.ExternalSharingCenter.Core;
using Bugfree.Spo.Cqrs.Core;

namespace Bugfree.Spo.ExternalSharingCenter.WebJob
{
    public class ExternalSharing
    {
        [NoAutomaticTrigger]
        public static void Start(TextWriter tw)
        {
            var logger = new TextWriterLogger(tw);
            var settings = new Settings();
            try
            {
                var controller = new Controller(logger, settings);
                controller.ExpireUsers();
                controller.SendExpirationWarnings();
                controller.UpdateSiteCollections();
            }
            catch (Exception e)
            {
                logger.Error($"{e.Message} {e.StackTrace}");
                throw;
            }
        }
    }
}

