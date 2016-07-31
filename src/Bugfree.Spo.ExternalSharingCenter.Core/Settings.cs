using System;
using System.Configuration;

namespace Bugfree.Spo.ExternalSharingCenter
{
    public class Settings
    {
        public Uri ExternalSharingCenterUrl { get; set; }
        public string TenantAdministratorUsername { get; private set; }
        public string TenantAdministratorPassword { get; private set; }
        public int ExpirationWarningMailsMinimumDaysBetween { get; set; }
        public int ExpirationWarningDays { get; set; }

        public string MailFrom { get; set; }

        public Settings()
        {
            Initialize();
        }

        private string GetString(string k) => ConfigurationManager.AppSettings[k];
        
        private void Initialize()
        {
            ExternalSharingCenterUrl = new Uri(GetString("ExternalSharingCenterUrl"));
            TenantAdministratorUsername = GetString("TenantAdministratorUsername");
            TenantAdministratorPassword = GetString("TenantAdministratorPassword");
            ExpirationWarningDays = int.Parse(GetString("ExpirationWarningDays"));
            ExpirationWarningMailsMinimumDaysBetween = int.Parse(GetString("ExpirationWarningMailsMinimumDaysBetween"));
            MailFrom = GetString("MailFrom");
        }
    }
}
