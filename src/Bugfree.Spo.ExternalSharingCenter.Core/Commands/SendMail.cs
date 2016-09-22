using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Bugfree.Spo.Cqrs.Core;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Commands
{
    public class SendMail : Command
    {
        public SendMail(ILogger l) : base(l) { }

        public void Execute(ClientContext ctx, Email mail)
        {
            Logger.Verbose($"About to execute {nameof(SendMail)} for recipient '{mail.To}'");
            Utility.SendEmail(ctx, new EmailProperties
            {
                // Starting September 2, 2016 emails sent through SharePoint Online
                // originate from no-reply@sharepointonline.com. Specifying a different
                // from address is ignored by SharePoint Online.
                From = mail.From,
                To = new[] { mail.To },
                Subject = mail.Subject,
                Body = mail.Body.ToString(),
            });
            ctx.ExecuteQuery();
        }
    }
}