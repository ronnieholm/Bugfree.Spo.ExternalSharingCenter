using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using C = Bugfree.Spo.ExternalSharingCenter.Core.Commands.CreateExternalSharingCenterWeb.SentMailColumns;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Commands
{
    public class AddSentMail : Command
    {
        public AddSentMail(ILogger l) : base(l) { }

        public void Execute(List sentMails, Email mail)
        {
            Logger.Verbose($"About to execute {nameof(AddSentMail)} for mail from '{mail.From}' to '{mail.To}'");
            var i = sentMails.AddItem(new ListItemCreationInformation());
            i[C.From] = mail.From;
            i[C.To] = mail.To;
            i[C.Subject] = mail.Subject;
            i[C.Body] = mail.Body;
            i[C.SentMailType] = (int)mail.Type;
            i[C.Comment] = "";
            i.Update();
            sentMails.Context.ExecuteQuery();
        }
    }
}
