using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using C = Bugfree.Spo.ExternalSharingCenter.Core.Commands.CreateExternalSharingCenterWeb.ExternalUserColumns;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Commands
{
    public class AddExternalUser : Command
    {
        public AddExternalUser(ILogger l) : base(l) { }

        public void Execute(List externalUsers, ExternalUser eu)
        {
            Logger.Verbose($"About to execute {nameof(AddExternalUser)} for mail '{eu.Mail}'");
            var i = externalUsers.AddItem(new ListItemCreationInformation());
            i[C.ExternalUserId] = eu.ExternaluserId;
            i[C.Mail] = eu.Mail;
            i[C.Comment] = eu.Comment;
            i.Update();
            externalUsers.Context.ExecuteQuery();
        }
    }
}
