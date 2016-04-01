using Bugfree.Spo.Cqrs.Core;
using Microsoft.SharePoint.Client;
using System;
using C = Bugfree.Spo.ExternalSharingCenter.Core.Commands.CreateExternalSharingCenterWeb.SiteCollectionColumns;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Commands {
    public class AddSiteCollection : Command 
    {
        public AddSiteCollection(ILogger l) : base(l) { }

        public void Execute(List siteCollections, SharedSiteCollection s) 
        {
            Logger.Verbose($"About to execute {nameof(AddSiteCollection)} for title '{s.Title}'");
            var i = siteCollections.AddItem(new ListItemCreationInformation());
            i[C.SiteCollectionId] = Guid.NewGuid();
            i[C.SiteCollectionTitle] = s.Title;
            i[C.SiteCollectionUrl] = s.Url;
            i[C.Comment] = "";
            i.Update();
            siteCollections.Context.ExecuteQuery();
        }
    }
}
