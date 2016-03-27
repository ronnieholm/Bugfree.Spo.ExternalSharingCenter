using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using Bugfree.Spo.Cqrs.Core;
using Bugfree.Spo.Cqrs.Core.Commands;
using E = System.Xml.Linq.XElement;
using A = System.Xml.Linq.XAttribute;

namespace Bugfree.Spo.ExternalSharingCenter.Core.Commands
{
    public class CreateExternalSharingCenterWeb : Command
    {
        public const string AllItems = "All Items";
        public const string StartAspx = "Start.aspx";
        public const string SiteCollectionExternalUserGuideAspx = "Site collection external user guide.aspx";

        public const string SiteCollectionExternalUsersTitle = "Site collection external users";
        public const string SiteCollectionExternalUsersUrl = "SiteCollectionExternalUsers";
        public class SiteCollectionExternalUserColumns
        {
            public const string SiteCollectionExternalUserId = nameof(SiteCollectionExternalUserId);
            public const string SiteCollectionUrl = nameof(SiteCollectionUrl);
            public const string ExternalUserId = nameof(ExternalUserId);
            public const string Start = nameof(Start);
            public const string End = nameof(End);
            public const string Comment = nameof(Comment);
        }
        
        public const string ExternalUsersTitle = "External users";
        public const string ExternalUsersUrl = "ExternalUsers";
        public class ExternalUserColumns
        {
            public const string ExternalUserId = nameof(ExternalUserId);
            public const string Mail = nameof(Mail);
            public const string Comment = nameof(Comment);
        }

        public const string SentMailTitle = "Sent mail";
        public const string SentMailUrl = "SentMail";
        public class SentMailColumns
        {
            public const string From = nameof(From);
            public const string To = nameof(To);
            public const string Subject = nameof(Subject);
            public const string Body = nameof(Body);
            public const string SentMailType = nameof(SentMailType);
            public const string Comment = nameof(Comment);
        }

        public CreateExternalSharingCenterWeb(ILogger l) : base(l) { }

        private void CreateExternalUsersList(ClientContext ctx)
        {
            var lists = ctx.Web.Lists;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            var candidate = lists.SingleOrDefault(l => l.Title == ExternalUsersTitle);
            if (candidate != null)
            {
                Logger.Warning($"List '{ExternalUsersTitle}' already exists");
                return;
            }

            new CreateListFromTemplate(Logger)
                .Execute(ctx, ListTemplateType.GenericList, ExternalUsersUrl, l =>
                {
                    l.OnQuickLaunch = false;
                    l.EnableVersioning = true;
                    l.Title = ExternalUsersTitle;

                    var title = l.Fields.GetByTitle("Title");
                    title.Required = false;
                    title.Update();
                });

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    ExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Text"),
                        new A("DisplayName", ExternalUserColumns.ExternalUserId),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "TRUE"),
                        // when EnforceUniqueValues = true, field must be indexed
                        // or provisioning throws an exception. Uniqueness constraint
                        // is validated by REST service on insert whereas required
                        // only applied working through the browser.
                        new A("Indexed", "TRUE"),
                        new A("MaxLength", "255"),
                        new A("StaticName", ExternalUserColumns.ExternalUserId),
                        new A("Name", ExternalUserColumns.ExternalUserId)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    ExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Text"),
                        new A("DisplayName", ExternalUserColumns.Mail),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "TRUE"),
                        new A("Indexed", "TRUE"),
                        new A("MaxLength", "255"),
                        new A("StaticName", ExternalUserColumns.Mail),
                        new A("Name", ExternalUserColumns.Mail)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    ExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Note"),
                        new A("DisplayName", ExternalUserColumns.Comment),
                        new A("Required", "FALSE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("NumLines", "6"),
                        new A("RichText", "FALSE"),
                        new A("Sortable", "FALSE"),
                        new A("AppendOnly", true),
                        new A("StaticName", ExternalUserColumns.Comment),
                        new A("Name", ExternalUserColumns.Comment)));

            new RemoveListView(Logger).Execute(ctx, ExternalUsersTitle, AllItems);
            new CreateListView(Logger)
                .Execute(
                    ctx,
                    ExternalUsersTitle,
                    AllItems,
                    new[]
                    {
                        "ID",
                        "LinkTitle",
                        ExternalUserColumns.ExternalUserId,
                        ExternalUserColumns.Mail,
                        ExternalUserColumns.Comment
                    },
                    v =>
                    {
                        v.DefaultView = true;
                        v.Paged = true;
                        v.ViewQuery =
                            new E("OrderBy",
                                new E("FieldRef",
                                    new A("Name", "ID"),
                                    new A("Ascending", "FALSE"))).ToString();
                    });
        }

        private void CreateSiteCollectionExternalUsersList(ClientContext ctx)
        {
            var lists = ctx.Web.Lists;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            var candidate = lists.SingleOrDefault(l => l.Title == SiteCollectionExternalUsersTitle);
            if (candidate != null)
            {
                Logger.Warning($"List '{SiteCollectionExternalUsersTitle}' already exists");
                return;
            }

            new CreateListFromTemplate(Logger)
                .Execute(ctx, ListTemplateType.GenericList, SiteCollectionExternalUsersUrl, l =>
                {
                    l.OnQuickLaunch = false;
                    l.EnableVersioning = true;
                    l.Title = SiteCollectionExternalUsersTitle;

                    var title = l.Fields.GetByTitle("Title");
                    title.Required = false;
                    title.Update();
                });

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SiteCollectionExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Text"),
                        new A("DisplayName", SiteCollectionExternalUserColumns.SiteCollectionExternalUserId),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "TRUE"),
                        new A("Indexed", "TRUE"),
                        new A("MaxLength", "255"),
                        new A("StaticName", SiteCollectionExternalUserColumns.SiteCollectionExternalUserId),
                        new A("Name", SiteCollectionExternalUserColumns.SiteCollectionExternalUserId)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SiteCollectionExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Text"),
                        new A("DisplayName", SiteCollectionExternalUserColumns.SiteCollectionUrl),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("MaxLength", "255"),
                        new A("StaticName", SiteCollectionExternalUserColumns.SiteCollectionUrl),
                        new A("Name", SiteCollectionExternalUserColumns.SiteCollectionUrl)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SiteCollectionExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Text"),
                        new A("DisplayName", SiteCollectionExternalUserColumns.ExternalUserId),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("MaxLength", "255"),
                        new A("StaticName", SiteCollectionExternalUserColumns.ExternalUserId),
                        new A("Name", SiteCollectionExternalUserColumns.ExternalUserId)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SiteCollectionExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "DateTime"),
                        new A("DisplayName", SiteCollectionExternalUserColumns.Start),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("Format", "DateOnly"),
                        new A("FriendlyDisplayFormat", "Disabled"),
                        new A("StaticName", SiteCollectionExternalUserColumns.Start),
                        new A("Name", SiteCollectionExternalUserColumns.Start)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SiteCollectionExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "DateTime"),
                        new A("DisplayName", SiteCollectionExternalUserColumns.End),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("Format", "DateOnly"),
                        new A("FriendlyDisplayFormat", "Disabled"),
                        new A("StaticName", SiteCollectionExternalUserColumns.End),
                        new A("Name", SiteCollectionExternalUserColumns.End)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SiteCollectionExternalUsersTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Note"),
                        new A("DisplayName", SiteCollectionExternalUserColumns.Comment),
                        new A("Required", "FALSE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("NumLines", "6"),
                        new A("RichText", "FALSE"),
                        new A("Sortable", "FALSE"),
                        new A("AppendOnly", true),
                        new A("StaticName", SiteCollectionExternalUserColumns.Comment),
                        new A("Name", SiteCollectionExternalUserColumns.Comment)));

            new RemoveListView(Logger).Execute(ctx, SiteCollectionExternalUsersTitle, AllItems);
            new CreateListView(Logger)
                .Execute(
                    ctx,
                    SiteCollectionExternalUsersTitle,
                    AllItems,
                    new[]
                    {
                        "ID",
                        "LinkTitle",
                        SiteCollectionExternalUserColumns.SiteCollectionExternalUserId,
                        SiteCollectionExternalUserColumns.SiteCollectionUrl,
                        SiteCollectionExternalUserColumns.ExternalUserId,
                        SiteCollectionExternalUserColumns.Start,
                        SiteCollectionExternalUserColumns.End,
                        SiteCollectionExternalUserColumns.Comment
                    },
                    v =>
                    {
                        v.DefaultView = true;
                        v.Paged = true;
                        v.ViewQuery =
                            new E("OrderBy",
                                new E("FieldRef",
                                    new A("Name", "ID"),
                                    new A("Ascending", "FALSE"))).ToString();
                    });
        }

        private void CreateSentMailList(ClientContext ctx)
        {
            var lists = ctx.Web.Lists;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            var candidate = lists.SingleOrDefault(l => l.Title == SentMailTitle);
            if (candidate != null)
            {
                Logger.Warning($"List '{SentMailTitle}' already exists");
                return;
            }

            new CreateListFromTemplate(Logger)
                .Execute(ctx, ListTemplateType.GenericList, SentMailUrl, l =>
                {
                    l.OnQuickLaunch = false;
                    l.EnableVersioning = true;
                    l.Title = SentMailTitle;

                    var title = l.Fields.GetByTitle("Title");
                    title.Required = false;
                    title.Update();
                });

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SentMailTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Text"),
                        new A("DisplayName", SentMailColumns.From),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("MaxLength", "255"),
                        new A("StaticName", SentMailColumns.From),
                        new A("Name", SentMailColumns.From)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SentMailTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Text"),
                        new A("DisplayName", SentMailColumns.To),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("MaxLength", "255"),
                        new A("StaticName", SentMailColumns.To),
                        new A("Name", SentMailColumns.To)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SentMailTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Text"),
                        new A("DisplayName", SentMailColumns.Subject),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("MaxLength", "255"),
                        new A("StaticName", SentMailColumns.Subject),
                        new A("Name", SentMailColumns.Subject)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SentMailTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Note"),
                        new A("DisplayName", SentMailColumns.Body),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("NumLines", "6"),
                        new A("RichText", "TRUE"),
                        new A("RichTextMode", "FullHtml"),
                        new A("IsolateStyles", "TRUE"),
                        new A("Sortable", "FALSE"),
                        new A("AppendOnly", false),
                        new A("StaticName", SentMailColumns.Body),
                        new A("Name", SentMailColumns.Body)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SentMailTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Number"),
                        new A("DisplayName", SentMailColumns.SentMailType),
                        new A("Required", "TRUE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("Decimals", "0"),
                        new A("MaxLength", "255"),
                        new A("StaticName", SentMailColumns.SentMailType),
                        new A("Name", SentMailColumns.SentMailType)));

            new CreateColumnOnList(Logger)
                .Execute(
                    ctx,
                    SentMailTitle,
                    new E("Field",
                        new A("ID", Guid.NewGuid()),
                        new A("Type", "Note"),
                        new A("DisplayName", SentMailColumns.Comment),
                        new A("Required", "FALSE"),
                        new A("EnforceUniqueValues", "FALSE"),
                        new A("Indexed", "FALSE"),
                        new A("NumLines", "6"),
                        new A("RichText", "FALSE"),
                        new A("Sortable", "FALSE"),
                        new A("AppendOnly", true),
                        new A("StaticName", SentMailColumns.Comment),
                        new A("Name", SentMailColumns.Comment)));

            new RemoveListView(Logger).Execute(ctx, SentMailTitle, AllItems);
            new CreateListView(Logger)
                .Execute(
                    ctx, 
                    SentMailTitle,
                    AllItems, 
                    new[] 
                    {
                        "ID",
                        "LinkTitle",
                        SentMailColumns.From,
                        SentMailColumns.To,
                        SentMailColumns.Subject,
                        SentMailColumns.SentMailType,
                        "Created"
                    },
                    v =>
                    {
                        v.DefaultView = true;
                        v.Paged = true;
                        v.ViewQuery =
                            new E("OrderBy",
                                new E("FieldRef",
                                    new A("Name", "ID"),
                                    new A("Ascending", "FALSE"))).ToString();
                    });
        }

        private void DisableMinimumDownloadStrategy(ClientContext ctx)
        {
            new EnsureFeatureState(Logger).Execute(ctx, EnsureFeatureState.Feature.MinimalDownloadStrategy, FeatureDefinitionScope.Web, DesiredFeatureState.Deactivated);
        }

        private E CreateContentEditorWebPart(string title, string contentLink)
        {
            const string webPart =
              @"<?xml version=""1.0"" encoding=""utf-8""?>
                <WebPart xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://schemas.microsoft.com/WebPart/v2"">
                    <Title>TITLE</Title>
                    <FrameType>Default</FrameType>
                    <Description>Allows authors to enter rich text content.</Description>
                    <IsIncluded>true</IsIncluded>
                    <ZoneID>Body</ZoneID>
                    <PartOrder>1</PartOrder>
                    <FrameState>Normal</FrameState>
                    <Height />
                    <Width />
                    <AllowRemove>true</AllowRemove>
                    <AllowZoneChange>true</AllowZoneChange>
                    <AllowMinimize>true</AllowMinimize>
                    <AllowConnect>true</AllowConnect>
                    <AllowEdit>true</AllowEdit>
                    <AllowHide>true</AllowHide>
                    <IsVisible>true</IsVisible>
                    <DetailLink />
                    <HelpLink />
                    <HelpMode>Modeless</HelpMode>
                    <Dir>Default</Dir>
                    <PartImageSmall />
                    <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
                    <PartImageLarge>/_layouts/15/images/mscontl.gif</PartImageLarge>
                    <IsIncludedFilter />
                    <Assembly>Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
                    <TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>
                    <ContentLink xmlns=""http://schemas.microsoft.com/WebPart/v2/ContentEditor"">CONTENTLINK</ContentLink>
                    <Content xmlns=""http://schemas.microsoft.com/WebPart/v2/ContentEditor"" />
                    <PartStorage xmlns=""http://schemas.microsoft.com/WebPart/v2/ContentEditor"" />
                </WebPart>";
            return E.Parse(webPart.Replace("TITLE", title).Replace("CONTENTLINK", contentLink));
        } 

        private void SetupPages(ClientContext ctx)
        {            
            const string sitePages = "Site Pages";

            new CreateWikiPage(Logger).Execute(ctx, sitePages, StartAspx, WikiPageTemplate.TwoColumnsHeader);
            new SetWelcomePage(Logger).Execute(ctx, sitePages, StartAspx);
            var localPath = new Uri(ctx.Url).LocalPath;
            new[]
            {
                CreateContentEditorWebPart("Loader", $"{localPath}/SiteAssets/App/Loader.html"),
                CreateContentEditorWebPart(
                    "Site collection external users overview", 
                    $"{localPath}/SiteAssets/App/SiteCollectionExternalUsersOverview.html")
            }
            .ToList()
            .ForEach(webPart => new AddWebPartToWikiPage(Logger).Execute(ctx, sitePages, StartAspx, webPart, 0, 0));

            new[] { "Home.aspx", "How%20To%20Use%20This%20Library.aspx" }
               .ToList()
               .ForEach(page => new RemoveFileFromLibrary(Logger).Execute(ctx, sitePages, page));

            new CreateWikiPage(Logger).Execute(ctx, sitePages, SiteCollectionExternalUserGuideAspx, WikiPageTemplate.TwoColumnsHeader);
            new[]
            {
                CreateContentEditorWebPart("Loader", $"{localPath}/SiteAssets/App/Loader.html"),
                CreateContentEditorWebPart(
                    "Site collection external user guide", 
                    $"{localPath}/SiteAssets/App/SiteCollectionExternalUserGuide.html")
            }
            .ToList()
            .ForEach(webPart => new AddWebPartToWikiPage(Logger).Execute(ctx, sitePages, SiteCollectionExternalUserGuideAspx, webPart, 0, 0));
        }

        private void ResetQuickLaunch(ClientContext ctx)
        {
            new[] {
                "Home", "Notebook", "Documents", "Recent", "Site Contents" }
                .ToList()
                .ForEach(title =>
                    new RemoveNavigationNode(Logger)
                        .Execute(ctx, RemoveNavigationNode.Navigation.QuickLaunch, title));

            var quickLaunch = ctx.Web.Navigation.QuickLaunch;
            ctx.Load(quickLaunch);
            ctx.ExecuteQuery();
            var localPath = new Uri(ctx.Url).LocalPath;

            new[] { Tuple.Create("Start", StartAspx), Tuple.Create("External user guide", SiteCollectionExternalUserGuideAspx) }
                .Reverse()
                .ToList()
                .ForEach(kvp =>
                {
                    var c = quickLaunch.SingleOrDefault(n => n.Title == kvp.Item1);
                    if (c == null)
                    {
                        quickLaunch.Add(new NavigationNodeCreationInformation
                        {
                            Title = kvp.Item1,
                            Url = $"{localPath}/SitePages/{kvp.Item2}"
                        });
                        ctx.Load(quickLaunch);
                        ctx.ExecuteQuery();
                    }
                });
        }

        public void ResetTopNavigationBar(ClientContext ctx)
        {
            new RemoveNavigationNode(Logger).Execute(ctx, RemoveNavigationNode.Navigation.TopNavigationBar, "Home");
            var topNavigationBar = ctx.Web.Navigation.TopNavigationBar;
            ctx.Load(topNavigationBar);
            ctx.ExecuteQuery();
            var localPath = new Uri(ctx.Url).LocalPath;

            new[] { Tuple.Create("Start", StartAspx) }
                .Reverse()
                .ToList()
                .ForEach(kvp =>
                {
                    var c = topNavigationBar.SingleOrDefault(n => n.Title == kvp.Item1);
                    if (c == null)
                    {
                        topNavigationBar.Add(new NavigationNodeCreationInformation
                        {
                            Title = kvp.Item1,
                            Url = $"{localPath}/SitePages/{kvp.Item2}"
                        });
                        ctx.Load(topNavigationBar);
                        ctx.ExecuteQuery();
                    }
                });
        }

        public void Execute(ClientContext ctx)
        {
            Logger.Verbose($"About to execute {nameof(CreateExternalSharingCenterWeb)} at '{ctx.Url}'");
            CreateExternalUsersList(ctx);
            CreateSiteCollectionExternalUsersList(ctx);
            CreateSentMailList(ctx);
            DisableMinimumDownloadStrategy(ctx);
            SetupPages(ctx);
            ResetQuickLaunch(ctx);
            ResetTopNavigationBar(ctx);
        }
    }
}
