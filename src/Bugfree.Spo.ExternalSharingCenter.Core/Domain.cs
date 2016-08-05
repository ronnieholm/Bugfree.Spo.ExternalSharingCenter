using System;
using System.Xml.Linq;
using System.Collections.Generic;

namespace Bugfree.Spo.ExternalSharingCenter
{
    public class Database
    {
        // conceptually, the relational data model looks like
        //
        //   SharedSiteCollection 1---*  SiteCollectionExternalUser *---1 ExternalUser
        //
        // with SharedSiteCollections populated by quering the tenant API and the other
        // entities coming from backing lists within the External Sharing Center. In practice,
        // the SharedSiteCollection instance is a Url field inside SiteCollectionExternalUsers.
        // Using the Url field, lookup of the SharedSiteCollection instance can take place.
        public List<SiteCollectionExternalUser> SiteCollectionExternalUsers { get; set; }
        public List<ExternalUser> ExternalUsers { get; set; }
        public List<SharedSiteCollection> SharedSiteCollections { get; set; }
    }

    public enum SentMailType
    {
        None = 0,
        Warning = 1,
        Expiration = 2
    }

    // constructed mail to be sent to a InvitedBy user
    public class Email
    {
        public string From { get; set; }
        public string To { get; set; }
        public string Subject { get; set; }
        public XElement Body { get; set; }
        public SentMailType Type { get; set; }
    }

    // archive of sent mail (useful for diagnostic purposes)
    public class SentMail
    {
        public int Id { get; set; }
        public string To { get; set; }
        public SentMailType Type { get; set; }
        public DateTime Created { get; set; }
    }

    // expiration for further processing
    public class Expiration
    {
        public SharedSiteCollection SiteCollection { get; set; }
        public SharePointExternalUser SharePointExternalUser { get; set; }
        public DateTime ExpirationDate { get; set; }
    }

    // expiration warning for further processing
    public class ExpirationWarning
    {
        public SharedSiteCollection SiteCollection { get; set; }
        public SharePointExternalUser SharePointExternalUser { get; set; }
        public DateTime ExpirationDate { get; set; }
        public TimeSpan TimeUntilExpiration { get; set; }
    }

    // external user as reported by SharePoint tenant
    public class SharePointExternalUser
    {
        public int UserId { get; set; }
        public string DisplayName { get; set; }
        public string AcceptedAs { get; set; }
        public string InvitedAs { get; set; }
        public string InvitedBy { get; set; }
        public DateTime WhenCreated { get; set; }
    }

    // site collection as reported by SharePoint tenant
    public class SharedSiteCollection {
        public Uri Url { get; set; }
        public string Title { get; set; }
        public string FallbackOwnerMail { get; set; }
        public List<SharePointExternalUser> ExternalUsers { get; set; }
    }

    // external user stored within ExternalUsers list
    public class ExternalUser
    {
        public int Id { get; set; }
        public Guid ExternaluserId { get; set; }
        public string Mail { get; set; }
        public string Comment { get; set; }
    }

    // sharing between Site collections and external users entities
    // Think of the properties as properties on the relation itself or the
    // junction table in an n:m relationship.
    public class SiteCollectionExternalUser
    {
        public int Id { get; set; }
        public Guid SiteCollectionExternalUserId { get; set; }
        public Uri SiteCollectionUrl { get; set; }
        public Guid ExternalUserId { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Comment { get; set; }
    }
}
