using System;
using System.Xml.Linq;
using System.Collections.Generic;

namespace Bugfree.Spo.ExternalSharingCenter
{
    public class Database
    {
        // conceptually, the relational data model looks like
        //
        //   SiteCollections 1---*  SiteCollectionExternalUsers *---1 ExternalUsers
        //
        // although site collections aren't stored in a separate table. Instead
        // the site collection is a Url field inside SiteCollectionExternalUsers, 
        // leading to a bit of controlled redudancy.
        public List<SiteCollection> SiteCollections { get; set; }
        public List<SiteCollectionExternalUser> SiteCollectionExternalUsers { get; set; }
        public List<ExternalUser> ExternalUsers { get; set; }
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

    // historical mail as stored within the Sent mail list
    public class SentMail
    {
        public int Id { get; set; }
        public string To { get; set; }
        public SentMailType Type { get; set; }
        public DateTime Created { get; set; }
    }

    // representation of expiration for furher processing
    public class Expiration
    {
        public SiteCollection SiteCollection { get; set; }
        public SharePointExternalUser SharePointExternalUser { get; set; }
        public DateTime ExpirationDate { get; set; }
    }

    // representation of expiration warning for further processing
    public class ExpirationWarning
    {
        public SiteCollection SiteCollection { get; set; }
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

    // historical user as stored within the External user list
    public class ExternalUser
    {
        public int Id { get; set; }
        public Guid ExternaluserId { get; set; }
        public string Mail { get; set; }
        public string Comment { get; set; }
    }

    // site collection as reported by SharePoint tenant
    public class SiteCollection
    {
        public Uri Url { get; set; }
        public string Title { get; set; }
        public string FallbackOwnerMail { get; set; }
        public List<SharePointExternalUser> ExternalUsers { get; set; }
    }

    // historical sharing between Site collections and external users entities
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
