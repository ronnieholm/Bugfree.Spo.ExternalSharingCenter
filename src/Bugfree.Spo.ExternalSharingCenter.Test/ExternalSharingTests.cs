﻿using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using Bugfree.Spo.Cqrs.Core;
using Bugfree.Spo.ExternalSharingCenter.Core.Queries;

namespace Bugfree.Spo.ExternalSharingCenter.Test
{
    public class Revoke_access_to_external_user
    {
        private ILogger _logger = new ColoredConsoleLogger();
        private List<SiteCollection> siteCollection = new List<SiteCollection>
        {
            new SiteCollection
            {
                Url = new Uri("http://test/siteCollection"),
                ExternalUsers = new List<SharePointExternalUser>
                {
                    new SharePointExternalUser
                    {
                        InvitedAs = "test@test.com"
                    }
                }
            }
        };

        private List<ExternalUser> emptyExternalUsers = new List<ExternalUser>();
        private List<SiteCollectionExternalUser> emptySiteCollectionExternalUsers = new List<SiteCollectionExternalUser>();

        [Fact]
        public void When_not_in_external_users_list()
        {
            var db = new Database
            {
                SiteCollections = siteCollection,
                ExternalUsers = emptyExternalUsers,
                SiteCollectionExternalUsers = 
                    new List<SiteCollectionExternalUser>
                    {
                        new SiteCollectionExternalUser
                        {
                        }
                    }
            };

            var expirations = new GenerateExpirations(_logger).Execute(db, DateTime.MinValue);
            Assert.Equal(1, expirations.Count());
            Assert.Equal(db.SiteCollections[0].Url, expirations[0].SiteCollection.Url);
            Assert.Equal(db.SiteCollections[0].ExternalUsers[0].InvitedAs, expirations[0].SharePointExternalUser.InvitedAs);
        }

        [Fact]
        public void When_in_external_users_list_but_for_wrong_site_collection()
        {
            var db = new Database
            {
                SiteCollections = siteCollection,
                SiteCollectionExternalUsers = emptySiteCollectionExternalUsers,
                ExternalUsers = new List<ExternalUser>
                {
                    new ExternalUser
                    {
                        Mail = "test@test.com"
                    }
                }
            };

            var evictions = new GenerateExpirations(_logger).Execute(db, DateTime.MinValue);
            Assert.Equal(1, evictions.Count());
            Assert.Equal(db.SiteCollections[0].Url, evictions[0].SiteCollection.Url);
            Assert.Equal(db.SiteCollections[0].ExternalUsers[0].InvitedAs, evictions[0].SharePointExternalUser.InvitedAs);
        }

        [Fact]
        public void Whose_site_collection_external_user_start_date_is_in_the_future()
        {
            var siteCollectionSharings = new List<SiteCollectionExternalUser>
            {
                new SiteCollectionExternalUser
                {
                    ExternalUserId = Guid.NewGuid(),
                    SiteCollectionUrl = new Uri("http://dr.dk"),
                    Start = new DateTime(2100, 1, 1),
                    End = new DateTime(2200, 1, 1)
                }
            };

            var db = new Database
            {
                SiteCollectionExternalUsers = siteCollectionSharings,
                SiteCollections = siteCollection,
                ExternalUsers = new List<ExternalUser>
                {
                    new ExternalUser
                    {
                        Id = 42,
                        Mail = "test@test.com"
                    }
                }
            };

            var evictions = new GenerateExpirations(_logger).Execute(db, new DateTime(2000, 1, 1));
            Assert.Equal(1, evictions.Count());
            Assert.Equal(db.SiteCollections[0].Url, evictions[0].SiteCollection.Url);
            Assert.Equal(db.SiteCollections[0].ExternalUsers[0].InvitedAs, evictions[0].SharePointExternalUser.InvitedAs);
        }

        [Fact]
        public void Whose_site_collection_external_user_end_date_is_in_the_past()
        {
            var db = new Database
            {
                SiteCollectionExternalUsers = new List<SiteCollectionExternalUser>
                {
                    new SiteCollectionExternalUser
                    {
                        ExternalUserId = Guid.NewGuid(),
                        SiteCollectionUrl = new Uri("http://dr.dk"),
                        Start = new DateTime(1900, 1, 1),
                        End = new DateTime(1950, 1, 1)
                    }
                },
                ExternalUsers = new List<ExternalUser>
                {
                    new ExternalUser
                    {
                        Id = 42,
                        Mail = "test@test.com"
                    }
                },
                SiteCollections = siteCollection
            };

            var expirations = new GenerateExpirations(_logger).Execute(db, new DateTime(2000, 1, 1));
            Assert.Equal(1, expirations.Count());
            Assert.Equal(db.SiteCollections[0].Url, expirations[0].SiteCollection.Url);
            Assert.Equal(db.SiteCollections[0].ExternalUsers[0].InvitedAs, expirations[0].SharePointExternalUser.InvitedAs);
        }
    }

    public class Keep_external_user
    {
        private readonly ILogger _logger = new ColoredConsoleLogger();
        private readonly List<SiteCollection> siteCollectionWithExternalUsers = new List<SiteCollection>
        {
            new SiteCollection
            {
                Url = new Uri("http://test/siteCollection"),
                ExternalUsers = new List<SharePointExternalUser>
                {
                    new SharePointExternalUser
                    {
                        InvitedAs = "test@test.com"
                    }
                }
            }
        };

        [Fact]
        public void Whose_site_collection_external_user_start_date_and_end_date_surrounds_current_time()
        {
            var siteCollectionSharings = new List<SiteCollectionExternalUser>
            {
                new SiteCollectionExternalUser
                {
                    ExternalUserId = Guid.Empty,
                    SiteCollectionUrl = new Uri("http://test/siteCollection"),
                    Start = new DateTime(2000, 1, 1),
                    End = new DateTime(2100, 1, 1)
                }
            };

            var db = new Database
            {
                SiteCollections = siteCollectionWithExternalUsers,
                SiteCollectionExternalUsers = siteCollectionSharings,
                ExternalUsers = new List<ExternalUser>
                {
                    new ExternalUser
                    {
                        Id = 42,
                        Mail = "test@test.com"
                    }
                }
            };

            var expiration = new GenerateExpirations(_logger).Execute(db, new DateTime(2050, 1, 1));
            Assert.Equal(0, expiration.Count());
        }
    }

    public class Issue_expiration_warning
    {
        private readonly ILogger _logger = new ColoredConsoleLogger();
        private readonly List<SiteCollection> siteCollections = new List<SiteCollection>
        {
            new SiteCollection
            {               
                //Id = 44,
                Url = new Uri("http://test/siteCollection"),
                ExternalUsers = new List<SharePointExternalUser>
                {
                    new SharePointExternalUser
                    {
                        InvitedAs = "test@test.com"
                    }
                }
            }
        };

        readonly List<ExternalUser> externalUser = new List<ExternalUser>
        {
            new ExternalUser
            {
                Id = 42,
                Mail = "test@test.com"
            }
        };

        [Fact]
        public void In_advance()
        {
            var db = new Database
            {
                SiteCollections = siteCollections,
                ExternalUsers = externalUser,
                SiteCollectionExternalUsers = 
                    new List<SiteCollectionExternalUser>
                    {
                        new SiteCollectionExternalUser
                        {
                            ExternalUserId = Guid.Empty,
                            SiteCollectionUrl = new Uri("http://test/siteCollection"),
                            Start = new DateTime(2015, 11, 16),
                            End = new DateTime(2015, 11, 30, 23, 59, 59)
                        },
                        new SiteCollectionExternalUser
                        {
                            ExternalUserId = Guid.Empty,
                            SiteCollectionUrl = new Uri("http://test/siteCollection"),
                            Start = new DateTime(2015, 11, 16),
                            End = new DateTime(2015, 12, 30, 23, 59, 59)
                        },
                        new SiteCollectionExternalUser
                        {
                            ExternalUserId = Guid.Empty,
                            SiteCollectionUrl = new Uri("http://test/siteCollection"),
                            Start = new DateTime(2015, 11, 16),
                            End = new DateTime(2015, 11, 19)
                        } 
                    }           
            };

            var now = new DateTime(2015, 11, 28);
            var expirationWarnings = new GenerateExpirationWarnings(_logger).Execute(db, now);
            Assert.Equal(db.SiteCollectionExternalUsers.Count, expirationWarnings.Count());

            Assert.Equal(db.SiteCollectionExternalUsers[0].End - now, new TimeSpan(2, 23, 59, 59));
            Assert.Equal(db.SiteCollectionExternalUsers[1].End - now, new TimeSpan(32, 23, 59, 59));
            Assert.Equal(db.SiteCollectionExternalUsers[2].End - now, new TimeSpan(-9, 00, 00, 00));
        }
    }
}
