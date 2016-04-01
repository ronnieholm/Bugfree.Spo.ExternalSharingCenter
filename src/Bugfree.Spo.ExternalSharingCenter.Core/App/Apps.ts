"use strict";

/// <reference path="jquery.d.ts" />
/// <reference path="angular.d.ts" />

module Bugfree.Spo.ExternalSharingCenter.ViewModels {
    export interface ISiteCollectionExternalUserGuide extends ng.IScope {
        siteCollections: Services.GetSiteCollections.SiteCollection[];
        externalUsers: Services.GetExternalUsers.ExternalUser[];

        validSteps: boolean[];
        currentStep: number;
        siteCollectionUrl: string;
        externalUserId: string;
        mail: string;
        start: Date;
        end: Date;
    }
}

module Bugfree.Spo.ExternalSharingCenter {
    export class Helpers {
        public static generateGuid() {
            var d = new Date().getTime();
            var uuid = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {
                var r = (d + Math.random() * 16) % 16 | 0;
                d = Math.floor(d / 16);
                return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
            });
            return uuid;
        }

        public static getLocationOrigin() {
            // only Chrome supports window.location.origin so we hand-craft a common implementation
            const l = window.location;
            return l.protocol + "//" + l.hostname + (l.port !== "" ? ":" + l.port : "");
        }
    }
}

module Bugfree.Spo.ExternalSharingCenter.Controllers {
    export class GuideController {
        private siteCollectionExternalUsers: Services.GetSiteCollectionExternalUsers.SiteCollectionExternalUser[];

        constructor(private vm: Bugfree.Spo.ExternalSharingCenter.ViewModels.ISiteCollectionExternalUserGuide, private $sce: ng.ISCEService, private $q: ng.IQService, private $location: ng.ILocationService, private getSiteCollections: Services.GetSiteCollections.Query, private getSiteCollectionExternalUsers: Services.GetSiteCollectionExternalUsers.Query, private getExternalUsers: Services.GetExternalUsers.Query, private upsertSiteCollectionExternalUser: Services.UpsertSiteCollectionExternalUser.Command, private createExternalUser: Services.CreateExternalUser.Command) {
            vm.validSteps = [];
            const siteCollectionsPromise = getSiteCollections.execute().then(r => {
                switch (r.state) {
                    case Services.GetSiteCollections.GetSiteCollectionState.Error:
                        this.log("Unspecified client error");
                        break;
                    case Services.GetSiteCollections.GetSiteCollectionState.Ok:
                        vm.siteCollections = r.siteCollections;
                        break;
                }
            }, _ => this.log("Unable to retrieve site collections"));

            const siteCollectionExternalUsersPromise = getSiteCollectionExternalUsers.execute().then(r => {
                switch (r.state) {
                    case Services.GetSiteCollectionExternalUsers.GetSiteCollectionExternalUsersState.Error:
                        this.log("Unspecified client error");
                        break;
                    case Services.GetSiteCollectionExternalUsers.GetSiteCollectionExternalUsersState.Ok:
                        this.siteCollectionExternalUsers = r.siteCollectionExternalUsers;
                        break;
                }
            }, _ => this.log("Unable to retrieve site collection external users"));

            const getExternalUsersPromise = getExternalUsers.execute().then(r => {
                switch (r.state) {
                    case Services.GetExternalUsers.GetExternalUsersState.Error:
                        this.log("Unspecified client error");
                        break;
                    case Services.GetExternalUsers.GetExternalUsersState.Ok:
                        vm.externalUsers = r.externalUsers;
                        break;
                }
            }, _ => this.log("Unable to retrieve external users"));

            $q.all([siteCollectionsPromise, getExternalUsersPromise, siteCollectionExternalUsersPromise]).then(r => {
                // example of query parameters: #?siteCollectionUrl=<url>&externalUserId=<guid>
                const params = $location.search();
                const siteCollectionUrl = params.siteCollectionUrl;
                const externalUserId = params.externalUserId;

                if (siteCollectionUrl !== undefined && externalUserId === undefined) {
                    const sc = this.vm.siteCollections.filter(s => s.url === siteCollectionUrl);

                    if (sc.length === 1) {
                        this.vm.siteCollectionUrl = siteCollectionUrl;
                        vm.validSteps[1] = true;
                        vm.validSteps[2] = true;
                        vm.validSteps[3] = true;
                        vm.currentStep = 3;
                    }
                } else if (siteCollectionUrl !== undefined && externalUserId !== undefined) {
                    const sc = this.vm.siteCollections.filter(s => s.url === siteCollectionUrl);
                    const sceu =
                        this.siteCollectionExternalUsers
                            .filter(s =>
                                s.siteCollectionUrl === siteCollectionUrl &&
                                s.externalUserId === externalUserId);

                    if (sceu.length === 1) {
                        this.vm.siteCollectionUrl = sc[0].url;
                        this.vm.externalUserId = sceu[0].externalUserId;
                        this.vm.start = new Date(sceu[0].start.toString());
                        this.vm.end = new Date(sceu[0].end.toString());
                        vm.validSteps[1] = true;
                        vm.validSteps[2] = true;
                        vm.validSteps[3] = true;
                        vm.currentStep = 4;
                    } else {
                        vm.currentStep = 1;
                    }
                } else {
                    vm.currentStep = 1;
                }
            });

            vm.$watchGroup(["siteCollectionUrl", "externalUserId", "mail", "start", "end"], () => this.updateValidity());
        }

        private updateValidity() {
            const vm = this.vm;
            vm.validSteps[2] = vm.siteCollectionUrl !== undefined;
            if (vm.externalUserId === undefined && vm.mail === undefined) {
                vm.validSteps[3] = false;
            } else if (vm.externalUserId !== undefined && vm.mail === undefined) {
                this.vm.validSteps[3] = true;
            } else if (vm.externalUserId === undefined && vm.mail !== undefined) {
                vm.mail = vm.mail.toLowerCase();
                vm.validSteps[3] = this.vm.externalUsers.filter(u => u.mail === vm.mail).length === 0;
            } else {
                vm.validSteps[3] = false;
            }
            vm.validSteps[4] = vm.start !== undefined && vm.end !== undefined && vm.start < vm.end;
        }

        private log(s) {
            console.log(s);
        }

        public getMail() {
            const vm = this.vm;
            const externalUserId = vm.externalUserId;
            return vm.externalUserId !== undefined
                ? vm.externalUsers.filter(s => s.externalUserId === externalUserId)[0].mail
                : vm.mail;
        }

        public getSiteCollection() {
            const vm = this.vm;
            const siteCollectionUrl = vm.siteCollectionUrl;
            return vm.siteCollectionUrl !== undefined
                ? vm.siteCollections.filter(s => s.url === siteCollectionUrl)[0]
                : "";
        }

        public stepToCssStyle(step: number) {
            return step === this.vm.currentStep ? { "display": "normal" } : { "display": "none" };
        }

        public getCandidateSiteCollectionExternalUser(siteCollectionUrl: string, externalUserId: string) {
            return this.siteCollectionExternalUsers.filter(sceu =>
                sceu.siteCollectionUrl === this.vm.siteCollectionUrl &&
                sceu.externalUserId === this.vm.externalUserId);
        }

        public enterStep(step: number) {
            const vm = this.vm;
            vm.currentStep = step;

            if (step === 4) {
                const candidate = this.getCandidateSiteCollectionExternalUser(vm.siteCollectionUrl, vm.externalUserId);
                if (candidate.length === 1) {
                    vm.start = new Date(candidate[0].start.toString());
                    vm.end = new Date(candidate[0].end.toString());
                }
            }
        }

        private ensureEndOfDayTimeComponent(d: Date) {
            return d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0
                ? new Date(d.getTime() + ((60 * 60 * 24 * 1000) - 1000))
                : d;
        }

        public save() {
            const vm = this.vm;

            if (vm.mail !== undefined) {
                // new sharing because external user doesn't exist
                const externalUser = {
                    externalUserId: Helpers.generateGuid(),
                    mail: vm.mail.toLowerCase()
                };

                this.createExternalUser.execute(externalUser)
                    .then(response => externalUser.mail)
                    .then(mail => this.getExternalUsers.execute().then(r => r.externalUsers.filter(s => s.mail === mail)[0].externalUserId))
                    .then(externalUserId => {
                        const siteCollectionExternalUser = {
                            etag: "0",
                            uri: "",
                            siteCollectionExternalUserId: Helpers.generateGuid(),
                            siteCollectionUrl: vm.siteCollectionUrl,
                            externalUserId: externalUserId,
                            start: vm.start,
                            end: this.ensureEndOfDayTimeComponent(vm.end)
                        };

                        this.upsertSiteCollectionExternalUser.execute(siteCollectionExternalUser).then(r => {
                            if (r.state === Services.UpsertSiteCollectionExternalUser.UpsertSiteCollectionSharingState.Ok) {
                                this.log("Site collection external user successfully created");
                                window.location.replace("Start.aspx");
                            } else {
                                this.log("Unable to create Site collection external user");
                            }
                        });
                    }, _ => this.log("Unable to create Site collection external user"));
            } else {
                const candidate = this.getCandidateSiteCollectionExternalUser(vm.siteCollectionUrl, vm.externalUserId);
                const siteCollectionExternalUser = {
                    etag: candidate.length === 0 ? "0" : candidate[0].etag,
                    uri: candidate.length === 0 ? "" : candidate[0].uri,
                    siteCollectionExternalUserId: candidate.length === 0 ? Helpers.generateGuid() : candidate[0].siteCollectionExternalUserId,
                    siteCollectionUrl: vm.siteCollectionUrl,
                    externalUserId: vm.externalUserId,
                    start: vm.start,
                    end: this.ensureEndOfDayTimeComponent(vm.end)
                };
                this.upsertSiteCollectionExternalUser.execute(siteCollectionExternalUser).then(r => {
                    if (r.state === Services.UpsertSiteCollectionExternalUser.UpsertSiteCollectionSharingState.Ok) {
                        this.log("Site collection external user successfully created");
                        window.location.replace("Start.aspx");
                    } else {
                        this.log("Unable to create Site collection external user");
                    }
                }, _ => this.log("Unable to create Site collection external user"));
            }
        }
    }
}

module Bugfree.Spo.ExternalSharingCenter.ViewModels.SiteCollectionExternalUsers {
    export class ExternalUserBySiteCollection {
        public url: string;
        public start: Date;
        public end: Date;
        public mail: string;
    }

    export interface SiteCollectionExternalUsers extends ng.IScope {
        overview: ExternalUserBySiteCollection[];
    }
}

module Bugfree.Spo.ExternalSharingCenter.Controllers {
    export class OverviewController {
        constructor(private vm: ViewModels.SiteCollectionExternalUsers.SiteCollectionExternalUsers, private $sce: ng.ISCEService, private $q: ng.IQService, private getSiteCollections: Services.GetSiteCollections.Query, private getSiteCollectionExternalUsers: Services.GetSiteCollectionExternalUsers.Query, private getExternalUsers: Services.GetExternalUsers.Query) {
            let siteCollections: Services.GetSiteCollections.SiteCollection[];
            let siteCollectionExternalUsers: Services.GetSiteCollectionExternalUsers.SiteCollectionExternalUser[];
            let externalUsers: Services.GetExternalUsers.ExternalUser[];

            const siteCollectionComparer = (a: Services.GetSiteCollections.SiteCollection, b: Services.GetSiteCollections.SiteCollection) => {
                if (a.title < b.title) {
                    return -1;
                } else if (a.title > b.title) {
                    return 1;
                } else {
                    return 0;
                };
            };

            const siteCollectionsPromise = getSiteCollections.execute().then(r => {
                switch (r.state) {
                    case Services.GetSiteCollections.GetSiteCollectionState.Error:
                        this.log("Unspecified client error");
                        break;
                    case Services.GetSiteCollections.GetSiteCollectionState.Ok:
                        // requires ordering here rather than in view itself to make
                        // site collection and respective external users match up in 
                        // output (external user rows have empty site collection url).
                        siteCollections = r.siteCollections.sort(siteCollectionComparer);
                        break;
                }
            }, _ => this.log("Unable to retrieve site collections"));

            const siteCollectionExternalUsersPromise = getSiteCollectionExternalUsers.execute().then(r => {
                switch (r.state) {
                    case Services.GetSiteCollectionExternalUsers.GetSiteCollectionExternalUsersState.Error:
                        this.log("Unspecified client error");
                        break;
                    case Services.GetSiteCollectionExternalUsers.GetSiteCollectionExternalUsersState.Ok:
                        siteCollectionExternalUsers = r.siteCollectionExternalUsers;
                        break;
                }
            }, _ => this.log("Unable to retrieve Site collection external users"));

            const getExternalUsersPromise = getExternalUsers.execute().then(r => {
                switch (r.state) {
                    case Services.GetExternalUsers.GetExternalUsersState.Error:
                        this.log("Unspecified client error");
                        break;
                    case Services.GetExternalUsers.GetExternalUsersState.Ok:
                        externalUsers = r.externalUsers;
                        break;
                }
            }, _ => this.log("Unable to retrieve external users"));

            $q.all([siteCollectionsPromise, getExternalUsersPromise, siteCollectionExternalUsersPromise]).then(r => {
                vm.overview = this.generateOverview(siteCollections, siteCollectionExternalUsers, externalUsers);
            });
        }

        private generateOverview(siteCollections: Services.GetSiteCollections.SiteCollection[], siteCollectionExternalUsers: Services.GetSiteCollectionExternalUsers.SiteCollectionExternalUser[], externalUsers: Services.GetExternalUsers.ExternalUser[]) {
            const result = [];
            siteCollections.forEach((sc, _, __) => {
                result.push({
                    siteCollectionUrl: sc.url,
                    title: sc.title
                });

                const sceu = siteCollectionExternalUsers.filter(s => s.siteCollectionUrl === sc.url);
                if (sceu.length > 0) {
                    let s;
                    for (s of sceu) {
                        const users = externalUsers.filter(eu => eu.externalUserId === s.externalUserId)[0];
                        result.push({
                            siteCollectionUrl: s.siteCollectionUrl,
                            externalUserId: s.externalUserId,
                            url: "",
                            title: "",
                            start: s.start,
                            end: s.end,
                            mail: users.mail
                        });
                    }
                }
            });
            return result;
        }

        private log(s) {
            console.log(s);
        }
    }
}

module Bugfree.Spo.ExternalSharingCenter.Services.GetSiteCollectionExternalUsers {
    export enum GetSiteCollectionExternalUsersState {
        None = 0,
        Ok,
        Error
    }

    export class SiteCollectionExternalUser {
        public etag: string;
        public id: number;
        public uri: string;
        public siteCollectionExternalUserId: string;
        public siteCollectionUrl: string;
        public externalUserId: string;
        public start: Date;
        public end: Date;
        public comment: string;
    }

    export class GetSiteCollectionExternalUsersResponse {
        public state: GetSiteCollectionExternalUsersState;
        public siteCollectionExternalUsers: SiteCollectionExternalUser[];
    }

    export class Query {
        constructor(private $http: ng.IHttpService, private $q: ng.IQService) { }

        public execute() {
            const deferred = this.$q.defer<GetSiteCollectionExternalUsersResponse>();
            this.$http.defaults.headers.common["Accept"] = "application/json;odata=verbose";
            this.$http
                .get("../_api/web/lists/getbytitle('site collection external users')/items?$top=5000")
                .then(response => {
                    const data: any = response.data;
                    const results: any[] = data.d.results;
                    let siteCollectionExternalUsers = results.map((s, _, __) => {
                        return {
                            etag: s.__metadata.etag,
                            uri: s.__metadata.uri,
                            id: s.Id,
                            siteCollectionExternalUserId: s.SiteCollectionExternalUserId,
                            siteCollectionUrl: s.SiteCollectionUrl,
                            externalUserId: s.ExternalUserId,
                            start: s.Start,
                            end: s.End,
                            comment: s.Comment
                        };
                    });

                    // for generating demo screen
                    //siteCollectionExternalUsers = [
                    //    { etag: 0, uri: "uri1", id: 1, siteCollectionExternalUserId: "guid10", siteCollectionUrl: "http://accounting", externalUserId: "guid1", start: new Date(2016, 3, 1), end: new Date(2016, 4, 30), comment: "" },
                    //    { etag: 0, uri: "uri2", id: 2, siteCollectionExternalUserId: "guid11", siteCollectionUrl: "http://accounting", externalUserId: "guid2", start: new Date(2016, 3, 1), end: new Date(2016, 5, 15), comment: "" },
                    //    { etag: 0, uri: "uri3", id: 3, siteCollectionExternalUserId: "guid12", siteCollectionUrl: "http://sales", externalUserId: "guid3", start: new Date(2016, 1, 1), end: new Date(2016, 4, 30), comment: "" }
                    //];                

                    deferred.resolve({ state: GetSiteCollectionExternalUsersState.Ok, siteCollectionExternalUsers: siteCollectionExternalUsers });
                },
                _ => deferred.reject({ state: GetSiteCollectionExternalUsersState.Error, siteCollectionSharing: [] }));
            return deferred.promise;
        }
    }
}

module Bugfree.Spo.ExternalSharingCenter.Services.CreateExternalUser {
    export enum CreateExternalUserState {
        None = 0,
        Ok,
        Error
    }

    export class ExternalUser {
        public externalUserId: string;
        public mail: string;
    }

    export class CreateToSharingResponse {
        public state: CreateExternalUserState;
    }

    export class Command {
        constructor(private $http: ng.IHttpService, private $q: ng.IQService) { }

        public execute(e: ExternalUser) {
            const deferred = this.$q.defer<CreateToSharingResponse>();
            const body = {
                "Title": null,
                "ExternalUserId": e.externalUserId,
                "Mail": e.mail
            };

            this.$http.defaults.headers.common["Accept"] = "application/json;odata=verbose";
            this.$http.defaults.headers.common["contentType"] = "application/json;odata=verbose";
            this.$http.defaults.headers.common["X-RequestDigest"] = (<any>document.getElementById("__REQUESTDIGEST")).value;
            this.$http
                .post("../_api/web/lists/getbytitle('external users')/items", JSON.stringify(body))
                .then(response => deferred.resolve({ state: CreateExternalUserState.Ok }),
                      _ => deferred.reject({ state: CreateExternalUserState.Error }));
            return deferred.promise;
        }
    }
}

module Bugfree.Spo.ExternalSharingCenter.Services.UpsertSiteCollectionExternalUser {
    export enum UpsertSiteCollectionSharingState {
        None = 0,
        Ok,
        Error
    }

    export class SiteCollectionExternalUser {
        public etag: string;
        public uri: string;
        public siteCollectionExternalUserId: string;
        public siteCollectionUrl: string;
        public externalUserId: string;
        public start: Date;
        public end: Date;
    }

    export class UpsertSiteCollectionSharingResponse {
        public state: UpsertSiteCollectionSharingState;
    }

    export class Command {
        constructor(private $http: ng.IHttpService, private $q: ng.IQService) { }

        public execute(u: SiteCollectionExternalUser) {
            const deferred = this.$q.defer<UpsertSiteCollectionSharingResponse>();
            const body = {
                "Title": null,
                "SiteCollectionExternalUserId": u.siteCollectionExternalUserId,
                "SiteCollectionUrl": u.siteCollectionUrl,
                "ExternalUserId": u.externalUserId,
                "Start": u.start,
                "End": u.end
            };

            let c = this.$http.defaults.headers.common;
            c["Accept"] = "application/json;odata=verbose";
            c["contentType"] = "application/json;odata=verbose";
            c["X-RequestDigest"] = (<any>document.getElementById("__REQUESTDIGEST")).value;

            if (u.etag === "0") {
                this.$http
                    .post("../_api/web/lists/getbytitle('site collection external users')/items", JSON.stringify(body))
                    .then(response =>
                        deferred.resolve({ state: UpsertSiteCollectionSharingState.Ok }),
                        _ => deferred.reject({ state: UpsertSiteCollectionSharingState.Error }));
            } else {
                c["X-HTTP-Method"] = "MERGE";
                c["If-Match"] = u.etag;
                this.$http
                    .post(u.uri, JSON.stringify(body))
                    .then(response =>
                        deferred.resolve({ state: UpsertSiteCollectionSharingState.Ok }),
                    _ => deferred.reject({ state: UpsertSiteCollectionSharingState.Error }));
            }

            return deferred.promise;
        }
    }
}

module Bugfree.Spo.ExternalSharingCenter.Services.GetExternalUsers {
    export enum GetExternalUsersState {
        None = 0,
        Ok,
        Error
    }

    export class ExternalUser {
        public id: number;
        public externalUserId: string;
        public mail: string;
        public comment: string;
    }

    export class GetExternalUserResponse {
        public state: GetExternalUsersState;
        public externalUsers: ExternalUser[];
    }

    export class Query {
        constructor(private $http: ng.IHttpService, private $q: ng.IQService) { }

        public execute() {
            const deferred = this.$q.defer<GetExternalUserResponse>();
            this.$http.defaults.headers.common["Accept"] = "application/json;odata=verbose";
            this.$http
                .get("../_api/web/lists/getbytitle('external users')/items?$top=5000&$select=Id,ExternalUserId,Mail,Comment")
                .then(response => {
                    const data: any = response.data;
                    const results: any[] = data.d.results;
                    let externalUsers = results.map((s, _, __) => {
                        return {
                            id: s.Id,
                            externalUserId: s.ExternalUserId,
                            mail: s.Mail,
                            comment: s.Comment
                        };
                    });

                    // for generating demo screen
                    //externalUsers = [
                    //    { id: 1, externalUserId: "guid1", mail: "john@acmecorp.com", comment: "" },
                    //    { id: 2, externalUserId: "guid2", mail: "jane@acmecorp.com", comment: "" },
                    //    { id: 3, externalUserId: "guid3", mail: "julian@contoso.com", comment: "" },
                    //];

                    deferred.resolve({ state: GetExternalUsersState.Ok, externalUsers: externalUsers });
                },
                _ => deferred.reject({ state: GetExternalUsersState.Error, sharings: [] }));

            return deferred.promise;
        }
    }
}

module Bugfree.Spo.ExternalSharingCenter.Services.GetSiteCollections {
    export enum GetSiteCollectionState {
        None = 0,
        Ok,
        Error
    }

    export class GetSiteCollectionResponse {
        state: GetSiteCollectionState;
        siteCollections: any[];
    }

    export interface SiteCollection {
        url: string;
        title: string;
    }

    export class Query {
        constructor(private $http: ng.IHttpService, private $q: ng.IQService) { }

        private getSearchEngineResultsRecursive(url: string, startRow: number, siteCollections: SiteCollection[], deferred: ng.IDeferred<any[]>) {
            const batchSize = 500;
            this.$http.get(`${url}&rowlimit=${batchSize}&startRow=${startRow}`)
                .then(response => {
                    const relevantResults: any = (<any>response.data).PrimaryQueryResult.RelevantResults;
                    const rows: any[] = relevantResults.Table.Rows;

                    for (var i = 0; i < rows.length; i++) {
                        const row = rows[i];
                        const cells = row.Cells;

                        var title = "";
                        var url = "";
                        for (var j = 0; j < cells.length; j++) {
                            var key = cells[j]["Key"];
                            var value = cells[j]["Value"];

                            if (key === "Title") {
                                title = value;
                            }

                            if (key === "OriginalPath") {
                                url = value;
                            }
                        }

                        siteCollections.push({
                            title: title,
                            url: url
                        });
                    }

                    const totalRows = relevantResults.TotalRows;
                    if (totalRows > startRow + batchSize) {
                        this.getSearchEngineResultsRecursive(url, (startRow + batchSize), siteCollections, deferred);
                    } else {
                        deferred.resolve(siteCollections);
                    }
                }, _ => {
                    deferred.reject();
                });

            return deferred.promise;
        }

        execute() {
            const deferred = this.$q.defer<GetSiteCollectionResponse>();

            // primitive switching between site collection retrieval strategies
            if (true) {
                this.$http.defaults.headers.common["Accept"] = "application/json;odata=verbose";
                this.$http
                    .get("../_api/web/lists/getbytitle('site collections')/items?$top=5000&$select=Id,SiteCollectionTitle,SiteCollectionUrl")
                    .then(response => {
                        const data: any = response.data;
                        const results: any[] = data.d.results;
                        let siteCollections = results.map((s, _, __) => {
                            return {
                                id: s.Id,
                                title: s.SiteCollectionTitle,
                                url: s.SiteCollectionUrl
                            };
                        });
                        deferred.resolve({ state: GetSiteCollectionState.Ok, siteCollections: siteCollections });
                    },
                    _ => deferred.reject({ state: GetSiteCollectionState.Error, siteCollections: [] }));
            } else {
                const baseUrl = Helpers.getLocationOrigin();
                const queryText = `ContentClass:STS_Site Path:${baseUrl}/sites Path:${baseUrl}/teams`;
                const selectProperties = ["Title", "OriginalPath"];

                //const deferred = this.$q.defer<GetSiteCollectionResponse>();
                const selectPropertiesAsDelimitedString = ["", ...selectProperties].reduce((acc, c) => acc.length === 0 ? c : acc + "," + c);
                const selectPropertiesAsString =
                    selectProperties.length === 0
                        ? ""
                        : `&selectproperties='${selectPropertiesAsDelimitedString}'`;

                const url = `${baseUrl}/_api/search/query?querytext='${queryText}'${selectPropertiesAsString}`;
                const recursiveDeferred = this.$q.defer<any[]>();
                this.getSearchEngineResultsRecursive(url, 0, [], recursiveDeferred).then(results => {
                    // for generating demo screen
                    //results = [
                    //    { title: "HR", url: "http://hr" },
                    //    { title: "Sales", url: "http://sales" },
                    //    { title: "Accounting", url: "http://accounting" }
                    //];
                    deferred.resolve({ state: GetSiteCollectionState.Ok, siteCollections: results });
                }, _ => {
                    deferred.reject({ state: GetSiteCollectionState.Error, siteCollections: [] });
                });
            }

            return deferred.promise;
        }
    }
}

angular.module("app").service("GetSiteCollections", Bugfree.Spo.ExternalSharingCenter.Services.GetSiteCollections.Query);
angular.module("app").service("GetSiteCollectionExternalUsers", Bugfree.Spo.ExternalSharingCenter.Services.GetSiteCollectionExternalUsers.Query);
angular.module("app").service("GetExternalUsers", Bugfree.Spo.ExternalSharingCenter.Services.GetExternalUsers.Query);
angular.module("app").service("UpsertSiteCollectionExternalUser", Bugfree.Spo.ExternalSharingCenter.Services.UpsertSiteCollectionExternalUser.Command);
angular.module("app").service("CreateExternalUser", Bugfree.Spo.ExternalSharingCenter.Services.CreateExternalUser.Command);
angular.module("app").controller("OverviewController", ["$scope", "$sce", "$q", "GetSiteCollections", "GetSiteCollectionExternalUsers", "GetExternalUsers", Bugfree.Spo.ExternalSharingCenter.Controllers.OverviewController]);
angular.module("app").controller("GuideController", ["$scope", "$sce", "$q", "$location", "GetSiteCollections", "GetSiteCollectionExternalUsers", "GetExternalUsers", "UpsertSiteCollectionExternalUser", "CreateExternalUser", Bugfree.Spo.ExternalSharingCenter.Controllers.GuideController]);
