/// <reference path="../Typings/angularjs/angular.d.ts" />

declare module AngularSP {
    interface SharePoint {
        GetItemTypeForListName(name: string): angular.IPromise<any>
        SanitizeWebUrl(url: string): string
        CreateListItem(listName: string, webUrl: string, item: any, hostUrl?: string): angular.IPromise<any>
        UpdateListItem(itemId: number, listName: string, webUrl: string, item: any, hostUrl?: string): angular.IPromise<any>
        DeleteListItem(itemId: number, listName: string, webUrl?: string, hostUrl?: string): angular.IPromise<any>
        GetGroup(groupName: string, includeMembers?: boolean, webUrl?: string, hostUrl?: string): angular.IPromise<any>
        GetSiteUsers(webUrl?: string, options?: any, hostUrl?: string): angular.IPromise<any>
        GetUserById(userId: number, webUrl?: string, hostUrl?: string): angular.IPromise<any>
        AddUsertoGroup(groupId: number, loginName: string, webUrl?: string, hostUrl?: string): angular.IPromise<any>
        CreateSubSite(options: any, webUrl?: string, hostUrl?: string): angular.IPromise<any>
        GetWebData(webUrl?: string, hostUrl?: string): angular.IPromise<any>
        UpdateWebData(webUrl: string, updateData: any, hostUrl?: string): angular.IPromise<any>
        AddFileToLibrary(listName: string, webUrl: string, fileName: string, file: File): angular.IPromise<any>
    }
    interface ISearch {
        Get(webUrl: string, options: any): angular.IPromise<any>
        Post(webUrl: string, options: any)?: angular.IPromise<any>
    }
    module AngularSPREST {
        interface IAngularSPREST {
            GetUpdateDigest(webUrl: string, hostUrl: string): angular.IPromise<any>    
            GetItemById(itemId: number, listName: string, webUrl?: string, extraParams?: any, hostUrl?: string): angular.IPromise<any>
            GetListItems(listName: string, webUrl?: string, options?: any, hostUrl?: string): angular.IPromise<any>
            GetListItemsByCAML(listName: string, webUrl: string, camlQuery: string, options?: any, hostUrl?: string): angular.IPromise<any>
        }
        module Profile {
            interface IProfile {
                GetCurrentUser(webUrl:string, options:any, hostUrl?:string): angular.IPromise<any>
                GetForUser(webUrl: string, userName: string, options: any, hostUrl?: string): angular.IPromise<any>
            }
        }
    }

    module AngularSPCSOM {
        interface ICSOM {
            GetItemTypeForListName(name: string): string
            SanitizeWebUrl(url: string): string
            CreateListItem(listName: string, webUrl: string, item: any, hostUrl?: string): angular.IPromise<any>
            GetItemById(itemId: number, listName: string, webUrl: string, hostUrl?: string): angular.IPromise<any>
            GetListItems(listName: string, webUrl: string, camlquery: string, hostUrl?: string): angular.IPromise<any>
            
        }
    }
}