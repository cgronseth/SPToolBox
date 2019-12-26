import { LogAx } from "../spt.logax";
import { SPView, SPData } from "./spt.sharepoint.entities";
import { Strings } from "../spt.strings";
import { SP } from "./spt.sharepoint";

export enum RestQueryType {
    ODataJSON,
    ArrayBuffer
}

export class SPRest {
    /**
    * Lanza consulta REST y devuelve los resultados en un Promise
    * @param qry Consulta completa: HTTPS://<servidor>/_api/ProjectServer/Projects('2c7d134c- [...]
    * @param reintentos Número de reintentos de 6s. Por defecto 10
    */
    public static restQuery(qry: string, type: RestQueryType, reintentos?: number): Promise<any> {
        let tiempoReintento: number = 30000;
        if (reintentos === undefined) {
            reintentos = 10;
        }

        return new Promise<any>((resolve, reject) => {
            let xhr: XMLHttpRequest = new XMLHttpRequest();
            xhr.open('GET', qry, true);
            if (type === RestQueryType.ODataJSON) {
                xhr.setRequestHeader("Accept", "application/json;odata=nometadata");
            } else if (type == RestQueryType.ArrayBuffer) {
                xhr.responseType = "arraybuffer";
            }

            xhr.onreadystatechange = () => {
                if (xhr.readyState === 4) {

                    //Debug. Random error. Descomentar para habilitar:
                    /*
                    let debugError: boolean = (Math.random() < 0.3); //30% error
                    */

                    if (/*!debugError &&*/ type === RestQueryType.ODataJSON && xhr.status === 200 && xhr.getResponseHeader('content-type').startsWith('application/json')) {
                        let result = JSON.parse(xhr.responseText);

                        let coolOffDelay = setTimeout(() => {
                            clearTimeout(coolOffDelay);
                            resolve(result);
                        }, Math.floor(Math.random() * 100) + 50)

                    } else if (type === RestQueryType.ArrayBuffer && xhr.status === 200 && xhr.getResponseHeader('content-type') === "application/octet-stream") {
                        resolve(xhr.response);
                    } else if (xhr.status === 401 || xhr.status === 403) {
                        reject('RESTQuery - Unauthorized/Forbidden');
                    } else {
                        LogAx.trace('RESTQuery - Status==' + xhr.status.toString() + ': ' + xhr.responseText);
                        tiempoReintento = Math.floor(Math.random() * 1000) + tiempoReintento;

                        // Adapt query if field fails
                        if (xhr.status === 400) {
                            // Example responseText --> {\"odata.error\":{\"code\":\"-1, Microsoft.SharePoint.SPException\",\"message\":{\"lang\":\"en-US\",\"value\":\"The field or property 'AppAuthor' does not exist.\"}}}"
                            let matches = xhr.responseText.match(/property '(.*)' does not exist/);
                            if (matches.length === 2) {
                                // Erase field in query, be it a simple field or in a expand pair --> [",Author/Title", ",Author/EMail", "Author,"]
                                let rgx = new RegExp('(,' + matches[1] + '\/[a-zA-Z_0-9]+)|(' + matches[1] + ',)|(,' + matches[1] + ')', 'g');
                                qry = qry.replace(rgx, "");

                                // Not a connection problem, so restore retries and send off immediatly
                                tiempoReintento = 100;
                                reintentos++;
                                LogAx.trace("RESTQuery - Reintento eliminando el campo problemático '" + matches[1] + "'. Query:" + qry);
                            }
                        }

                        // Recursive retry
                        if (reintentos && --reintentos >= 0) {
                            LogAx.trace('RESTQuery - Reintentos restantes: ' + reintentos.toString());

                            let retryDelay = setTimeout(() => {
                                clearTimeout(retryDelay);
                                SPRest.restQuery(qry, type, reintentos).then((recResult) => {
                                    resolve(recResult);
                                }, (e) => {
                                    reject(e);
                                })
                            }, tiempoReintento);

                        } else {
                            reject('RESTQuery - Reintentos agotados');
                        }
                    }
                }
            };
            xhr.onerror = () => {
                reject('RESTQuery - OnError: ' + xhr.responseText);
            };
            xhr.send();
        });
    }

    /**
     * Sends POST request
     * @param qry 
     * @param requestDigestValue 
     * @param newFolder 
     * @param isDelete 
     * @param fileData 
     * @param listItemEntity 
     * @param itemData 
     */
    public static restPost(qry: string, requestDigestValue?: string, newFolder?: string, isDelete?: boolean, fileData?: ArrayBuffer, listItemEntity?: string, itemData?: SPData[]): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            let jsonBody: any = null;
            let spRequest = new XMLHttpRequest();
            spRequest.open('POST', qry, true);
            //cabeceras indicados en https://pwmather.wordpress.com/2018/05/21/using-rest-in-javascript-to-update-projectonline-project-custom-fields-ppm-pmot-jquery-office365/

            spRequest.setRequestHeader("Accept", "application/json;odata=verbose");
            if (requestDigestValue) {
                spRequest.setRequestHeader("X-RequestDigest", requestDigestValue);
            }
            if (!fileData) {
                spRequest.setRequestHeader("If-Match", "*");
            }
            if (newFolder) {
                jsonBody = JSON.stringify({ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': newFolder });
                spRequest.setRequestHeader("content-type", "application/json;odata=verbose");
            }
            if (isDelete) {
                spRequest.setRequestHeader("X-HTTP-Method", "DELETE");
            }
            if (itemData && listItemEntity) {
                let body: any = { '__metadata': { 'type': listItemEntity } };
                for (let spd of itemData) {
                    let internalName: string = spd.InternalName;
                    if (spd.Type === 7) {   //Lookup
                        internalName += "Id";
                    }
                    body[internalName] = SP.assignPostValue(spd);
                }
                jsonBody = JSON.stringify(body);
                spRequest.setRequestHeader("content-type", "application/json;odata=verbose");
            }

            spRequest.onreadystatechange = () => {
                if (spRequest.readyState === 4) {
                    if (spRequest.status >= 200 && spRequest.status < 300) {
                        if (spRequest.responseText !== '') {
                            resolve(JSON.parse(spRequest.responseText));
                        } else if (spRequest.statusText.toLowerCase() === "ok") {
                            resolve({});
                        } else {
                            reject("RestPost - Status: " + spRequest.status + ", StatusText: " + spRequest.statusText);
                        }
                    } else {
                        reject('RestPost - Status==' + spRequest.status.toString() + ': ' + spRequest.responseText);
                    }
                }
            };
            spRequest.onerror = () => {
                reject('Resultados: restPost - Error: ' + spRequest.responseText);
            };

            if (fileData) {
                spRequest.send(fileData);
            } else if (jsonBody) {
                spRequest.send(jsonBody);
            } else {
                spRequest.send();
            }
        });
    }

    /**
     * Sets up POST request for adding Files
     * @param url 
     * @param qry 
     * @param data 
     */
    public static restPostAddFile(url: string, qry: string, data: ArrayBuffer): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, null, false, data).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }

    /**
     * Sets up POST request for adding Folders
     * @param url 
     * @param qry 
     * @param serverRelativefolder 
     */
    public static restPostAddFolder(url: string, qry: string, serverRelativefolder: string): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, serverRelativefolder, false).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }

    /**
     * Sets up POST request for deleting files or folders
     * @param url 
     * @param qry 
     * @param data 
     */
    public static restPostDelete(url: string, qry: string, data?: ArrayBuffer): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, null, true, data).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }

    /**
     * Sets up POST request for deleting items
     * @param url 
     * @param qry 
     * @param data 
     */
    public static restPostDeleteItem(url: string, qry: string): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, null, true).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }

    /**
     * Sets up POST request for adding Files
     * @param url 
     * @param qry 
     * @param data 
     */
    public static restPostAddItem(url: string, qry: string, listItemEntity: string, data: SPData[]): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, null, false, null, listItemEntity, data).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }

    /* No utiliza de momento
    public static restPut(qry: string, data?: string): Promise<any> {
        return new Promise<void>((resolve, reject) => {
            var spRequest = new XMLHttpRequest();
            spRequest.open('PUT', qry, true);
            spRequest.setRequestHeader("Accept", "application/json");
            spRequest.onreadystatechange = () => {
                if (spRequest.readyState === 4 && spRequest.status === 200) {
                    var result = JSON.parse(spRequest.responseText);
                    resolve(result);
                } else if (spRequest.readyState === 4 && spRequest.status !== 200) {
                    LogAx.trace('Resultados: restPut - Status==' + spRequest.status.toString() + ': ' + spRequest.responseText);
                    reject();
                }
            };
            spRequest.onerror = () => {
                LogAx.trace('Resultados: restPut - Error: ' + spRequest.responseText);
                reject();
            };
            if (data) {
                spRequest.send(JSON.stringify(data));
            } else {
                spRequest.send();
            }
        });
    }*/

    /**
     * DigestValue cache: avoid sending a request if already available and not expired
     */
    private static cacheRequestDigest: string = null;
    private static cacheRequestDigestExpire: number = null;

    /**
     * Read digest value necessary for POST and PUT updates
     * Since this code reads from outside a SharePoint site, a request is necessary
     * Apps in a SharePoint page can read this value directly from "__REQUESTDIGEST" input element
     * @param url 
     */
    public static requestDigestValue(url: string): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheRequestDigest && this.cacheRequestDigestExpire > new Date().getTime()) {
                resolve(this.cacheRequestDigest);
            }
            let qry: string = SPRest.queryDigestValue(url);
            this.restPost(qry).then((resultadoJSON) => {
                let ctxWebInfo: any = resultadoJSON.d.GetContextWebInformation;
                let expireSeconds: number = +ctxWebInfo.FormDigestTimeoutSeconds;   // 1800 seconds could be the limiting factor for most big uploads, before 2GB size limit. Just hope it doesn't expire in middle of operation
                this.cacheRequestDigest = ctxWebInfo.FormDigestValue;
                this.cacheRequestDigestExpire = new Date().setSeconds(expireSeconds - 60); // save in cache the expiration date (less a minute just in case)
                resolve(this.cacheRequestDigest);
            }, (e) => {
                reject("RequestDigestValue read Error: " + e);
            });
        });
    }

    /* Querys GET */
    public static queryPermissionsList(url: string, listId: string): string {
        return Strings.safeURL(url) + "_api/web/Lists(guid'" + listId + "')/EffectiveBasePermissions";
    }

    public static querySiteInfo(url: string): string {
        let qry: string = Strings.safeURL(url) + "_api/Web";
        //Select fields
        return qry + "?$Select=Id,Title,UIVersion,Language";
    }

    public static queryLists(url: string, baseType: number, showHidden: boolean): string {
        let qry: string = Strings.safeURL(url) + "_api/Web/Lists";
        //Select fields
        qry += "?$Select=Id,Title,Description,Hidden,EntityTypeName,ItemCount,Created,LastItemModifiedDate";
        //Where
        qry += "&$filter=BaseType eq " + baseType.toString() + " and Hidden eq " + (showHidden ? "true" : "false");
        //Order
        return qry + "&$OrderBy=Title asc";
    }

    public static queryList(url: string, id: string): string {
        let qry: string = Strings.safeURL(url) + "_api/Web/Lists(guid'" + id + "')";
        //Select fields
        qry += "?$Select=Id,Title,Description,Hidden,EntityTypeName,ItemCount,Created,LastItemModifiedDate,EntityTypeName,ListItemEntityTypeFullName,RootFolder/Name,RootFolder/UniqueID,RootFolder/ItemCount,RootFolder/ServerRelativeUrl";
        qry += "&$Expand=RootFolder"
        return qry;
    }

    public static queryListFields(url: string, idList: string): string {
        let qry: string = Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Fields";
        //Select fields
        qry += "?$Select=Id,Title,InternalName,StaticName,Description,Hidden,Required,FieldTypeKind,LookupField";
        //Where
        qry += "&$filter=Hidden eq false";
        return qry;
    }

    public static queryLibraryItemsWithView(url: string, idList: string, view: SPView): string {
        let expands: string[] = ["Author", "Editor", "Folder", "Folder/ParentFolder", "File"];
        let fixedFields: string[] = ["Id", "ID", "EncodedAbsUrl", "FileRef", "FileLeafRef", "Folder", "File", "Author", "Created", "Editor", "Modified"];
        let qry: string = Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items";

        //Select fields. Long queries get cut off, so use wildcard selector for normal columns      
        if (view.ViewFields.length > 32) {
            qry += "?$Select=*,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Folder/Name,Folder/UniqueId,Folder/ItemCount,Folder/ServerRelativeUrl,Folder/ParentFolder/UniqueId,File/Length";
        } else {
            qry += "?$Select=Id,EncodedAbsUrl,FileRef,FileLeafRef,Created,Modified,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Folder/Name,Folder/UniqueId,Folder/ItemCount,Folder/ServerRelativeUrl,Folder/ParentFolder/UniqueId,File/Length";
        }

        for (let field of view.ViewFields) {
            if (fixedFields.indexOf(field.InternalName) !== -1)
                continue;

            let odataInternalName: string = SP.safeOdataField(field.InternalName);

            switch (field.Type) {
                case 7: //Lookup
                    qry += "," + odataInternalName + "/" + SP.safeOdataField(field.LookupField);
                    expands.push(odataInternalName);
                    break;
                case 20: //User
                    qry += "," + odataInternalName + "/Title," + odataInternalName + "/EMail";
                    expands.push(odataInternalName);
                    break;
                default:
                    if (view.ViewFields.length <= 32) { // not necessary with wildcard
                        qry += "," + odataInternalName;
                    }
                    break;
            }
        }
        //Select - Expand
        if (expands.length) {
            qry += "&$Expand=" + expands.join(",");
        }
        //Limit
        qry += "&$Top=" + ((view.RowLimit) ? view.RowLimit : 500);
        return qry
    }

    public static queryListItems(url: string, idList: string, rowLimit?: number): string {
        let qry: string = Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items";
        //Select fields
        qry += "?$Select=*,EncodedAbsUrl,FileRef,FileLeafRef,Author/Title,Author/EMail,Editor/Title,Editor/EMail";
        //Select - Expand
        qry += "&$Expand=Author,Editor";
        //Limit
        qry += "&$Top=" + ((rowLimit) ? rowLimit : 100);
        return qry
    }

    public static queryListItemsWithView(url: string, idList: string, view: SPView): string {
        let expands: string[] = ["Author", "Editor"];
        let fixedFields: string[] = ["Id", "ID", "Title", "Author", "Created", "Editor", "Modified"];
        let qry: string = Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items";

        //Select fields. Long queries get cut off, so use wildcard selector for normal columns      
        if (view.ViewFields.length > 32) {
            qry += "?$Select=*,Author/Title,Author/EMail,Editor/Title,Editor/EMail";
        } else {
            qry += "?$Select=Id,Title,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Created,Modified";
        }

        for (let field of view.ViewFields) {
            if (fixedFields.indexOf(field.InternalName) !== -1)
                continue;

            let odataInternalName: string = field.InternalName.startsWith("_x") ? "OData_" + field.InternalName : field.InternalName;

            switch (field.Type) {
                case 7: //Lookup
                    let odataLookupField: string = field.LookupField.startsWith("_x") ? "OData_" + field.LookupField : field.LookupField;
                    qry += "," + odataInternalName + "/" + odataLookupField + "," + odataInternalName + "/ID";
                    expands.push(odataInternalName);
                    break;
                case 20: //User
                    qry += "," + odataInternalName + "/Title," + odataInternalName + "/EMail";
                    expands.push(odataInternalName);
                    break;
                default:
                    if (view.ViewFields.length <= 32) { // not necessary with wildcard
                        qry += "," + odataInternalName;
                    }
                    break;
            }
        }
        //Select - Expand
        if (expands.length) {
            qry += "&$Expand=" + expands.join(",");
        }
        //Limit
        qry += "&$Top=" + ((view.RowLimit) ? view.RowLimit : 100);
        return qry
    }

    public static queryFileData(url: string, idList: string, idItem: number): string {
        return Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items(" + idItem + ")/File/$value";
    }

    public static queryWebs(url: string) {
        let qry: string = Strings.safeURL(url) + "_api/Web/Webs";
        //Select fields
        qry += "?$Select=Id,Title,Url,ServerRelativeUrl";
        //Order
        return qry + "&$OrderBy=Title asc";
    }

    public static queryWeb(url: string) {
        let qry: string = Strings.safeURL(url) + "_api/Web";
        //Select fields
        qry += "?$Select=Id,Title,Description,Url,ServerRelativeUrl,Created";
        //Order
        return qry + "&$OrderBy=Title asc";
    }

    public static queryDigestValue(url: string) {
        return Strings.safeURL(url) + "_api/contextinfo";
    }

    public static queryFolder(url: string, path: string) {
        return Strings.safeURL(url) + "_api/Web/GetFolderByServerRelativeUrl('" + path + "')/Folders?$Select=Name";
    }

    public static queryViews(url: string, idList: string): string {
        let qry: string = Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Views";
        //Select fields
        qry += "?$Select=Id,Title,DefaultView,Paged,PersonalView,RowLimit,ServerRelativeUrl,ViewFields/Items";
        //Expando to other api methods
        qry += "&$Expand=ViewFields"
        //Filter
        qry += "&$filter=Hidden eq false";
        return qry;
    }

    public static querySiteUsers(url: string): string {
        return Strings.safeURL(url) + "_api/web/SiteUsers?$Select=Id,Title,Email,IsSiteAdmin,UserPrincipalName";
    }

    public static querySiteGroupsUsers(url: string): string {
        return Strings.safeURL(url) + "_api/Web/SiteGroups?$Select=Title,PrincipalType,Users/Id,Users/Title,Users/Email,Users/IsSiteAdmin&$Expand=Users";
    }

    /* Querys POST */
    public static queryPostFolders(url: string) {
        return Strings.safeURL(url) + "_api/web/folders"
    }

    public static queryPostAddFile(url: string, path: string, name: string) {
        return Strings.safeURL(url) + "_api/Web/GetFolderByServerRelativeUrl('" + path + "')/Files/Add(overwrite=true, url='" + name + "')";
    }

    public static queryPostDeleteFolder(url: string, pathAndName: string, recycle: boolean) {
        let qry: string = Strings.safeURL(url) + "_api/Web/GetFolderByServerRelativeUrl('" + pathAndName + "')";
        if (recycle) {
            qry += "/Recycle()";
        }
        return qry;
    }

    public static queryPostDeleteFile(url: string, pathAndName: string, recycle: boolean) {
        let qry: string = Strings.safeURL(url) + "_api/Web/GetFileByServerRelativeUrl('" + pathAndName + "')";
        if (recycle) {
            qry += "/Recycle()";
        }
        return qry;
    }

    public static queryPostDeleteItem(url: string, idList: string, idItem: number) {
        return Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/items(" + idItem.toString() + ")";
    }

    public static queryPostAddItem(url: string, idList: string) {
        return Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/items";
    }


}
