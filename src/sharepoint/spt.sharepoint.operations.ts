import { SPList, SPFolder, SPListItem, SPView, SPField, SPData } from "./spt.sharepoint.entities";
import { SP, FileData } from "./spt.sharepoint";
import { IListItemLight } from "../spt.storage";
import { SPRest, RestQueryType } from "./spt.sharepoint.rest";
import { LogAx } from "../spt.logax";
import { Constants } from "../spt.constants";
import { string } from "prop-types";

interface IIdPathDictionary {
    id: number,
    path: string,
    isFolder: boolean
}

export class SPOps {
    public copyingProcessTotal: number;
    public copyingProcessedOK: number;
    public copyingProcessedError: number;
    public deletingProcessTotal: number;
    public deletingProcessedOK: number;
    public deletingProcessedError: number;
    public cancelOperation: boolean;

    //Temporal control variable, try to use the next member instead. This value is set on paste event, while the other is set on folder checked event.
    //The reason this is still here is because sometimes on normal paste event the folder is set. I'll have to figure this one out sometime...
    public copyIntoFolder: boolean;
    public copyIntoFolderTarget: SPFolder;

    //private batchSize: number = 5;

    constructor() {
        this.cancelOperation = false;
    }

    public libraryDeleteItems(url: string, list: SPList, items: SPListItem[]): Promise<void> {
        this.deletingProcessTotal = items.length;
        this.deletingProcessedOK = 0;
        this.deletingProcessedError = 0;

        return new Promise<void>(async (resolve) => {
            let paths: IIdPathDictionary[] = [];
            for (let item of items) {
                paths.push({
                    id: item.ID,
                    path: SP.getFolderPathWithoutParent(item.Folder.ParentFolder, "") + item.Name,
                    isFolder: item.Folder.UniqueId !== null
                });
            }
            // Order: have deepest levels processed first, all while current API doesn't permit recursive delete. Otherwise it would have been much easier
            paths.sort((a, b) => {
                let levelA: number = a.path.split("/").length;
                let levelB: number = b.path.split("/").length;
                if (levelA === levelB)
                    return b.path.localeCompare(a.path);
                return (levelA < levelB) ? 1 : -1;
            });
            // Delete in order
            let itemIdx: number = 0;
            do {
                let item: IIdPathDictionary = paths[itemIdx];

                let ok: boolean = await this.deleteSharePointFile(url, list.RootFolder.ServerRelativeUrl, item.path, item.isFolder, true);
                if (ok) {
                    this.deletingProcessedOK++;
                } else {
                    this.deletingProcessedError++;
                }
            } while (!this.cancelOperation && ++itemIdx < this.deletingProcessTotal);

            resolve();
        });
    }

    public listDeleteItems(url: string, list: SPList, items: SPListItem[]): Promise<void> {
        this.deletingProcessTotal = items.length;
        this.deletingProcessedOK = 0;
        this.deletingProcessedError = 0;
        return new Promise<void>(async (resolve) => {
            let itemIdx: number = 0;
            do {
                let item: SPListItem = items[itemIdx];

                let ok: boolean = await this.deleteSharePointItem(url, list.ID, item.ID);
                if (ok) {
                    this.deletingProcessedOK++;
                } else {
                    this.deletingProcessedError++;
                }
            } while (!this.cancelOperation && ++itemIdx < this.deletingProcessTotal);

            resolve();
        });
    }

    public libraryCopyPaste(urlOrigen: string, idListOrigen: string, urlDestino: string, listDestino: SPList, items: IListItemLight[]): Promise<void> {
        this.copyingProcessTotal = items.length;
        this.copyingProcessedOK = 0;
        this.copyingProcessedError = 0;

        let targetRelativePath: string = this.copyIntoFolder ? this.copyIntoFolderTarget.ServerRelativeUrl : listDestino.RootFolder.ServerRelativeUrl;

        return new Promise<void>(async (resolve) => {
            // Create folder structure array. For each relativePath it checks and adds if not included every folder level.
            // To only create the necessary parent folders, omit folders that aren't in the items array!
            // Create a dictionary with the origin folder and the target folder, necessary when uploading files.
            let folderStructure: string[] = [];
            let folderStructureTranslation: IIdPathDictionary[] = [];

            for (let item of items) {
                let targetFolderLevel: string = "";
                if (!folderStructure.includes(item.File.FilePath)) {
                    let folderLevels: string[] = item.File.FilePath.split('/').filter(f => items.filter(i => f === i.File.FileName).length);

                    for (let folderLevel of folderLevels) {
                        targetFolderLevel += "/" + folderLevel;
                        if (!folderStructure.includes(targetFolderLevel)) {
                            folderStructure.push(targetFolderLevel);
                        }
                    }
                }
                folderStructureTranslation.push({
                    id: item.ID,
                    path: targetFolderLevel + "/" + item.File.FileName,
                    isFolder: item.File.Length === 0   // I'll have to find something more robust than this, but for this process it isn't used
                });
            }
            // Order: have smaller levels processed first, that way next levels will have the previous folder created
            folderStructure.sort((a, b) => {
                let levelA: number = a.split("/").length;
                let levelB: number = b.split("/").length;
                if (levelA === levelB)
                    return a.localeCompare(b);  // could get away returning directly 0, but hey... do it right I guess.
                return (levelA < levelB) ? -1 : 1;
            });
            // Create folders in target
            for (let folder of folderStructure) {
                await this.uploadSharePointFolder(urlDestino, targetRelativePath, folder);
            }

            // Download & Upload
            let itemIdx: number = 0;
            do {
                let item: IListItemLight = items[itemIdx];
                if (item.File.Length) {
                    let fd = await this.downloadSharePointFile(urlOrigen, idListOrigen, item);
                    if (fd) {
                        let uploaded = await this.uploadSharePointFile(urlDestino, targetRelativePath, item.ID, fd, folderStructureTranslation);
                        if (uploaded) {
                            this.copyingProcessedOK++;
                        } else {
                            this.copyingProcessedError++;
                        }
                    } else {
                        this.copyingProcessedError++;
                    }
                } else {
                    this.copyingProcessedOK++;
                }
            } while (!this.cancelOperation && ++itemIdx < this.copyingProcessTotal);

            resolve();
        });
    }

    public listPaste(urlDestino: string, listDestino: SPList, items: IListItemLight[]): Promise<void> {
        this.copyingProcessTotal = items.length;
        this.copyingProcessedOK = 0;
        this.copyingProcessedError = 0;

        return new Promise<void>(async (resolve) => {
            // Remove item data fields that are not present in target list
            let targetItems: IListItemLight[] = await this.removeFieldsNotInList(urlDestino, listDestino, items);

            // Insert items
            let itemIdx: number = 0;
            do {
                let item: IListItemLight = targetItems[itemIdx];
                let uploaded = await this.insertSharePointItem(urlDestino, listDestino, item);
                if (uploaded) {
                    this.copyingProcessedOK++;
                } else {
                    this.copyingProcessedError++;
                }
            } while (!this.cancelOperation && ++itemIdx < this.copyingProcessTotal);

            resolve();
        });
    }

    private removeFieldsNotInList(urlDestino: string, listDestino: SPList, items: IListItemLight[]): Promise<IListItemLight[]> {
        return new Promise<IListItemLight[]>(async (resolve) => {
            SPOps.loadFields(urlDestino, listDestino.ID).then((targetFields) => {
                for (let iField = items[0].ItemData.length - 1; iField >= 0; iField--) {
                    let dataField: SPData = items[0].ItemData[iField];
                    if (!targetFields.find(tf => tf.InternalName === dataField.InternalName && tf.Type === dataField.Type)) {
                        for (let item of items) {
                            item.ItemData.splice(iField, 1);
                        }
                    }
                }
                resolve(items);
            });
        });
    }

    /**
     * Download files form SharePoint, using REST interface. 
     * Creates batches of download processes.
     * Updates information after each batch.
     * @param zip 
     * @param items 
     */
    private async downloadSharePointFile(url: string, idList: string, item: IListItemLight): Promise<FileData> {
        return new Promise<FileData>((resolve) => {
            SPOps.getFileDataFromLightItem(url, idList, item).then((value) => {
                resolve(value);
            });
        });
    }

    /**
     * Upload files to SharePoint, using REST interface.
     * Is capable of uploading to specific folders with targetServerPath
     * @param url 
     * @param targetServerPath 
     * @param fileID 
     * @param fileData 
     * @param folderStructureTranslation 
     */
    private async uploadSharePointFile(url: string, targetServerPath: string, fileID: number, fileData: FileData, folderStructureTranslation: IIdPathDictionary[]): Promise<boolean> {
        return new Promise<boolean>((resolve) => {
            let tranlatedFileName: string = folderStructureTranslation.find(f => f.id === fileID).path;
            let name: string = tranlatedFileName;
            let path: string = "";
            if (tranlatedFileName.indexOf("/") !== -1) {
                name = tranlatedFileName.substring(tranlatedFileName.lastIndexOf("/") + 1);
                path = tranlatedFileName.substring(0, tranlatedFileName.lastIndexOf("/"));
                if (path.startsWith("/")) {
                    path = path.substring(1);
                }
            }

            targetServerPath = targetServerPath + "/" + path;

            let query = SPRest.queryPostAddFile(url, targetServerPath, name);
            SPRest.restPostAddFile(url, query, fileData.FileData).then((postResult) => {
                resolve(true);
            }, (e) => {
                LogAx.trace("uploadSharePointFile failed origin file '" + fileData.FileName + "' to target path '" + targetServerPath + "': " + e);
                resolve(false);
            });
        });
    }

    /**
     * Create folder in SharePoint
     * @param url 
     * @param targetRelativePath 
     * @param folder 
     */
    private async uploadSharePointFolder(url: string, targetRelativePath: string, folder: string): Promise<boolean> {
        return new Promise<boolean>((resolve) => {
            // First detect if creating folder is necessary. Read current folders in parent folder structure / root structure
            let parentFolder: string = "/";
            let newChildFolder: string = folder.substring(1);
            if (folder.lastIndexOf("/") > 0) {
                parentFolder = folder.substring(0, folder.lastIndexOf("/"));
                newChildFolder = folder.substring(folder.lastIndexOf("/") + 1);
            }
            let folderQuery: string = SPRest.queryFolder(url, targetRelativePath + parentFolder);
            SPRest.restQuery(folderQuery, RestQueryType.ODataJSON).then((r) => {
                let foundFolder: boolean = r.value.find((i: any) => i.Name === newChildFolder);
                if (foundFolder) {
                    LogAx.trace("uploadSharePointFolder skipping folder creation, already exists: '" + folder + "'");
                    resolve(true);
                }

                let serverRelativeCompleteFolderPath: string = targetRelativePath + folder;
                let query: string = SPRest.queryPostFolders(url);

                SPRest.restPostAddFolder(url, query, serverRelativeCompleteFolderPath).then((postResult) => {
                    resolve(true);
                }, (e) => {
                    LogAx.trace("uploadSharePointFolder failed for folder '" + folder + "' in path '" + targetRelativePath + "': " + e);
                    resolve(false);
                });

            }, (e) => {
                LogAx.trace("uploadSharePointFolder failed read folders in path '" + targetRelativePath + "': " + e);
                resolve(false);
            });
        });
    }

    private async deleteSharePointFile(url: string, targetRelativePath: string, itemPath: string, isFolder: boolean, recycle: boolean): Promise<boolean> {
        return new Promise<boolean>((resolve) => {
            let fullPath: string = targetRelativePath;
            if (!itemPath.startsWith("/")) {
                fullPath += "/"
            }
            fullPath += itemPath;

            let query: string = "";
            if (isFolder) {
                query = SPRest.queryPostDeleteFolder(url, fullPath, recycle);
            } else {
                query = SPRest.queryPostDeleteFile(url, fullPath, recycle);
            }

            SPRest.restPostDelete(url, query).then((postResult) => {
                resolve(true);
            }, (e) => {
                LogAx.trace("deleteSharePointFile failed for file '" + fullPath + "': " + e);
                resolve(false);
            });
        });
    }

    private readonly omitInsertInternalFields = ["ID", "Author", "Editor", "Created", "Modified"];

    private async insertSharePointItem(url: string, targetList: SPList, items: IListItemLight): Promise<boolean> {
        return new Promise<boolean>((resolve) => {
            let cleanedItems: SPData[] = items.ItemData.filter(i => this.omitInsertInternalFields.includes(i.InternalName) === false);

            let query = SPRest.queryPostAddItem(url, targetList.ID);
            SPRest.restPostAddItem(url, query, targetList.ListItemEntityTypeFullName, cleanedItems).then((postResult) => {
                resolve(true);
            }, (e) => {
                LogAx.trace("insertSharePointItem failed: " + e);
                resolve(false);
            });
        });
    }

    private async deleteSharePointItem(url: string, idList: string, idItem: number): Promise<boolean> {
        return new Promise<boolean>((resolve) => {
            let query: string = SPRest.queryPostDeleteItem(url, idList, idItem);

            SPRest.restPostDeleteItem(url, query).then((postResult) => {
                resolve(true);
            }, (e) => {
                LogAx.trace("deleteSharePointFile failed for item '" + idItem.toString() + "': " + e);
                resolve(false);
            });
        });
    }

    /**
     * Loads List/Library from SharePoint
     */
    public static async loadList(url: string, idList: string): Promise<SPList> {
        return new Promise((resolve, reject) => {
            let qry: string = SPRest.queryList(url, idList);
            LogAx.trace("Query List/Library:" + qry);
            SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
                try {
                    resolve({
                        ID: r.Id,
                        Title: r.Title,
                        ItemCount: r.ItemCount,
                        Hidden: r.Hidden,
                        InternalName: r.EntityTypeName,
                        ListItemEntityTypeFullName: r.ListItemEntityTypeFullName,
                        RootFolder: {
                            UniqueId: r.RootFolder.UniqueId,
                            Name: r.RootFolder.Name,
                            ItemCount: r.RootFolder.ItemCount,
                            ServerRelativeUrl: r.RootFolder.ServerRelativeUrl,
                            Level: 0,
                            Expand: true
                        }
                    });
                } catch (e) {
                    reject("SPT.SharePoint.Operations.LoadList exception: " + e);
                }
            }, (e) => {
                reject(e);
            });
        });
    }

    /**
     * Loads Views for the library.
     */
    public static async loadViews(url: string, idList: string, listFields: SPField[]): Promise<SPView[]> {
        return new Promise((resolve, reject) => {
            let qry: string = SPRest.queryViews(url, idList);
            LogAx.trace("Query Views:" + qry);

            SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
                try {
                    resolve(r.value.map((f: any) => ({
                        ID: f.Id,
                        Title: f.Title,
                        DefaultView: f.DefaultView as boolean,
                        PersonalView: f.PersonalView as boolean,
                        RowLimit: f.Rowlimit as number,
                        ServerRelativeUrl: f.ServerRelativeUrl,
                        ViewFields: SPOps.viewFieldArrayToSPFields(f.ViewFields.Items, listFields)
                    })));
                } catch (e) {
                    reject("SPT.SharePoint.Operations.LoadViews ODataAx.restQuery error: " + e);
                }
            }, (e) => {
                reject(e);
            });
        });
    }

    /**
     * Assign real available fields from a views viewfields property
     */
    private static viewFieldArrayToSPFields(fields: string[], listFields: SPField[]): SPField[] {
        let r: SPField[] = [];
        for (let field of fields) {
            if (field === "LinkTitle") { // LinkTitle gives problems on certain operations, use Title
                field = "Title";
            }
            let spf: SPField = listFields.find(f => f.InternalName === field);
            if (spf) {
                r.push(spf);
            }
        }
        return r;
    }

    /**
     * Loads Field data for the library
     */
    public static async loadFields(url: string, idList: string): Promise<SPField[]> {
        return new Promise((resolve, reject) => {
            let qry: string = SPRest.queryListFields(url, idList);
            LogAx.trace("Query Fields:" + qry);
            SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
                try {
                    resolve(r.value.map((f: any) => ({
                        ID: f.Id,
                        Title: f.Title,
                        InternalName: f.InternalName,
                        StaticName: f.StaticName,
                        Description: f.Description,
                        Hidden: f.Hidden as boolean,
                        Required: f.Required as boolean,
                        Type: f.FieldTypeKind,    //https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee540543%28v%3doffice.15%29
                        LookupField: f.LookupField
                    })));
                } catch (e) {
                    reject("SPT.SharePoint.Operations.loadFields ODataAx.restQuery error: " + e);
                }
            }, (e) => {
                reject(e);
            });
        });
    }

    /**
     * Load item data from list, based on view
     * @param url 
     * @param idList 
     * @param view 
     * @param isLibrary 
     */
    public static async loadItems(url: string, idList: string, view: SPView, isLibrary: boolean): Promise<any> {
        return new Promise((resolve, reject) => {
            let qry: string;
            if (isLibrary) {
                qry = SPRest.queryLibraryItemsWithView(url, idList, view);
            } else {
                qry = SPRest.queryListItemsWithView(url, idList, view);
            }
            LogAx.trace("Query Items:" + qry);
            SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
                resolve(r);
            }, (e) => {
                reject(e);
            });
        });
    }

    /**
     * Read contents from file in binary. Returns string of type "data:application/octet-stream;base64"
     * @param url
     * @param idLista 
     * @param item 
     */
    public static getFileData(url: string, idLista: string, item: SPListItem): Promise<FileData> {
        return new Promise<FileData>((resolve) => {
            let qry: string = SPRest.queryFileData(url, idLista, item.ID);
            SPRest.restQuery(qry, RestQueryType.ArrayBuffer).then((result: ArrayBuffer) => {
                resolve({
                    FileName: SP.getFolderPath(item.Folder.ParentFolder, item.Name),
                    FileLength: item.Length,
                    FileData: result
                });
            }, (e) => {
                LogAx.trace("GetFile data error on item '" + item.ID + "' in list '" + idLista + "': " + e);
                resolve(null);
            });
        });
    }

    /**
     * Read contents from file in binary. Returns string of type "data:application/octet-stream;base64"
     * Slightly modified version of previous function compatible with IListItemLight instead of SPListItem, necessary for low localstorage use
     * @param url 
     * @param idLista 
     * @param item 
     */
    public static getFileDataFromLightItem(url: string, idLista: string, item: IListItemLight): Promise<FileData> {
        return new Promise<FileData>((resolve) => {
            let qry: string = SPRest.queryFileData(url, idLista, item.ID);
            SPRest.restQuery(qry, RestQueryType.ArrayBuffer).then((result: ArrayBuffer) => {
                resolve({
                    FileName: item.File.FilePath + item.File.FileName,
                    FileLength: item.File.Length,
                    FileData: result
                });
            }, (e) => {
                LogAx.trace("GetFileFromLightItem data error on item '" + item.ID + "' in list '" + idLista + "': " + e);
                resolve(null);
            });
        });
    }

    public static analizeTargetListFields(url: string, idList: string, sourceData: IListItemLight[]): Promise<string[]> {
        return new Promise<string[]>((resolve) => {
            let result: string[] = [];
            let maxErrors: number = 10;
            let errorRequiredNotFound: string = Constants.getLiteral("analisysErrorRequired");
            let errorInternalNotFound: string = Constants.getLiteral("analisysErrorInternalName");
            let errorInternalAndTypeNotFound: string = Constants.getLiteral("analisysErrorInternalNameAndType");

            try {
                this.loadFields(url, idList).then((fields) => {
                    //Test all required fields are included
                    for (let field of fields) {
                        if (result.length > maxErrors) {
                            break;
                        }
                        if (field.Required) {
                            if (!sourceData[0].ItemData.find(i => i.InternalName === field.InternalName)) {
                                result.push(errorRequiredNotFound.replace("%1", field.InternalName));
                                continue;
                            }
                        }
                    }

                    //Test source fields exist in target
                    for (let dataField of sourceData[0].ItemData) {
                        if (result.length > maxErrors) {
                            break;
                        }
                        //Error InternalName doesn't exist
                        if (!fields.find(f => f.InternalName === dataField.InternalName)) {
                            result.push(errorInternalNotFound.replace("%1", dataField.InternalName));
                            continue;
                        }
                        //Type for field not same
                        if (!fields.find(f => f.InternalName === dataField.InternalName && f.Type === dataField.Type)) {
                            result.push(errorInternalAndTypeNotFound.replace("%1", dataField.InternalName));
                        }
                    }

                    //
                    // TODO: Add other checks: Lookups that exist, options that are available, users are added, ...
                    //

                    resolve(result);
                }, (e) => {
                    throw e;
                });
            } catch (e) {
                LogAx.trace("AnalizeTargelListFields error: " + e);
                resolve(["Error"]);
            }
        });
    }
}

