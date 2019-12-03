import { SPData } from "./sharepoint/spt.sharepoint.entities";

export interface IListItemLight {
    ID: number;
    Author: string;
    Editor: string;
    Created: Date;
    Modified: Date;
    File?: {
        FilePath: string;
        FileName: string;
        Length: number;
    }
    ItemData?: SPData[];
}

export interface ICopyPasteInstruction {
    site: string;
    listId: string;
    listType: number;
    items?: IListItemLight[];
}

/**
 * Storage helper class that deals with webextension storage API and exposes "simple to use" functions for common usage
 */
export class WebExStorage {
    private static readonly storageKeyInfo = "sptStorageCyPInfo";
    private static readonly storageKeyData = "sptStorageCyPData";

    /**
     * Retrieve info of C&P Instuctions 
     */
    public static info(): Promise<ICopyPasteInstruction> {
        return new Promise<ICopyPasteInstruction>((resolve, reject) => {
            try {
                browser.storage.local.get(this.storageKeyInfo).then((result: any) => {
                    if (result[this.storageKeyInfo]) {
                        resolve({
                            site: result[this.storageKeyInfo].site,
                            listId: result[this.storageKeyInfo].listId,
                            listType: result[this.storageKeyInfo].listType
                        });
                    }
                    resolve(null);
                }, (e) => {
                    reject("Error in Storage.Info: " + e);
                });
            } catch (e) {
                reject("Exception in Storage.Info: " + e);
            }
        });
    }

    /**
     * Read and Delete C&P Items from storage
     */
    public static get(): Promise<ICopyPasteInstruction> {
        return new Promise<ICopyPasteInstruction>((resolve, reject) => {
            try {
                let getPromiseInfo = new Promise<any>((resolve, reject) => {
                    browser.storage.local.get(this.storageKeyInfo).then((result) => {
                        resolve((result[this.storageKeyInfo]) ? result[this.storageKeyInfo] as any : {});
                    }, (e) => { reject("Error al leer storage list id: " + e) });
                });
                let getPromiseData = new Promise<IListItemLight[]>((resolve, reject) => {
                    browser.storage.local.get(this.storageKeyData).then((result) => {
                        resolve((result[this.storageKeyData]) ? result[this.storageKeyData] as unknown as IListItemLight[] : []);
                    }, (e) => { reject("Error al leer storage data: " + e) });
                });

                Promise.all([getPromiseInfo, getPromiseData]).then((resultados) => {
                    resolve({
                        site: resultados[0].site,
                        listId: resultados[0].listId,
                        listType: resultados[0].listType,
                        items: resultados[1]
                    });
                }, (e) => {
                    reject("Error in Storage.Get promise all: " + e);
                }).finally(() => {
                    // Delete storage used
                    this.clear();
                });
            } catch (e) {
                reject("Exception in Storage.Get: " + e);
            }
        });
    }

    /**
     * Delete storage entries used in this application
     */
    public static clear(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            try {
                let promiseClearBatch: Promise<void>[] = [
                    browser.storage.local.remove(this.storageKeyInfo),
                    browser.storage.local.remove(this.storageKeyData)
                ];
                Promise.all(promiseClearBatch).then(() => {
                    resolve();
                }, (e) => {
                    reject("Error in Storage.Clear promise all: " + e);
                });
            } catch (e) {
                reject("Exception in Storage.Clear: " + e);
            }
        });
    }

    /**
     * Save C&P Items to storages
     * @param cpInstruction
     */
    public static set(cpInstruction: ICopyPasteInstruction): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            try {
                let info: any = {
                    [this.storageKeyInfo]: {
                        site: cpInstruction.site,
                        listId: cpInstruction.listId,
                        listType: cpInstruction.listType
                    }
                };
                browser.storage.local.set(info).then(() => {
                    let data: any = {
                        [this.storageKeyData]: cpInstruction.items
                    };
                    browser.storage.local.set(data).then(() => {
                        resolve(true);
                    }, (e) => {
                        reject("Error al escribir Items en storage: " + e);
                    });
                }, (e) => {
                    reject("Error al escribir Info en storage: " + e);
                });
            } catch (e) {
                reject("Exception in Storage.Set: " + e);
            }
        });
    }

    public static size(): Promise<number> {
        let items: string[] = [this.storageKeyInfo, this.storageKeyData];
        return new Promise<number>((resolve) => {
            browser.storage.local.get(items).then((result) => {
                resolve(JSON.stringify(result).length);
            });
        });
    }
}  
