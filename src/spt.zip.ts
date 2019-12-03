import * as JSZip from "jszip";
import { SPListItem, SPList } from "./sharepoint/spt.sharepoint.entities";
import { FileData } from "./sharepoint/spt.sharepoint";
import { SPOps } from "./sharepoint/spt.sharepoint.operations";

export class ZIP {
    public downloading: boolean;
    public downloadProcessTotal: number;
    public downloadProcessedItems: number;
    public downloadErrorItems: number;
    public downloadCurrentBytes: number;
    public compressing: boolean;
    public compressProcessedPercentage: number;
    public cancelOperation: boolean;
    public batchSize: number = 2;
    public compression: number = 6;

    private url: string;
    private list: SPList;

    constructor(url: string, list: SPList) {
        this.url = url;
        this.list = list;
        this.downloading = false;
        this.downloadProcessTotal = 0;
        this.downloadProcessedItems = 0;
        this.downloadErrorItems = 0;
        this.downloadCurrentBytes = 0;
        this.compressing = false;
        this.compressProcessedPercentage = 0;
        this.cancelOperation = false;
    }

    /**
     * Create new ZIP into Blob ready for download
     * @param items 
     */
    public createDownloadZip(items: SPListItem[]): Promise<string> {
        let zip: JSZip = new JSZip();
        this.downloadProcessTotal = items.length;

        return new Promise<string>(async (resolve, reject) => {
            //Fase 1: Descarga desde SharePoint. 
            //Se mantienen actualizadas las métricas (processedItems, errorItems)
            //Se permite la cancelación (cancelOperation)
            await this.downloadSharePointFiles(zip, items);

            //Fase 2: Se crea el objeto zip. Se almacena como blob en el DOM a la espera de que el usuario inicie la descarga
            //No se actualizan métricas, pero se podría si el delay es significativo
            this.compressing = true;
            let copyThis: this = this;
            zip.generateAsync({
                type: "blob",
                compression: "DEFLATE",
                compressionOptions: {
                    level: this.compression
                }
            }, function updateCallback(metadata) {
                copyThis.compressProcessedPercentage = metadata.percent;
            }).then((content) => {
                let blob: Blob = new Blob([content], { type: 'application/octet-binary' });
                resolve(URL.createObjectURL(blob));
            }).finally(() => {
                this.compressing = false;
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
    private async downloadSharePointFiles(zip: JSZip, items: SPListItem[]): Promise<void> {
        let index = 0;
        this.downloading = true;

        do {
            let batchIndex: number;
            let arrPromises: Promise<FileData>[] = [];
            for (batchIndex = 0; batchIndex < this.batchSize && index < items.length; batchIndex++ , index++) {
                arrPromises.push(
                    SPOps.getFileData(this.url, this.list.ID, items[index])
                );
            }

            if (arrPromises.length) {
                let currentItemsOk: number = await this.processBatch(zip, arrPromises);
                this.downloadProcessedItems += batchIndex;
                this.downloadErrorItems += batchIndex - currentItemsOk;
            }
        } while (!this.cancelOperation && this.downloadProcessedItems < this.downloadProcessTotal);

        this.downloading = false;
    }

    /**
     * Downloads a controlled amount of files at a time. Not a sliding batch, until the last element
     * is downloaded, it doesn't start a new batch
     * @param zip 
     * @param arrPromises 
     */
    private processBatch(zip: JSZip, arrPromises: Promise<FileData>[]): Promise<number> {
        return new Promise<number>((resolve) => {
            Promise.all(arrPromises).then((resultados: FileData[]) => {
                let resultadosOk: FileData[] = resultados.filter(f => !!f);
                for (let fd of resultadosOk) {
                    zip.file(fd.FileName, fd.FileData);
                    this.downloadCurrentBytes += fd.FileLength;
                }
                resolve(resultadosOk.length);
            });
        });
    }
}

