import * as React from "react";
import { SPList, SPListItem, SPUser, SPData, SPView, SPField } from "../sharepoint/spt.sharepoint.entities";
import { Constants } from "../spt.constants";
import { ListMenu } from "./spt.explorer.menu";
import { ListPaging } from "./spt.explorer.paging";
import { SPOps } from "../sharepoint/spt.sharepoint.operations";
import { LogAx } from "../spt.logax";
import { WebExStorage, ICopyPasteInstruction, IListItemLight } from "../spt.storage";
import { SP } from "../sharepoint/spt.sharepoint";
import { Dates } from "../spt.dates";
import { SPTModalDownloadExcel, SPTModalDeleteList, SPTModalCopyList, SPTModalImportExcel } from "./spt.modal";
import { ListTable } from "./spt.explorer.listtable";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { CSV, CSVDelimiters } from "../spt.csv";
import { Strings } from "../spt.strings";

export interface IExplorerListProps {
    Url: string;
    ID: string;
    Nivel: number;
}

export interface IExplorerListState {
    list: SPList;
    listFields: SPField[];
    availableViews: SPView[];
    view: SPView;
    listItems: SPListItem[];
    listItemsLoaded: boolean;
    checkedListItemsIDs: number[];
    storageSize: number;
    windowWidth: number;
    windowHeight: number;
    downloadExcelProcessing: boolean;
    uploadExcelProcessing: boolean;
    uploadExcelAnalyzing: boolean;
    uploadExcelAnalyzed: boolean;
    uploadExcelUploading: boolean;
    uploadExcelUploaded: boolean;
    uploadExcelAnalisisResults: string[];
    uploadExcelReadingFileProgress: number;
    uploadExcelUploadingRowsOK: number;
    uploadExcelUploadingRowsError: number;
    copyingModalOpen: boolean;
    copyingAnalizing: boolean;
    copyingAnalisisMessages: string[];
    copying: boolean;
    copyingOK: number;
    copyingError: number;
    copyingTotal: number;
    copyingCPI: ICopyPasteInstruction;
    deletingModalOpen: boolean;
    deleting: boolean;
    deletingOK: number;
    deletingError: number;
    deletingTotal: number;
    cancelOperation: boolean;
}

export class ListViewer extends React.Component<IExplorerListProps, IExplorerListState> {
    constructor(props: IExplorerListProps) {
        super(props);
        this.state = {
            list: null,
            listFields: null,
            availableViews: [],
            view: null,
            listItems: [],
            listItemsLoaded: false,
            checkedListItemsIDs: [],
            storageSize: 0,
            windowWidth: 0,
            windowHeight: 0,
            downloadExcelProcessing: false,
            uploadExcelProcessing: false,
            uploadExcelAnalyzing: false,
            uploadExcelAnalyzed: false,
            uploadExcelUploading: false,
            uploadExcelUploaded: false,
            uploadExcelAnalisisResults: [],
            uploadExcelReadingFileProgress: 0,
            uploadExcelUploadingRowsOK: 0,
            uploadExcelUploadingRowsError: 0,
            copyingModalOpen: false,
            copyingAnalizing: false,
            copyingAnalisisMessages: [],
            copying: false,
            copyingOK: 0,
            copyingError: 0,
            copyingTotal: 0,
            copyingCPI: null,
            deletingModalOpen: false,
            deleting: false,
            deletingOK: 0,
            deletingError: 0,
            deletingTotal: 0,
            cancelOperation: false
        };
        this.onDescargar = this.onDescargar.bind(this);
        this.onCloseDescargar = this.onCloseDescargar.bind(this);
        this.onCloseImportar = this.onCloseImportar.bind(this);
        this.onAcceptCopyClause = this.onAcceptCopyClause.bind(this);
        this.onCloseCopy = this.onCloseCopy.bind(this);
        this.onAcceptDeleteClause = this.onAcceptDeleteClause.bind(this);
        this.onCloseDelete = this.onCloseDelete.bind(this);
    }

    render() {
        return <div>
            {
                this.state.list && this.state.view &&
                <div id="listViewer">
                    <div id="listTitle">{this.state.list.Title}
                        <span>[{Constants.getLiteral("explorerBibliotecaInternal")}: {this.state.list.InternalName}]</span>
                        <span>[{Constants.getLiteral("explorerBibliotecaEntityType")}: {this.state.list.ListItemEntityTypeFullName}]</span><br />
                    </div>
                    <ListMenu
                        Url={this.props.Url}
                        ListID={this.props.ID}
                        itemsSelected={this.state.checkedListItemsIDs.length}
                        listType={0}
                        views={this.state.availableViews}
                        onDescargar={() => this.onPrepararExcelDescargable()}
                        onCargarExcel={() => this.onImportarExcel()}
                        onCopiar={() => this.clickCopiar()}
                        onPegar={(mode) => this.clickPegar(mode)}
                        onEliminar={() => this.clickDelete()}
                        onRefresh={() => this.clickRefresh()}
                        onViewUpdate={(viewId) => this.selectedViewUpdate(viewId)} />
                    <ListTable
                        listItems={this.state.listItems}
                        loaded={this.state.listItemsLoaded}
                        windowHeight={this.state.windowHeight}
                        windowWidth={this.state.windowWidth}
                        view={this.state.view}
                        onSelected={(selectedItemIDs, clickedID) => this.updateSelectedItems(selectedItemIDs, clickedID)} />
                    <ListPaging
                        AvailableHeight={this.state.windowHeight}
                        TotalItems={this.state.list.ItemCount}
                        LoadedItems={this.state.listItems.length}
                        onLoadFullList={() => this.loadDataFull()} />
                </div>
            }
            {
                !this.state.list &&
                <div>
                    <img src="icons/ajax-loader.gif" width="12px" /> {Constants.getLiteral("libraryViewerCargando")}
                </div>
            }
            <SPTModalDownloadExcel
                Processing={this.state.downloadExcelProcessing}
                onDownload={this.onDescargar}
                onClose={this.onCloseDescargar} />

            <SPTModalImportExcel
                Processing={this.state.uploadExcelProcessing}
                Analizing={this.state.uploadExcelAnalyzing}
                Analized={this.state.uploadExcelAnalyzed}
                AnalyzedMessage={this.state.uploadExcelAnalisisResults}
                Uploading={this.state.uploadExcelUploading}
                Uploaded={this.state.uploadExcelUploaded}
                ReadingFileProgress={this.state.uploadExcelReadingFileProgress}
                UploadingRowOK={this.state.uploadExcelUploadingRowsOK}
                UploadingRowError={this.state.uploadExcelUploadingRowsError}
                UploadingRowTotal={this.excelItemData.length}
                onAnalyze={(separador, texto) => this.onAnalizarImportar(separador, texto)}
                onUpload={() => this.onImportar()}
                onClose={this.onCloseImportar} />

            <SPTModalCopyList
                Open={this.state.copyingModalOpen}
                Analysing={this.state.copyingAnalizing}
                AnalisisMessages={this.state.copyingAnalisisMessages}
                RowsOK={this.state.copyingOK}
                RowsError={this.state.copyingError}
                RowsTotal={this.state.copyingTotal}
                onAcceptClause={this.onAcceptCopyClause}
                onCloseModal={this.onCloseCopy} />

            <SPTModalDeleteList
                Open={this.state.deletingModalOpen}
                RowsOK={this.state.deletingOK}
                RowsError={this.state.deletingError}
                RowsTotal={this.state.deletingTotal}
                onAcceptClause={this.onAcceptDeleteClause}
                onCloseModal={this.onCloseDelete} />
        </div>
    }

    public componentDidUpdate(prevProps: any, prevState: any) {
        if (prevProps.ID !== this.props.ID) {
            this.setState({
                list: null,
                availableViews: [],
                view: null,
                listItems: [],
                listItemsLoaded: false,
                checkedListItemsIDs: []
            }, () => {
                this.loadData();
            });
        }
    }

    private storageSizeIntervalHandler: number;
    public componentDidMount(): void {
        this.loadData();
        window.addEventListener("resize", this.updateDimensions.bind(this));
        this.updateDimensions();
        this.storageSizeIntervalHandler = window.setInterval(() => {
            WebExStorage.size().then((size) => {
                this.setState({
                    storageSize: size
                });
            });
        }, 2000);
    }

    public componentWillUnmount(): void {
        window.removeEventListener("resize", this.updateDimensions.bind(this));
        window.clearInterval(this.storageSizeIntervalHandler);
    }

    private updateDimensions() {
        this.setState({ windowWidth: window.innerWidth, windowHeight: window.innerHeight });
    }

    /**
    * Load all data for the list
    */
    private loadData() {
        SPOps.loadList(this.props.Url, this.props.ID).then((list: SPList) => {
            this.setState({ list: list });
            SPOps.loadFields(this.props.Url, list.ID).then((fields) => {
                this.setState({ listFields: fields });
                SPOps.loadViews(this.props.Url, list.ID, fields).then((views) => {
                    this.setState({
                        availableViews: views,
                        view: this.state.view ? this.state.view : views.find(v => v.DefaultView === true)
                    }, () => {
                        //After setting state, run initial load with default view
                        this.loadPartialData();
                    });
                }, (e) => {
                    LogAx.trace("SPT.Explorer.ListViewer loadViews error: " + e);
                    this.setState({ view: null });
                });
            }, (e) => {
                LogAx.trace("SPT.Explorer.ListViewer loadFields error: " + e);
                this.setState({ listFields: null });
            });
        }, (e) => {
            LogAx.trace("SPT.Explorer.ListViewer loadList error: " + e);
            this.setState({ list: null });
        });
    }

    private loadPartialData() {
        this.setState({
            listItems: [],
            listItemsLoaded: false,
            checkedListItemsIDs: []
        });

        if (!this.state.list.ID) {
            LogAx.trace("SPT.Explorer.ListViewer loadPartialData error: List ID not set");
            return;
        }
        if (!this.state.view) {
            LogAx.trace("SPT.Explorer.ListViewer loadPartialData error: View not set");
            return;
        }

        SPOps.loadItems(this.props.Url, this.state.list.ID, this.state.view, false).then((dataSet: any) => {
            // Save item data to collection
            this.loadItemData(dataSet);
        }, (e) => {
            LogAx.trace("SPT.Explorer.ListViewer loadItems error: " + e);
        });
    }

    private loadDataFull(): Promise<void> {
        return new Promise((resolve) => {
            this.setState({
                listItems: []
            });

            // For full list loading make sure the rowlimit is somewhat large (about 500 max), because there is less overhead.
            let fullAdaptedView: SPView = this.state.view;
            fullAdaptedView.RowLimit = Math.max(500, fullAdaptedView.RowLimit);

            let qry: string = SPRest.queryListItemsWithView(this.props.Url, this.props.ID, fullAdaptedView);
            LogAx.trace("Query FULL List:" + qry);

            this.loadDataFullRecursive(qry).then(() => {
                resolve();
            });
        });
    }

    private loadDataFullRecursive(qry: string): Promise<void> {
        return new Promise<void>((resolve) => {
            SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
                this.loadItemData(r);

                if (!r["odata.nextLink"]) {
                    resolve();
                }
                this.loadDataFullRecursive(r["odata.nextLink"]).then(() => {
                    resolve();
                });
            }, (e) => {
                LogAx.trace("SPT.explorer.listviewer loadDataFullRecursive error: " + e)
                resolve();
            });
        });
    }

    /**
     * Load data read from web service into local collection
     * @param dataResult
     */
    private loadItemData(dataResult: any): void {
        let batchListItems: SPListItem[] = [];

        dataResult.value.forEach((i: any) => {
            let item: SPListItem = new SPListItem();

            // Process entity data
            item.ID = i["ID"] as number;
            item.Created = i["Created"] as Date;
            item.Author = new SPUser(i.Author.Title, i.Author.EMail);
            item.Modified = i["Modified"] as Date;
            item.Editor = new SPUser(i.Editor.Title, i.Editor.EMail);

            // All usefull data in mapped collection for visualization, using view order
            for (let field of this.state.view.ViewFields) {
                item.Items.push(SP.parseItemJsonResult(i, field));
            }

            batchListItems.push(item);
        });

        this.setState({
            listItems: this.state.listItems.concat(batchListItems),
            listItemsLoaded: true
        });
    }

    /**
    * Handle item selection updates from child table component, mostly used by the menu
    * @param selectedItemIDs 
    * @param clickedID 
    */
    private updateSelectedItems(selectedItemIDs: number[], clickedID: number) {
        this.setState({
            checkedListItemsIDs: selectedItemIDs
        });
    }

    /**
     * Click on menu option to start downloading data and create CSV archive
     */
    private onPrepararExcelDescargable() {
        this.setState({
            downloadExcelProcessing: true
        });
    }

    /**
     * Click on Download in modal. Proceed to download the zip archive
     */
    private onDescargar(delimiter: CSVDelimiters) {
        let items: SPListItem[] = this.state.listItems.filter(i => this.state.checkedListItemsIDs.includes(i.ID));
        let csvData = CSV.generateCSV(items, this.state.view, delimiter);

        let blob = new Blob([Strings.UTFBOMStartCode + csvData], { type: "text/csv;charset=UTF-8" }); //add code to force as UTF-8 with BOM

        browser.downloads.download({
            url: URL.createObjectURL(blob),
            filename: this.state.list.Title + " " + Dates.getFileSuffix() + ".csv",
            saveAs: true
        }).finally(() => {
            this.setState({
                downloadExcelProcessing: false
            });
        });
    }

    private onCloseDescargar() {
        this.setState({
            downloadExcelProcessing: false
        });
    }

    /**
     * Click on menu option to import CSV archive into list
     */
    private excelItemData: IListItemLight[] = [];

    private onImportarExcel() {
        this.setState({
            uploadExcelProcessing: true,
            uploadExcelAnalyzing: false,
            uploadExcelAnalyzed: false,
            uploadExcelAnalisisResults: [],
            uploadExcelUploading: false,
            uploadExcelUploaded: false,
            uploadExcelReadingFileProgress: 0,
            uploadExcelUploadingRowsOK: 0,
            uploadExcelUploadingRowsError: 0,
            cancelOperation: false
        });
    }

    private onAnalizarImportar(delimiter: CSVDelimiters, file: File) {
        this.setState({
            uploadExcelAnalyzing: true
        });

        let reader = new FileReader()
        reader.onprogress = ((e: ProgressEvent<FileReader>) => {
            if (e.lengthComputable) {
                this.setState({
                    uploadExcelReadingFileProgress: Math.round((e.loaded / e.total) * 100)
                });
            }
        });
        reader.onload = ((e: ProgressEvent<FileReader>) => {
            let csvData: string;

            //Read data uploaded
            try {
                let base64Data = e.target.result.toString();
                let headerToken = "base64,";
                let headerIdx: number = base64Data.indexOf(headerToken); //Remove header if exists (data:application/vnd.ms-excel;base64,)
                if (headerIdx !== -1) {
                    base64Data = base64Data.substring(headerIdx + headerToken.length);
                }

                csvData = Strings.b64DecodeUnicode(base64Data);

                this.setState({
                    uploadExcelReadingFileProgress: 100
                });

                if (!csvData) {
                    throw "Nothing read";
                }
            }
            catch (e) {
                LogAx.trace("SPT.explorer.onImportarExcel Read error: " + e);
                return;
            }

            //Parse data
            this.excelItemData = CSV.parseCSV(csvData, this.state.listFields, delimiter);
            if (this.excelItemData.length === 0) {
                LogAx.trace("SPT.explorer.onImportarExcel Data empty");
                return;
            }

            //Analyze data
            try {
                SPOps.analizeTargetListFields(this.props.Url, this.state.list.ID, this.excelItemData).then((results) => {
                    this.setState({
                        uploadExcelAnalyzed: true,
                        uploadExcelAnalisisResults: results
                    });
                });
            }
            catch (e) {
                LogAx.trace("SPT.explorer.onImportarExcel Analyze error: " + e);
                return;
            }
        });
        reader.onerror = ((e: ProgressEvent<FileReader>) => {
            LogAx.trace("SPT.explorer.onImportarExcel error: [" + e.target.error.code + "] " + e.target.error.message);
            this.onCloseImportar();
        });
        reader.readAsDataURL(file);
    }

    private onImportar() {
        this.setState({
            uploadExcelUploading: true
        });

        let spOp: SPOps = new SPOps();

        //This timer keeps modal updated
        let updateImportStatus: number = window.setInterval(() => {
            spOp.cancelOperation = this.state.cancelOperation;
            this.setState({
                uploadExcelUploadingRowsOK: spOp.copyingProcessedOK,
                uploadExcelUploadingRowsError: spOp.copyingProcessedError
            });
        }, 500);

        spOp.listPaste(this.props.Url, this.state.list, this.excelItemData).then(() => {
            clearInterval(updateImportStatus);
            this.setState({
                uploadExcelUploadingRowsOK: spOp.copyingProcessedOK,
                uploadExcelUploadingRowsError: spOp.copyingProcessedError,
                uploadExcelUploaded: true
            });
        });
    }

    private onCloseImportar() {
        this.setState({
            uploadExcelProcessing: false,
            cancelOperation: true
        });
        if (this.state.uploadExcelUploaded) {
            this.loadData();
        }
    }

    /**
     * Click on Copy. Add IDs to internal web storage.
     */
    private clickCopiar() {
        let itemData: IListItemLight[] = this.state.listItems
            .filter(i => this.state.checkedListItemsIDs.includes(i.ID))
            .map(li => ({
                ID: li.ID,
                Author: li.Author.Email,
                Editor: li.Editor.Email,
                Created: li.Created,
                Modified: li.Modified,
                ItemData: li.Items.map(spd => ({
                    InternalName: spd.InternalName,
                    StringValue: spd.StringValue,
                    LookupId: spd.LookupId,
                    Type: spd.Type
                }) as SPData)
            }) as IListItemLight);

        if (itemData.length) {
            let cpInstruction: ICopyPasteInstruction = {
                site: this.props.Url,
                listId: this.props.ID,
                listType: 0,
                items: itemData
            };
            WebExStorage.set(cpInstruction).catch((reason) => {
                LogAx.trace("Error clickCopiar Storage.set: " + reason);
            });
        } else {
            WebExStorage.clear().catch((reason) => {
                LogAx.trace("Error clickCopiar Storage.clear: " + reason);
            });
        }
    }

    /**
     * Click on Paste. Paste files from previous IDs saved. Even on another site.
     * @param mode True for pasting in specific folder
     */
    private clickPegar(mode: boolean) {
        WebExStorage.get().then((cpi: ICopyPasteInstruction) => {
            this.setState({
                copyingModalOpen: true,
                copyingAnalizing: true,
                copyingAnalisisMessages: [],
                copying: false,
                cancelOperation: false,
                copyingOK: 0,
                copyingError: 0,
                copyingTotal: 0,
                copyingCPI: cpi
            });
            SPOps.analizeTargetListFields(this.props.Url, this.state.list.ID, cpi.items).then((results) => {
                this.setState({
                    copyingAnalizing: false,
                    copyingAnalisisMessages: results
                });
            });
        }, (e) => {
            LogAx.trace("Error in clickPegar WebStorage Get: " + e);
            this.onCloseCopy();
        });
    }

    private onAcceptCopyClause() {
        let spOp: SPOps = new SPOps();

        this.setState({
            copying: true,
            copyingTotal: this.state.copyingCPI.items.length
        });

        //This timer keeps modal updated
        let updateCopyPasteStatus: number = window.setInterval(() => {
            spOp.cancelOperation = this.state.cancelOperation;
            this.setState({
                copyingOK: spOp.copyingProcessedOK,
                copyingError: spOp.copyingProcessedError
            });
        }, 500);

        spOp.listPaste(this.props.Url, this.state.list, this.state.copyingCPI.items).then(() => {
            clearInterval(updateCopyPasteStatus);
            this.setState({
                copyingOK: spOp.copyingProcessedOK,
                copyingError: spOp.copyingProcessedError
            });
        });
    }

    private onCloseCopy() {
        this.setState({
            cancelOperation: true,
            copying: false,
            copyingModalOpen: false,
            copyingCPI: null
        });
        this.loadData();
    }

    /**
     * Click on Delete. Removes selected items in reverse fashion to avoid having the system error for deleting a non empty folder.
     */
    private clickDelete() {
        this.setState({
            deletingModalOpen: true,
            deleting: false,
            cancelOperation: false,
            deletingOK: 0,
            deletingError: 0,
            deletingTotal: 0
        });
    }

    private onAcceptDeleteClause() {
        let itemsDelete: SPListItem[] = this.state.listItems.filter(i => this.state.checkedListItemsIDs.includes(i.ID));

        let spOp: SPOps = new SPOps();
        this.setState({
            deleting: true,
            deletingTotal: itemsDelete.length
        });

        //This timer keeps modal updated
        let updateDeleteStatus: number = window.setInterval(() => {
            spOp.cancelOperation = this.state.cancelOperation;
            this.setState({
                deletingOK: spOp.deletingProcessedOK,
                deletingError: spOp.deletingProcessedError
            });
        }, 500);

        spOp.listDeleteItems(this.props.Url, this.state.list, itemsDelete).then(() => {
            clearInterval(updateDeleteStatus);
            this.setState({
                deletingOK: spOp.deletingProcessedOK,
                deletingError: spOp.deletingProcessedError
            });
        });
    }

    private onCloseDelete() {
        this.setState({
            cancelOperation: true,
            deleting: false,
            deletingModalOpen: false,
        });
        this.loadData();
    }

    /**
     * Click on Refresh. Reload data.
     */
    private clickRefresh() {
        this.loadData();
    }

    /**
     * Update items after changing the selected view
     * @param view 
     */
    private selectedViewUpdate(viewId: string) {
        this.setState({
            view: this.state.availableViews.find(v => v.ID === viewId)
        }, () => {
            //After setting state, load items with new view
            this.loadPartialData();
        });
    }
}