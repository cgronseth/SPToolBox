import * as React from "react";
import { SPList, SPListItem, SPUser, SPFolder, SPView } from "../sharepoint/spt.sharepoint.entities";
import { LogAx } from "../spt.logax";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { Constants } from "../spt.constants";
import { SP } from "../sharepoint/spt.sharepoint";
import { ListMenu } from "./spt.explorer.menu";
import { ZIP } from "../spt.zip";
import { SPTModalDownloadZip, SPTModalCopyLibrary, SPTModalDeleteLibrary } from "./spt.modal";
import { Dates } from "../spt.dates";
import { ListPaging } from "./spt.explorer.paging";
import { LibraryTable } from "./spt.explorer.librarytable";
import { WebExStorage, ICopyPasteInstruction, IListItemLight } from "../spt.storage";
import { SPOps } from "../sharepoint/spt.sharepoint.operations";

export interface IExplorerLibraryProps {
    Url: string;
    ID: string;
    Nivel: number;
}

export interface IExplorerLibraryState {
    library: SPList;
    availableViews: SPView[];
    view: SPView;
    listItems: SPListItem[];
    listItemsLoaded: boolean;
    downloading: boolean;
    downloadingCancel: boolean;
    downloadingCountTotal: number;
    downloadingCount: number;
    downloadingError: number;
    downloadingUrlBlob: string;
    downloadingBytesTotal: number;
    downloadingBytesCurrent: number;
    compressing: boolean;
    compressingPercentage: number;
    singleFolderSelected: SPFolder;
    copyingModalOpen: boolean;
    copyingIntoSingleFolder: boolean;
    copyingCancel: boolean;
    copying: boolean;
    copyingOK: number;
    copyingError: number;
    copyingTotal: number;
    deletingModalOpen: boolean;
    deleting: boolean;
    deletingCancel: boolean;
    deletingOK: number;
    deletingError: number;
    deletingTotal: number;
    storageSize: number;
    windowWidth: number;
    windowHeight: number;
}

export class LibraryViewer extends React.Component<IExplorerLibraryProps, IExplorerLibraryState> {
    constructor(props: IExplorerLibraryProps) {
        super(props);
        this.state = {
            library: null,
            availableViews: [],
            view: null,
            listItems: [],
            listItemsLoaded: false,
            downloading: false,
            downloadingCancel: false,
            downloadingCount: 0,
            downloadingError: 0,
            downloadingCountTotal: 0,
            downloadingBytesTotal: 0,
            downloadingBytesCurrent: 0,
            downloadingUrlBlob: null,
            compressing: false,
            compressingPercentage: 0,
            singleFolderSelected: null,
            copyingModalOpen: false,
            copyingIntoSingleFolder: false,
            copyingCancel: false,
            copying: false,
            copyingOK: 0,
            copyingError: 0,
            copyingTotal: 0,
            deletingModalOpen: false,
            deleting: false,
            deletingCancel: false,
            deletingOK: 0,
            deletingError: 0,
            deletingTotal: 0,
            storageSize: 0,
            windowWidth: 0,
            windowHeight: 0
        };
        this.onDescargar = this.onDescargar.bind(this);
        this.onCancelDescargar = this.onCancelDescargar.bind(this);
        this.onCloseDescargar = this.onCloseDescargar.bind(this);
        this.onAcceptCopyClause = this.onAcceptCopyClause.bind(this);
        this.onCloseCopy = this.onCloseCopy.bind(this);
        this.onAcceptDeleteClause = this.onAcceptDeleteClause.bind(this);
        this.onCloseDelete = this.onCloseDelete.bind(this);
    }

    private onCancel: boolean = false;
    private checkedListItemsIDs: number[] = [];
    private cacheFolders: Map<string, SPFolder> = new Map<string, SPFolder>();

    render() {
        return <div>
            {
                this.state.library && this.state.view &&
                <div id="libraryViewer">
                    <div id="libraryTitle">{this.state.library.Title}
                        <span>[{Constants.getLiteral("explorerBibliotecaInternal")}: {this.state.library.InternalName}]</span>
                        <span>[{Constants.getLiteral("explorerBibliotecaEntityType")}: {this.state.library.ListItemEntityTypeFullName}]</span><br />
                    </div>
                    <ListMenu
                        Url={this.props.Url}
                        ListID={this.props.ID}
                        itemsSelected={this.checkedListItemsIDs.length}
                        singleFolderSelected={(this.state.singleFolderSelected) ? true : false}
                        listType={1}
                        views={this.state.availableViews}
                        onDescargar={() => this.onPrepararZipDescargable()}
                        onCopiar={() => this.clickCopiar()}
                        onPegar={(mode) => this.clickPegar(mode)}
                        onEliminar={() => this.clickDelete()}
                        onRefresh={() => this.clickRefresh()}
                        onViewUpdate={(viewId) => this.selectedViewUpdate(viewId)} />
                    <LibraryTable
                        listItems={this.state.listItems}
                        loaded={this.state.listItemsLoaded}
                        windowHeight={this.state.windowHeight}
                        windowWidth={this.state.windowWidth}
                        view={this.state.view}
                        onSelected={(selectedItemIDs, clickedID) => this.updateSelectedItems(selectedItemIDs, clickedID)} />
                    <ListPaging
                        AvailableHeight={this.state.windowHeight}
                        TotalItems={this.state.library.ItemCount}
                        LoadedItems={this.state.listItems.length}
                        onLoadFullList={() => this.loadFullData()} />
                </div>
            }
            {
                !this.state.library &&
                <div>
                    <img src="icons/ajax-loader.gif" width="12px" /> {Constants.getLiteral("libraryViewerCargando")}
                </div>
            }
            {/*<small>
                <span>View: </span>{this.state.windowWidth}x{this.state.windowHeight}<br />
                <span>Storage: </span>{Strings.closestByteMetric(this.state.storageSize)} <span>(Max 5MB)</span>
            </small>*/}

            <SPTModalDownloadZip
                Open={this.state.downloading}
                DownloadTotalCount={this.state.downloadingCountTotal}
                DownloadOkCount={this.state.downloadingCount}
                DownloadErrorCount={this.state.downloadingError}
                DownloadTotalBytes={this.state.downloadingBytesTotal}
                DownloadCurrentBytes={this.state.downloadingBytesCurrent}
                CompressPercentage={this.state.compressingPercentage}
                onDownload={this.onDescargar}
                onClose={this.onCloseDescargar}
                onCancel={this.onCancelDescargar} />

            <SPTModalCopyLibrary
                Open={this.state.copyingModalOpen}
                FilesOK={this.state.copyingOK}
                FilesError={this.state.copyingError}
                FilesTotal={this.state.copyingTotal}
                CopyInFolder={this.state.copyingIntoSingleFolder}
                onAcceptClause={this.onAcceptCopyClause}
                onCloseModal={this.onCloseCopy} />

            <SPTModalDeleteLibrary
                Open={this.state.deletingModalOpen}
                FilesOK={this.state.deletingOK}
                FilesError={this.state.deletingError}
                FilesTotal={this.state.deletingTotal}
                onAcceptClause={this.onAcceptDeleteClause}
                onCloseModal={this.onCloseDelete} />
        </div>;
    }

    public componentDidUpdate(prevProps: IExplorerLibraryProps, prevState: IExplorerLibraryState) {
        if (prevProps.ID !== this.props.ID) {
            this.setState({
                library: null,
                availableViews: [],
                view: null,
                listItems: [],
                listItemsLoaded: false
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
        this.onCancel = true;
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
        SPOps.loadList(this.props.Url, this.props.ID).then((library: SPList) => {
            this.setState({ library: library });
            SPOps.loadFields(this.props.Url, library.ID).then((fields) => {
                SPOps.loadViews(this.props.Url, library.ID, fields).then((views) => {
                    this.setState({
                        availableViews: views,
                        view: this.state.view ? this.state.view : views.find(v => v.DefaultView === true)
                    }, () => {
                        //After setting state, run initial load with default view
                        this.loadPartialData();
                    });
                }, (e) => {
                    LogAx.trace("SPT.Explorer.LibraryViewer loadViews error: " + e);
                    this.setState({ view: null });
                });
            }, (e) => {
                LogAx.trace("SPT.Explorer.LibraryViewer loadFields error: " + e);
            });
        }, (e) => {
            LogAx.trace("SPT.Explorer.LibraryViewer loadList error: " + e);
            this.setState({ library: null });
        });
    }

    private loadPartialData() {
        this.setState({
            listItems: [],
            listItemsLoaded: false,
            copyingIntoSingleFolder: false,
            singleFolderSelected: null
        });
        this.checkedListItemsIDs = [];

        if (!this.state.library.ID) {
            LogAx.trace("SPT.Explorer.LibraryViewer loadPartialData error: Library ID not set");
            return;
        }
        if (!this.state.view) {
            LogAx.trace("SPT.Explorer.LibraryViewer loadPartialData error: View not set");
            return;
        }

        SPOps.loadItems(this.props.Url, this.state.library.ID, this.state.view, true).then((dataSet: any) => {
            // Initialize cache with root folder
            this.cacheFolders.set(this.state.library.RootFolder.UniqueId, this.state.library.RootFolder);
            // Save item data to collection
            this.loadItemData(dataSet);
        }, (e) => {
            LogAx.trace("SPT.Explorer.LibraryViewer loadPartialData error: " + e);
        });
    }

    private loadFullData(): Promise<void> {
        return new Promise((resolve) => {
            this.setState({
                listItems: []
            });

            let qry: string = SPRest.queryLibraryItemsWithView(this.props.Url, this.props.ID, this.state.view);
            LogAx.trace("Query FULL Library:" + qry);

            this.loadFullDataRecursive(qry).then(() => {
                resolve();
            });
        });
    }

    private loadFullDataRecursive(qry: string): Promise<void> {
        return new Promise<void>((resolve) => {
            SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
                this.loadItemData(r);

                if (!r["odata.nextLink"]) {
                    resolve();
                }
                this.loadFullDataRecursive(r["odata.nextLink"]).then(() => {
                    resolve();
                });
            }, (e) => {
                LogAx.trace("SPT.explorer.LibraryViewer loadFullDataRecursive error: " + e)
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
        let relativeFolderLevel: number = this.state.library.RootFolder.ServerRelativeUrl.split('/').length;

        dataResult.value.forEach((i: any) => {
            let item: SPListItem = new SPListItem;

            // Process entity data
            item.ID = i["ID"] as number;
            item.Name = i["FileLeafRef"];
            item.SPFileSystemObjectType = i["SPFileSystemObjectType"];
            item.Created = i["Created"] as Date;
            item.Author = new SPUser(i.Author.Title, i.Author.EMail);
            item.Modified = i["Modified"] as Date;
            item.Editor = new SPUser(i.Editor.Title, i.Editor.EMail);
            if (i.Folder) {
                item.Folder = {
                    UniqueId: i.Folder.UniqueId,
                    Name: i.Folder.Name,
                    ItemCount: i.Folder.ItemCount as number,
                    ServerRelativeUrl: i.Folder.ServerRelativeUrl,
                    Level: ((i.Folder.ServerRelativeUrl as string).split('/').length - relativeFolderLevel),
                    Expand: true
                };
                if (i.Folder.ParentFolder) {
                    item.Folder.ParentFolder = this.cacheFolders.get(i.Folder.ParentFolder.UniqueId);
                }
                this.cacheFolders.set(item.Folder.UniqueId as string, item.Folder);
            } else {
                let fileref: string = (i.FileRef as string);
                fileref = fileref.substring(0, fileref.lastIndexOf('/'));
                item.Folder = {
                    UniqueId: null,
                    Name: null,
                    Level: ((i.FileRef as string).split('/').length - relativeFolderLevel),
                    ParentFolder: SP.findFolderByPath(this.cacheFolders, fileref),
                    Expand: true
                };
                item.Length = +i.File.Length;
            }

            // All usefull data in mapped collection for visualization access
            for (let field of this.state.view.ViewFields) {
                item.Items.push(SP.parseItemJsonResult(i, field));
            }

            batchListItems.push(item);
        });

        let tempListItems: SPListItem[] = this.state.listItems.concat(batchListItems);

        this.setState({
            listItems: SP.orderLibraryDataByFields(tempListItems, ["Folder", "ID"]),
            listItemsLoaded: true
        });
    }

    /**
     * Click on menu option to start downloading data and create zip archive
     */
    private onPrepararZipDescargable() {
        let items: SPListItem[] = this.state.listItems.filter(i => this.checkedListItemsIDs.includes(i.ID) && !i.Folder.Name);

        this.setState({
            downloading: true,
            downloadingCancel: false,
            downloadingCountTotal: items.length,
            downloadingCount: 0,
            downloadingError: 0,
            downloadingUrlBlob: null,
            downloadingBytesCurrent: 0,
            downloadingBytesTotal: items.map(i => i.Length).reduce((a, b) => (+a) + (+b)),
            compressing: false,
            compressingPercentage: 0
        });

        let zipped = new ZIP(this.props.Url, this.state.library);

        //This timer keeps modal updated with the data saved during the async method
        let updateZippedStatus: number = window.setInterval(() => {
            zipped.cancelOperation = this.state.downloadingCancel;
            this.setState({
                downloadingCount: zipped.downloadProcessedItems,
                downloadingError: zipped.downloadErrorItems,
                downloadingBytesCurrent: zipped.downloadCurrentBytes,
                compressingPercentage: zipped.compressProcessedPercentage
            });
        }, 500);

        zipped.createDownloadZip(items).then((urlBlob: string) => {
            this.setState({
                downloadingCount: zipped.downloadProcessedItems,
                downloadingError: zipped.downloadErrorItems,
                downloadingBytesCurrent: zipped.downloadCurrentBytes,
                compressingPercentage: zipped.compressProcessedPercentage,
                downloadingUrlBlob: urlBlob
            });
            clearInterval(updateZippedStatus);
        });
    }

    /**
     * Click on Download in modal. Proceed to download the zip archive
     */
    private onDescargar() {
        browser.downloads.download({
            url: this.state.downloadingUrlBlob,
            filename: this.state.library.Title + " " + Dates.getFileSuffix() + ".zip",
            saveAs: true
        }).finally(() => {
            this.setState({
                downloading: false
            });
        });
    }

    private onCancelDescargar() {
        this.setState({
            downloadingCancel: true,
            downloading: false
        });
    }

    private onCloseDescargar() {
        this.setState({
            downloading: false
        });
    }

    /**
     * Click on Copy. Add IDs to internal web storage.
     */
    private clickCopiar() {
        let itemData: IListItemLight[] = this.state.listItems
            .filter(i => this.checkedListItemsIDs.includes(i.ID))
            .map(li => ({
                ID: li.ID,
                Author: li.Author.Email,
                Editor: li.Editor.Email,
                Created: li.Created,
                Modified: li.Modified,
                File: {
                    FileName: li.Name,
                    FilePath: SP.getFolderPathWithoutParent(li.Folder.ParentFolder, ""),
                    Length: li.Length
                }
            }) as IListItemLight);
        if (itemData.length) {
            let cpInstruction: ICopyPasteInstruction = {
                site: this.props.Url,
                listId: this.props.ID,
                listType: 1,
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
        this.setState({
            copyingModalOpen: true,
            copyingIntoSingleFolder: mode,
            copying: false,
            copyingCancel: false,
            copyingOK: 0,
            copyingError: 0,
            copyingTotal: 0
        });
    }

    private onAcceptCopyClause() {
        WebExStorage.get().then((cpi: ICopyPasteInstruction) => {
            let spOp: SPOps = new SPOps();
            spOp.copyIntoFolder = this.state.copyingIntoSingleFolder;
            spOp.copyIntoFolderTarget = this.state.singleFolderSelected;

            this.setState({
                copying: true,
                copyingTotal: cpi.items.length
            });

            //This timer keeps modal updated
            let updateCopyPasteStatus: number = window.setInterval(() => {
                spOp.cancelOperation = this.state.copyingCancel;
                this.setState({
                    copyingOK: spOp.copyingProcessedOK,
                    copyingError: spOp.copyingProcessedError
                });
            }, 500);

            spOp.libraryCopyPaste(cpi.site, cpi.listId, this.props.Url, this.state.library, cpi.items).then(() => {
                clearInterval(updateCopyPasteStatus);
                this.setState({
                    copyingOK: spOp.copyingProcessedOK,
                    copyingError: spOp.copyingProcessedError
                });
            });
        }, (e) => {
            LogAx.trace("Error in onAcceptCopyClause WebStorage Get: " + e);
            this.onCloseCopy();
        });
    }

    private onCloseCopy() {
        this.setState({
            copyingCancel: true,
            copying: false,
            copyingModalOpen: false,
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
            deletingCancel: false,
            deletingOK: 0,
            deletingError: 0,
            deletingTotal: 0
        });
    }

    private onAcceptDeleteClause() {
        let itemsDelete: SPListItem[] = this.state.listItems.filter(i => this.checkedListItemsIDs.includes(i.ID));

        let spOp: SPOps = new SPOps();
        this.setState({
            deleting: true,
            deletingTotal: itemsDelete.length
        });

        //This timer keeps modal updated
        let updateDeleteStatus: number = window.setInterval(() => {
            spOp.cancelOperation = this.state.deletingCancel;
            this.setState({
                deletingOK: spOp.deletingProcessedOK,
                deletingError: spOp.deletingProcessedError
            });
        }, 500);

        spOp.libraryDeleteItems(this.props.Url, this.state.library, itemsDelete).then(() => {
            clearInterval(updateDeleteStatus);
            this.setState({
                deletingOK: spOp.deletingProcessedOK,
                deletingError: spOp.deletingProcessedError
            });
        });
    }

    private onCloseDelete() {
        this.setState({
            deletingCancel: true,
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
     * Handle item selection updates from child table component, mostly used by the menu
     * @param selectedItemIDs 
     * @param clickedID 
     */
    private updateSelectedItems(selectedItemIDs: number[], clickedID: number) {
        this.checkedListItemsIDs = selectedItemIDs;
        let clickedFolderItem: SPListItem = this.state.listItems.find(li => clickedID === li.ID && selectedItemIDs.includes(li.ID) && li.Folder.UniqueId);

        this.setState({
            singleFolderSelected: clickedFolderItem ? clickedFolderItem.Folder : null
        });
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