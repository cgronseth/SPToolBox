import * as React from "react";
import { SPListItem, SPData, SPView } from "../sharepoint/spt.sharepoint.entities";
import { Grid, AutoSizer, GridCellProps, Index, ScrollSync } from 'react-virtualized'
import { LogAx } from "../spt.logax";
import { Constants } from "../spt.constants";
import { CommonUI } from "../commonui/commonui";
import { SP } from "../sharepoint/spt.sharepoint";

export interface ILibraryTableProps {
    listItems: SPListItem[];
    loaded: boolean;
    windowHeight: number;
    windowWidth: number;
    view: SPView;
    onSelected: (itemIDs: number[], clickedItemId: number) => void;
}

export interface ILibraryTableState {
    allChecked: boolean;
    listItems: SPListItem[];
}

export class LibraryTable extends React.Component<ILibraryTableProps, ILibraryTableState> {

    private cacheListItems: SPListItem[] = [];
    private headerGrid: any;
    private lastClickedItem: number;

    constructor(props: ILibraryTableProps) {
        super(props);
        this.state = {
            allChecked: false,
            listItems: []
        };
    }

    render() {
        this.cacheListItems = this.getFilteredListItems();
        let columnCount: number = this.props.view.ViewFields.length + 1;
        let gridHeight: number = this.props.windowHeight - 180;
        let overscanColumnCount: number = 0;
        let overscanRowCount: number = 5;
        let rowHeight: number = 30;
        let rowCount: number = this.cacheListItems.length;

        return <div id="libraryTable">
            <ScrollSync>
                {({ clientHeight, clientWidth, onScroll, scrollHeight, scrollLeft, scrollTop, scrollWidth }) => (
                    <AutoSizer disableHeight>
                        {({ width }) => (
                            <div>
                                <div>
                                    <Grid
                                        cellRenderer={(props) => this.renderHeaderCell(props)}
                                        className="headerGrid"
                                        columnWidth={(params: Index) => this.getColumnWidth(params)}
                                        columnCount={columnCount}
                                        height={rowHeight}
                                        overscanColumnCount={overscanColumnCount}
                                        rowHeight={rowHeight}
                                        rowCount={1}
                                        width={width}
                                        scrollLeft={scrollLeft}
                                        ref={(r) => this.setGridReference(r)}
                                    />
                                </div>
                                <div>
                                    <Grid
                                        cellRenderer={(props) => this.renderBodyCell(props)}
                                        className="bodyGrid"
                                        columnWidth={(params: Index) => this.getColumnWidth(params)}
                                        columnCount={columnCount}
                                        height={gridHeight}
                                        noContentRenderer={() => this.noContentRenderer()}
                                        overscanColumnCount={overscanColumnCount}
                                        overscanRowCount={overscanRowCount}
                                        rowHeight={rowHeight}
                                        rowCount={rowCount}
                                        width={width}
                                        onScroll={onScroll}
                                    />
                                </div>
                            </div>
                        )}
                    </AutoSizer>
                )}
            </ScrollSync>
        </div>;
    }

    public setGridReference(ref: any) {
        this.headerGrid = ref
    }

    public componentDidUpdate(prevProps: ILibraryTableProps, prevState: ILibraryTableState) {
        if (prevProps.listItems.length !== this.props.listItems.length ||
            prevProps.view.ID !== this.props.view.ID) {
            this.updateDisplayListItems();
        }
    }

    public componentDidMount(): void {
        this.updateDisplayListItems();
    }

    private updateDisplayListItems() {
        this.setState({
            listItems: this.props.listItems,
            allChecked: false
        });
        this.columnWidthCache = {};
        this.headerGrid.recomputeGridSize({ columnIndex: 1, rowIndex: 0 });
    }

    private getFilteredListItems(): SPListItem[] {
        return this.state.listItems.filter(i => !i.Hidden);
    }

    private updateSelectedListItems(items: SPListItem[], clickedID: number) {
        this.props.onSelected(
            items.filter(li => li.Checked === 1).map(li => li.ID),
            clickedID
        );
    }

    private renderHeaderCell(props: GridCellProps) {
        if (props.columnIndex === 0) {
            // Header Checkbox
            return (
                <div key={props.key} className="cellCheck" style={props.style}>
                    <img className="cursor-pointer"
                        src={CommonUI.checkHeaderImage(this.state.allChecked)}
                        onClick={() => this.clickHeaderRow()}
                    ></img>
                </div>
            );
        } else if (props.columnIndex === 1) {
            // Header Folder stucture
            return (
                <div key={props.key} className="headerCell" style={props.style}>
                    {Constants.getLiteral("libraryViewerFolder")}
                </div>
            );
        } else {
            // Header field
            let columnIndex = props.columnIndex - 2;
            return (
                <div key={props.key} className="headerCell" style={props.style}>
                    <div>{this.props.view.ViewFields[columnIndex].Title}</div>
                    {
                        //Show internal if different
                        this.props.view.ViewFields[columnIndex].Title !== this.props.view.ViewFields[columnIndex].InternalName &&
                        <div className="internal">{this.props.view.ViewFields[columnIndex].InternalName}</div>
                    }
                </div>
            );
        }
    }

    private renderBodyCell(props: GridCellProps) {
        let rowListItem: SPListItem = this.cacheListItems[props.rowIndex];
        let rowStyleClass: string = (props.rowIndex % 2 === 0) ? "evenRow" : "oddRow";
        let rowFolderClass: string = (!!rowListItem.Folder.UniqueId) ? "folderRow" : "";

        if (props.columnIndex === 0) {
            // Row Checkbox
            return (
                <div key={props.key} className={`${rowStyleClass} ${rowFolderClass} cellCheck`} style={props.style}>
                    <img className="cursor-pointer"
                        src={CommonUI.checkRowImage(rowListItem, this.lastClickedItem)}
                        onClick={() => this.clickItemRow(props.rowIndex)}
                    ></img>
                </div>
            );

        } else if (props.columnIndex === 1) {
            // Row Folder structure
            let folderLevel = rowListItem.Folder.Level - 1;
            let levelStyle: React.CSSProperties = props.style;

            try {
                levelStyle["paddingLeft"] = (CommonUI.initialSeparation + (CommonUI.levelSeparation * folderLevel)) + "px";
                levelStyle["backgroundPosition"] = (folderLevel > 0 ? (CommonUI.levelSeparation * folderLevel) : -2) + "px 0px";
            } catch (e) {
                //Bug in React. When scrolling fast sometimes throws: TypeError: "paddingLeft" is read-only.
                //Try removing someday and checking if fixed.
            }

            // Folder  
            let folderDisplay: any[] = [];
            if (rowListItem.Folder.UniqueId) {
                // Expand/Contract icons
                let expand = rowListItem.Folder.Expand;
                folderDisplay.push(<i key={"401_" + rowListItem.ID}
                    className={`fas fa-caret-${expand ? "down" : "right"} cursor-pointer`}
                    onClick={(e) => this.clickExpandFolderRow(e, props.rowIndex, !expand)}></i>
                );
            }

            return (
                <div key={props.key} className={`${rowStyleClass} ${rowFolderClass} bodyCell folderHeirarchy`} style={levelStyle}>
                    {folderDisplay}
                </div>
            );
        }
        // Row field
        let item: SPData = rowListItem.Items[props.columnIndex - 2];
        let itemValue: string = item ? item.StringValue : "";
        return (
            <div key={props.key} className={`${rowStyleClass} ${rowFolderClass} bodyCell`} style={props.style}>{itemValue}</div>
        );
    }

    private noContentRenderer() {
        return <div className="noCells">{Constants.getLiteral(this.props.loaded ? "explorerNoRows" : "explorerCargando")}</div>;
    }

    private columnWidthCache: { [columnIndex: number]: number } = {};
    private getColumnWidth(params: Index) {
        if (!this.columnWidthCache[params.index]) {
            this.columnWidthCache[params.index] = CommonUI.getColumnTypeWidth(params.index, 2, this.props.view, SP.getFolderDepthLevel(this.state.listItems));
        }
        return this.columnWidthCache[params.index];
    }

    /**
    * Event Header row checkbox clicked
    */
    private clickHeaderRow(): void {
        let items: SPListItem[] = this.state.listItems;
        let setChecked: number = this.state.allChecked ? 0 : 1; //Reversed because previous setting changed allChecked value in state
        for (let item of items) {
            item.Checked = setChecked;
        }
        this.lastClickedItem = null;
        this.updateSelectedListItems(items, null);
        this.setState({
            listItems: items,
            allChecked: !this.state.allChecked
        });
    }

    /**
    * Event Item row checkbox clicked
    * @param selectedItem 
    */
    private clickItemRow(selectedRow: number): void {
        let items: SPListItem[] = this.state.listItems;
        let selectedItem: SPListItem = this.getFilteredListItems()[selectedRow];
        let setChecked: number = selectedItem.Checked <= 0 ? 1 : 0; //Reversed because previous setting changed allChecked value in state
        this.setItemCheckedRecursiveDown(items, selectedItem, setChecked);
        //this.setItemCheckedRecursiveUp(items, selectedItem, setChecked);
        this.updateSelectedListItems(items, selectedItem.ID);
        this.lastClickedItem = selectedItem.ID;
        this.setState({
            listItems: items,
            allChecked: false
        });
        LogAx.trace("SPT.explorer.libraryviewer click item: " + selectedItem.Name);
    }

    /**
    * Marca elemento con valor "checked" y mira si es carpeta y realiza la misma acción a todos los hijos
    * @param items 
    * @param selectedItem 
    * @param checked 
    */
    private setItemCheckedRecursiveDown(items: SPListItem[], selectedItem: SPListItem, checked: number): void {
        items.find(i => i.ID === selectedItem.ID).Checked = checked;
        let itemsHijos: SPListItem[] = items.filter(i => i.Folder.ParentFolder.UniqueId === selectedItem.Folder.UniqueId);
        /*for (let idx = 0; idx < itemsHijos.length; idx++) {
            this.setItemCheckedRecursiveDown(items, itemsHijos[idx], checked);
        }*/
        for (let itemHijo of itemsHijos) {
            this.setItemCheckedRecursiveDown(items, itemHijo, checked);
        }
    }

    /**
     * Busca si en niveles superiores existe una carpeta y refresca su estatus de si está desmarcado, marcado o parcial
     * @param items 
     * @param selectedItem 
     * @param checked 
     */
    /*private setItemCheckedRecursiveUp(items: SPListItem[], selectedItem: SPListItem, checked: number): void {
        let itemsPadre: SPListItem[] = items.filter(i => selectedItem.Folder.ParentFolder.UniqueId === i.Folder.UniqueId);
        for (let idx = 0; idx < itemsPadre.length; idx++) {
            let itemPadre: SPListItem = itemsPadre[idx];
            let checkIgualHijos = this.checkChildItemsCheckedValueRecursive(items, itemPadre, checked === 1);
            itemPadre.Checked = checkIgualHijos ? checked : -1;

            this.setItemCheckedRecursiveUp(items, itemPadre, checked);
        }
    }*/

    /**
     * Busca recursivamente si un elemento tiene un valor "checked" distinto
     * @param items 
     * @param selectedItem 
     * @param checked 
     */
    /*private checkChildItemsCheckedValueRecursive(items: SPListItem[], selectedItem: SPListItem, checked: boolean): boolean {
        let subitems: SPListItem[] = items.filter(i => i.Folder.ParentFolder.UniqueId === selectedItem.Folder.UniqueId);
        for (const item of subitems) {
            if (item.Checked === -1 || (item.Checked === 0 && checked) || (item.Checked === 1 && !checked)) {
                return false;
            }
            if (item.Folder) {
                let resultadoCheckCarpeta: boolean = this.checkChildItemsCheckedValueRecursive(subitems, item, checked);
                if (!resultadoCheckCarpeta) {
                    return false;
                }
            }
        }
        return true;
    }*/

    /**
     * Expand current folder and search child elements to expand or contract
     * @param e 
     * @param selectedRow
     * @param expand
     */
    private clickExpandFolderRow(e: React.MouseEvent<HTMLElement, MouseEvent>, selectedRow: number, expand: boolean): void {
        let selectedItem: SPListItem = this.getFilteredListItems()[selectedRow];
        let items: SPListItem[] = this.state.listItems;
        let itemsAfectados: number = this.setItemVisibilityRecursive(items, selectedItem, expand);
        items.find(i => i.ID === selectedItem.ID).Folder.Expand = expand;

        this.setState({
            listItems: items
        });

        LogAx.trace("SPT.explorer.libraryviewer despliegue de carpeta: " + selectedItem.Name + " (" + itemsAfectados + " items)");
        e.preventDefault();
    }

    /**
     * Search child elements and set their expand value to collapse or expand as necessary
     * @param items 
     * @param selectedItem 
     * @param visible 
     */
    private setItemVisibilityRecursive(items: SPListItem[], selectedItem: SPListItem, visible: boolean): number {
        let itemsAfectados: number = 0;
        let itemsHijos: SPListItem[] = items.filter(i => i.Folder.ParentFolder.UniqueId === selectedItem.Folder.UniqueId);
        /*for (let idx = 0; idx < itemsHijos.length; idx++) {
            let itemHijo: SPListItem = itemsHijos[idx];*/
        for (let itemHijo of itemsHijos) {
            itemHijo.Hidden = !visible;
            itemsAfectados++;
            // Set child visibility with following rules:
            // - If action is to collapse, all child folders and items set to hidden
            // - If action is to expand, inmediate child folders and items are shown, and next level depends on folder expand value
            if (itemHijo.Folder.UniqueId) {
                if (!visible) {
                    itemsAfectados += this.setItemVisibilityRecursive(items, itemHijo, false);
                } else {
                    if (itemHijo.Folder.Expand === true) {
                        itemsAfectados += this.setItemVisibilityRecursive(items, itemHijo, true);
                    }
                }
            }
        }
        return itemsAfectados;
    }

}
