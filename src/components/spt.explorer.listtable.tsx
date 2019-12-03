import * as React from "react";
import { SPListItem, SPView, SPItem, SPData } from "../sharepoint/spt.sharepoint.entities";
import { Grid, AutoSizer, GridCellProps, Index, ScrollSync } from 'react-virtualized'
import { LogAx } from "../spt.logax";
import { Constants } from "../spt.constants";
import { CommonUI } from "../commonui/commonui";

export interface IListTableProps {
    listItems: SPListItem[];
    loaded: boolean;
    windowHeight: number;
    windowWidth: number;
    view: SPView;
    onSelected: (itemIDs: number[], clickedItemId: number) => void;
}

export interface IListTableState {
    allChecked: boolean;
    lastClickedItem: number;
    listItems: SPListItem[];
}

export class ListTable extends React.Component<IListTableProps, IListTableState> {

    private headerGrid: any;

    constructor(props: IListTableProps) {
        super(props);
        this.state = {
            allChecked: false,
            lastClickedItem: null,
            listItems: []
        };
    }

    render() {
        let columnCount: number = this.props.view.ViewFields.length + 1;
        let gridHeight: number = this.props.windowHeight - 180;
        let overscanColumnCount: number = 0;
        let overscanRowCount: number = 5;
        let rowHeight: number = 30;
        let rowCount: number = this.state.listItems.length;

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

    public componentDidUpdate(prevProps: IListTableProps, prevState: IListTableState) {
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
        } else {
            // Header field
            let columnIndex = props.columnIndex - 1;
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
        let rowListItem: SPListItem = this.state.listItems[props.rowIndex];
        let rowStyleClass: string = (props.rowIndex % 2 === 0) ? "evenRow" : "oddRow";

        if (props.columnIndex === 0) {
            // Row Checkbox
            return (
                <div key={props.key} className={`${rowStyleClass} cellCheck`} style={props.style}>
                    <img className="cursor-pointer"
                        src={CommonUI.checkRowImage(rowListItem, this.state.lastClickedItem)}
                        onClick={() => this.clickItemRow(props.rowIndex)}
                    ></img>
                </div>
            );
        }
        // Row field
        let item: SPData = rowListItem.Items[props.columnIndex - 1];
        let itemValue: string = item ? item.StringValue : "";
        return (
            <div key={props.key} className={`${rowStyleClass} bodyCell`} style={props.style}>{itemValue}</div>
        );
    }

    private noContentRenderer() {
        return <div className="noCells">{Constants.getLiteral(this.props.loaded ? "explorerNoRows" : "explorerCargando")}</div>;
    }

    private columnWidthCache: { [columnIndex: number]: number } = {};
    private getColumnWidth(params: Index) {
        if (!this.columnWidthCache[params.index]) {
            this.columnWidthCache[params.index] = CommonUI.getColumnTypeWidth(params.index, 1, this.props.view);
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
        this.updateSelectedListItems(items, null);
        this.setState({
            listItems: items,
            allChecked: !this.state.allChecked,
            lastClickedItem: null
        });
    }

    /**
    * Event Item row checkbox clicked
    * @param selectedItem 
    */
    private clickItemRow(selectedRow: number): void {
        let items: SPListItem[] = this.state.listItems;
        let selectedItem: SPListItem = items[selectedRow];
        selectedItem.Checked = selectedItem.Checked <= 0 ? 1 : 0; //Reversed because previous setting changed allChecked value in state
        this.updateSelectedListItems(items, selectedItem.ID);
        this.setState({
            listItems: items,
            allChecked: false,
            lastClickedItem: selectedItem.ID
        });
        LogAx.trace("SPT.explorer.listviewer click item: " + selectedItem.ID);
    }

}