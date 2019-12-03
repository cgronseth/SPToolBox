import * as React from "react";
import { Constants } from "../spt.constants"
import { WebExStorage, ICopyPasteInstruction } from "../spt.storage";
import { LogAx } from "../spt.logax";
import { SPView, PermissionKind } from "../sharepoint/spt.sharepoint.entities";
import { SPTButton, SPTButtonType } from "./basics/spt.button";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { SP } from "../sharepoint/spt.sharepoint";

export interface IListMenuProps {
    Url: string;
    ListID?: string;
    itemsSelected: number;
    singleFolderSelected?: boolean;
    listType: number;
    views: SPView[];
    onDescargar: () => void;
    onCopiar: () => void;
    onPegar: (modo: boolean) => void;
    onEliminar: () => void;
    onRefresh: () => void;
    onViewUpdate: (viewId: string) => void;
    onCargarExcel?: () => void;
}

interface IListMenuState {
    clipboardActive: ICopyPasteInstruction;
    viewSelected: string;
    permissionRead: boolean;
    permissionWrite: boolean;
    permissionDelete: boolean;
}

export class ListMenu extends React.Component<IListMenuProps, IListMenuState> {
    constructor(props: IListMenuProps) {
        super(props);
        this.state = {
            clipboardActive: null,
            viewSelected: null,
            permissionRead: false,
            permissionWrite: false,
            permissionDelete: false
        };
        this.handleViewChange = this.handleViewChange.bind(this);
    }

    render() {
        return <div id="ListMenu">
            <SPTButton
                type={SPTButtonType.MenuItem}
                show={this.props.itemsSelected > 0}
                enabled={this.state.permissionRead}
                icon={"fas fa-file-download fa-lg"}
                textId={"explorerMenuBotonDescargar"}
                onClick={this.props.onDescargar} />
            <SPTButton
                type={SPTButtonType.MenuItem}
                show={this.props.listType === 0}
                enabled={this.state.permissionWrite}
                icon={"fas fa-file-upload fa-lg"}
                textId={"explorerMenuBotonSubirCSV"}
                onClick={() => this.props.onCargarExcel()} />
            <SPTButton
                type={SPTButtonType.MenuItem}
                show={this.props.itemsSelected > 0}
                enabled={this.state.permissionRead}
                icon={"fas fa-copy fa-lg"}
                textId={"explorerMenuBotonCopiar"}
                onClick={this.props.onCopiar} />
            <SPTButton
                type={SPTButtonType.MenuItem}
                show={this.state.clipboardActive && this.state.clipboardActive.listType === this.props.listType}
                enabled={this.state.permissionWrite}
                icon={"fas fa-paste fa-lg"}
                textId={"explorerMenuBotonPegar"}
                onClick={() => this.props.onPegar(false)} />
            <SPTButton
                type={SPTButtonType.MenuItem}
                show={this.state.clipboardActive && this.state.clipboardActive.listType === this.props.listType &&
                    this.props.listType === 1 &&
                    this.props.singleFolderSelected}
                enabled={this.state.permissionWrite}
                icon={"fas fa-paste fa-lg"}
                textId={"explorerMenuBotonPegarFolder"}
                onClick={() => this.props.onPegar(true)} />
            <SPTButton
                type={SPTButtonType.MenuItem}
                show={this.props.itemsSelected > 0}
                enabled={this.state.permissionDelete}
                icon={"fas fa-trash-alt fa-lg"}
                textId={"explorerMenuBotonEliminar"}
                onClick={() => this.props.onEliminar()} />
            <SPTButton
                type={SPTButtonType.MenuItem}
                show={true}
                enabled={true}
                icon={"fas fa-sync-alt fa-lg"}
                textId={"explorerMenuBotonRefrescar"}
                onClick={() => this.props.onRefresh()} />

            <div className="menuItemDropdown"> {Constants.getLiteral("explorerMenuDesplegableVista")}: &nbsp;
                <select value={this.state.viewSelected} onChange={this.handleViewChange}>
                    {
                        this.props.views.length > 0 && this.props.views.map((view) =>
                            <option className={view.PersonalView ? "personal" : ""} value={view.ID}>{view.Title}</option>
                        )
                    }
                </select>
            </div>
        </div>;
    }

    public componentDidMount(): void {
        this.updateClipboardStatus();
        this.updateButtonPermissions();
    }

    public componentWillUnmount(): void {
        window.clearInterval(this.clipBoardStatusTimeout);
        window.clearInterval(this.permissionsStatusTimeout);
    }

    private clipBoardStatusTimeout: number;
    private updateClipboardStatus() {
        WebExStorage.info().then((result) => {
            if (this.state.clipboardActive !== result) {
                this.setState({
                    clipboardActive: result
                });
            }
            // This is the best solution found to detect C&P in another window instance and show paste buttons
            this.clipBoardStatusTimeout = window.setTimeout(() => {
                this.updateClipboardStatus();
            }, 1200);
        }, (e) => {
            LogAx.trace("Error in ListMenu.updateClipboardStatus Storage.has: " + e);
        });
    }

    private permissionsStatusTimeout: number;
    private updateButtonPermissions() {
        try {
            if (this.props.ListID) {
                let qry: string = SPRest.queryPermissionsList(this.props.Url, this.props.ListID);
                SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
                    this.setState({
                        permissionRead: SP.checkEffectivePermission(r.High, r.Low, PermissionKind.viewListItems),
                        permissionWrite: SP.checkEffectivePermission(r.High, r.Low, PermissionKind.addListItems),
                        permissionDelete: SP.checkEffectivePermission(r.High, r.Low, PermissionKind.deleteListItems),
                    })
                }, (e) => {
                    LogAx.trace("Error in ListMenu.updateButtonPermissions Permissions query: " + e);
                });
            }

            // Every 2 minutes reload permissions
            this.permissionsStatusTimeout = window.setTimeout(() => {
                this.updateButtonPermissions();
            }, 120000);
        } catch (e) {
            LogAx.trace("Error in ListMenu.updateButtonPermissions: " + e);
        }
    }

    private handleViewChange(event: React.ChangeEvent<HTMLSelectElement>) {
        this.setState({
            viewSelected: event.target.value
        });
        this.props.onViewUpdate(event.target.value);
    }

}
