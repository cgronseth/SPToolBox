import * as React from "react";
import { SPUser, SPWeb, PermissionKind } from "../sharepoint/spt.sharepoint.entities";
import { LogAx } from "../spt.logax";
import { Constants } from "../spt.constants";
import { SP } from "../sharepoint/spt.sharepoint";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { Table, Column, Index, AutoSizer } from "react-virtualized";


export interface ISitePermDetailProps {
    url: string;
    web: SPWeb;
    user: SPUser;
}

export interface ISitePermDetailState {
    loading: boolean;
    permissions: IPermissionWeb[];
}

interface IPermissionBase {
    High: number;
    Low: number;
}

interface IPermissionItem {
    Id: number;
    Title: string;
    Permissions?: IPermissionBase[];
}

interface IPermissionList {
    Id: string;
    Title: string;
    Permissions?: IPermissionBase[];
    Items?: IPermissionItem[];
}

interface IPermissionWeb {
    Url: string;
    RelativeUrl: string;
    Id: string;
    Title: string;
    Level: number;
    Permissions?: IPermissionBase[];
    Lists?: IPermissionList[];
}

interface IRenderTable {
    web?: string;
    list?: string;
    item?: string;
    rd: string;
    wr: string;
    de: string;
}

export class SitePermissionsDetail extends React.Component<ISitePermDetailProps, ISitePermDetailState> {
    constructor(props: ISitePermDetailProps) {
        super(props);
        this.state = {
            loading: false,
            permissions: []
        };
    }

    private renderTable: IRenderTable[] = [];

    render() {
        this.renderTable = this.convertPermissionsToTable();

        return <div id="userPermissionDetails">
            {
                <div>{/*TODO: Virtualized table with columns Site, List, Item, Read, Write, Delete.
                Load async subwebs, async lists, async items that have broken permissions from parent*/}
                    <AutoSizer disableHeight>
                        {({ width }) => (
                            <Table
                                rowClassName={(rowInfo: Index) => this.rowClassName(rowInfo.index)}
                                headerHeight={28}
                                height={200}
                                noRowsRenderer={() => this.noRowRenderer()}
                                overscanRowCount={10}
                                rowHeight={28}
                                rowGetter={(info: Index) => this.renderTable[info.index]}
                                rowCount={this.renderTable.length}
                                width={width}>
                                <Column label={Constants.getLiteral("directoryTableWebColumn")} dataKey="web" width={200} headerClassName="tableHeaderCell" className="tableCell" />
                                <Column label={Constants.getLiteral("directoryTableListColumn")} dataKey="lst" width={140} headerClassName="tableHeaderCell" className="tableCell" />
                                <Column label={Constants.getLiteral("directoryTableItemColumn")} dataKey="itm" width={140} headerClassName="tableHeaderCell" className="tableCell" />
                                <Column label={Constants.getLiteral("directoryTableReadColumn")} dataKey="rd" width={60} headerClassName="tableCenteredHeaderCell" className="tableCenteredCell" />
                                <Column label={Constants.getLiteral("directoryTableWriteColumn")} dataKey="wr" width={60} headerClassName="tableCenteredHeaderCell" className="tableCenteredCell" />
                                <Column label={Constants.getLiteral("directoryTableDeleteColumn")} dataKey="de" width={60} headerClassName="tableCenteredHeaderCell" className="tableCenteredCell" />
                            </Table>
                        )}
                    </AutoSizer>
                </div>
            }
            {
                this.state.loading &&
                <div>
                    <img src="icons/ajax-loader.gif" />&nbsp;
                    <span>{Constants.getLiteral("generalCargando")}</span>
                </div>
            }
        </div>;
    }

    public componentDidUpdate(prevProps: ISitePermDetailProps, prevState: ISitePermDetailState) {
        if (prevProps.user.Email !== this.props.user.Email) {
            this.load();
        }
    }

    public componentDidMount(): void {
        this.load();
    }

    private noRowRenderer() {
        if (!this.state.loading) {
            return <div className="noCells">
                <span>{Constants.getLiteral("generalNoData")}</span>
            </div>;
        }
    }

    private rowClassName(index: number) {
        if (index < 0) {
            return "tableHeaderRow";
        } else {
            return index % 2 === 0 ? "tableEvenRow" : "tableOddRow";
        }
    }

    private convertPermissionsToTable(): IRenderTable[] {
        let table: IRenderTable[] = [];
        this.state.permissions.sort((a, b) => (a.RelativeUrl > b.RelativeUrl) ? 1 : -1).forEach((web) => {
            table.push({
                web: web.RelativeUrl,
                rd: this.checkPermissionKind(web.Permissions, PermissionKind.viewListItems),
                wr: this.checkPermissionKind(web.Permissions, PermissionKind.editListItems),
                de: this.checkPermissionKind(web.Permissions, PermissionKind.deleteListItems)
            } as IRenderTable);
            web.Lists && web.Lists.sort((a, b) => (a.Title > b.Title) ? 1 : -1).forEach((list) => {
                table.push({
                    list: list.Title,
                    rd: this.checkPermissionKind(list.Permissions, PermissionKind.viewListItems),
                    wr: this.checkPermissionKind(list.Permissions, PermissionKind.editListItems),
                    de: this.checkPermissionKind(list.Permissions, PermissionKind.deleteListItems)
                } as IRenderTable)
                list.Items && list.Items.sort((a, b) => (a.Title > b.Title) ? 1 : -1).forEach((item) => {
                    table.push({
                        item: item.Title,
                        rd: this.checkPermissionKind(item.Permissions, PermissionKind.viewListItems),
                        wr: this.checkPermissionKind(item.Permissions, PermissionKind.editListItems),
                        de: this.checkPermissionKind(item.Permissions, PermissionKind.deleteListItems)
                    } as IRenderTable)
                });
            });
        });
        return table;
    }

    private checkPermissionKind(permissions: IPermissionBase[], kind: PermissionKind): string {
        if (permissions === null) {
            return '<img src="icons/ajax-loader.gif" />';
        }

        let hasPermission = "";
        permissions && permissions.forEach(permission => {
            if (SP.checkEffectivePermission(permission.High, permission.Low, kind)) {
                hasPermission = "*";
                return;
            }
        });

        return hasPermission;
    }

    private load() {
        this.setState({
            loading: true,
            permissions: []
        });

        let qry: string = SPRest.queryWeb(this.props.url);
        LogAx.trace("SPT.directory.SitePermissionsDetail load Query: " + qry);
        SPRest.restQuery(qry, RestQueryType.ODataJSON, 0).then((w: any) => {
            LogAx.trace("SPT.directory.SitePermissionsDetail load() web result: " + JSON.stringify(w));

            //Load all webs and subwebs, then proceed to query permissions for web, lists and items
            this.loadWebsRecursive(w).then((webs) => {
                webs.forEach((web) => {
                    setTimeout(() => {
                        this.loadWebPermissions(web);
                    }, Math.floor(Math.random() * 2000) + 500);
                });
            });
        }, (e) => {
            LogAx.trace("SPT.directory.SitePermissionsDetail load() web error: " + e);
            this.setState({
                loading: false,
                permissions: []
            });
        });
    }

    private loadWebsRecursive(w: any): Promise<IPermissionWeb[]> {
        return new Promise<IPermissionWeb[]>((resolve) => {
            let webs: IPermissionWeb[] = [];

            webs.push({
                Id: w.Id,
                Title: w.Title,
                Url: w.Url,
                RelativeUrl: w.ServerRelativeUrl,
                Level: w.ServerRelativeUrl.split("/").length - 1
            } as IPermissionWeb);

            let qry = SPRest.queryWebs(w.Url);
            LogAx.trace("SPT.directory.SitePermissionsDetail loadWebsRecursive Query: " + qry);
            SPRest.restQuery(qry, RestQueryType.ODataJSON, 0).then((subwebs: any) => {
                LogAx.trace("SPT.directory.SitePermissionsDetail loadWebsRecursive() subwebs result: " + JSON.stringify(subwebs));

                if (subwebs.value) {
                    const promises: Promise<IPermissionWeb[]>[] = [];
                    subwebs.value.map((subweb: any) => {
                        promises.push(this.loadWebsRecursive(subweb))
                    });
                    Promise.all(promises).then((promiseResult) => {
                        promiseResult.forEach((resultSubWebs: IPermissionWeb[]) => {
                            //webs = [...new Set([...webs, ...resultSubWebs])]; //Union of arrays
                            webs = webs.concat(resultSubWebs);
                        });
                        resolve(webs);
                    });
                } else {
                    resolve(webs);
                }

            }, (e) => {
                LogAx.trace("SPT.directory.SitePermissionsDetail loadWebsRecursive() error: " + e);
                resolve(webs);
            });
        })
    }

    private loadWebPermissions(web: IPermissionWeb): Promise<void> {
        return new Promise<void>((resolve) => {
            const promises: Promise<IPermissionBase[]>[] = [];

            // Get permissions for user ID
            promises.push(new Promise<IPermissionBase[]>((resolve) => {
                let qry = SPRest.queryWebPermissionsForUser(web.Url, this.props.user.Id);
                LogAx.trace("SPT.directory.SitePermissionsDetail loadWebPermissions user Query: " + qry);
                SPRest.restQuery(qry, RestQueryType.ODataJSON, 0).then((p: any) => {
                    resolve(
                        p.value.map((basePerm: any) => ({
                            High: basePerm.BasePermissions.High,
                            Low: basePerm.BasePermissions.Low
                        } as IPermissionBase))
                    );
                }, (e) => {
                    LogAx.trace("SPT.directory.SitePermissionsDetail loadWebPermissions() user error: " + e);
                    resolve([]);
                });
            }));

            // Get permissions for each group user belongs to
            this.props.user.Groups && this.props.user.Groups.forEach((group) => {
                promises.push(new Promise<IPermissionBase[]>((resolve) => {
                    let qry = SPRest.queryWebPermissionsForUser(web.Url, group.ID);
                    LogAx.trace("SPT.directory.SitePermissionsDetail loadWebPermissions group Query: " + qry);
                    SPRest.restQuery(qry, RestQueryType.ODataJSON, 0).then((p: any) => {
                        resolve(
                            p.value.map((basePerm: any) => ({
                                High: basePerm.BasePermissions.High,
                                Low: basePerm.BasePermissions.Low
                            } as IPermissionBase))
                        );
                    }, (e) => {
                        LogAx.trace("SPT.directory.SitePermissionsDetail loadWebPermissions() group error: " + e);
                        resolve([]);
                    });
                }));
            });

            // Process all base permission requests, user and groups all combined
            Promise.all(promises).then((allBasePermissions) => {
                web.Lists = []; //Empty for now, fill in later
                web.Permissions = [];
                allBasePermissions.forEach((basePermissions) => {
                    basePermissions && basePermissions.forEach((permissions) => {
                        web.Permissions.push(permissions);
                    });
                });

                this.setState({
                    permissions: this.state.permissions.concat(web)
                }, () => {
                    setTimeout(() => {
                        this.loadListsPermissions(web).then(() => {
                            resolve();
                        });
                    }, Math.floor(Math.random() * 500) + 100);
                });
            });

        });
    }

    private loadListsPermissions(web: IPermissionWeb): Promise<void> {
        return new Promise<void>((resolve) => {
            let qry = SPRest.queryListsLight(web.Url);
            LogAx.trace("SPT.directory.SitePermissionsDetail loadListsPermissions Query: " + qry);
            SPRest.restQuery(qry, RestQueryType.ODataJSON, 1).then((p: any) => {
                resolve();
            }, (e) => {
                LogAx.trace("SPT.directory.SitePermissionsDetail loadListsPermissions() error: " + e);
                resolve();
            });

        });
    }
}
