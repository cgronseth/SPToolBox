import * as React from "react";
import { SPUser, SPGroup } from "../sharepoint/spt.sharepoint.entities";
import { List, ListRowProps } from 'react-virtualized'
import { LogAx } from "../spt.logax";
import { Constants } from "../spt.constants";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";

export interface ISitePermUsersProps {
    url: string;
    userFilter: string;
    onSelected: (user: SPUser) => void;
}

export interface ISitePermUsersState {
    loading: boolean;
    usersLoaded: SPUser[];
}

export class SitePermissionsUserList extends React.Component<ISitePermUsersProps, ISitePermUsersState> {
    constructor(props: ISitePermUsersProps) {
        super(props);
        this.state = {
            loading: false,
            usersLoaded: []
        };
    }

    usersFiltered: SPUser[];

    render() {
        this.usersFiltered = this.filterUsers();

        return <div id="userSearchResults">
            <br />
            <div className="subtitle">
                {Constants.getLiteral("directorySearchResults")}
                {
                    this.state.loading &&
                    <img src="icons/ajax-loader-gray.gif" />
                }
            </div>
            {
                !this.state.loading && this.usersFiltered.length > 0 &&
                <div>
                    <div style={{ position: "relative", float: "left", width: 420 }}>
                        <List
                            className="listTable"
                            width={400}
                            height={Math.min(this.usersFiltered.length * 35, 200)}
                            rowCount={this.usersFiltered.length}
                            rowHeight={35}
                            rowRenderer={(props) => this.renderRow(props)}
                        />
                    </div>
                    <div style={{ position: "relative", float: "left", width: 120 }}>
                        <span className="icono35" style={{ backgroundColor: "green", height:24 }}></span>&nbsp;<small>Site User</small><br /><br />
                        <span className="icono35" style={{ backgroundColor: "orange" , height:24 }}></span>&nbsp;<small>Site Admin</small>
                    </div>
                    <div style={{clear:"both"}}></div>
                </div>
            }
        </div>;
    }

    public componentDidUpdate(prevProps: ISitePermUsersProps, prevState: ISitePermUsersState) {
    }

    public componentDidMount(): void {
        this.updateItems();
    }

    private updateItems() {
        let qry: string = SPRest.querySiteGroupsUsers(this.props.url);
        SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
            let usuariosCargados: SPUser[] = [];
            r.value.forEach((group: any) => {
                if (group.Users && group.Users.length > 0) {
                    group.Users.forEach((user: any) => {
                        //Reject data that doesn't have valid title or email
                        if (!user.Email || user.Email === '' || !user.Title || user.Title === '') {
                            return;
                        }

                        let usuarioCargado: SPUser = usuariosCargados.find(u => u.Email === user.Email);
                        if (!usuarioCargado) {
                            usuariosCargados.push({
                                Id: user.Id,
                                DisplayName: user.Title,
                                Email: user.Email,
                                IsSiteAdmin: user.IsSiteAdmin,
                                Groups: []
                            } as SPUser);
                            usuarioCargado = usuariosCargados.find(u => u.Email === user.Email);
                        }
                        usuarioCargado.Groups.push({
                            ID: group.Id,
                            Name: group.Title,
                            PrincipalType: group.PrincipalType
                        } as SPGroup);
                    });
                }
            });

            this.setState({
                loading: false,
                usersLoaded: usuariosCargados.sort((a, b) => (a.DisplayName.toLowerCase() > b.DisplayName.toLowerCase()) ? 1 : -1)
            });
        }, (e) => {
            LogAx.trace("SPT.directory.siteperm.user updateItems error: " + e);
        });

        this.setState({
            loading: true
        });
    }

    private renderRow(props: ListRowProps) {
        let rowStyleClass: string = (props.index % 2 === 0) ? "evenRow" : "oddRow";
        let item: SPUser = this.usersFiltered[props.index];

        return (
            <div key={props.key}
                className={`${rowStyleClass} bodyCell pointer`}
                style={props.style}
                onClick={() => this.props.onSelected(item)}>

                {this.renderIcono(item.Id, item.IsSiteAdmin)}

                <div style={{ position: "relative", float: "left", width: 300, height: 35, marginLeft: 3, overflow: "hidden" }}>
                    <b>{item.DisplayName}</b><br />
                    <i>{item.Email}</i>
                </div>
                <div style={{ clear: "both" }}></div>
            </div>
        );
    }

    private renderIcono(id?: number, siteAdmin?: boolean) {
        let color: string = siteAdmin ? "orange" : "green";
        return (
            <div className="icono35" style={{ backgroundColor: color }}>{id}</div>
        );
    }

    private filterUsers(): SPUser[] {
        let filter = this.props.userFilter.toLowerCase();

        return this.state.usersLoaded.filter(u =>
            u.DisplayName.toLowerCase().indexOf(filter) !== -1 ||
            u.Email.toLowerCase().indexOf(filter) !== -1
        );
    }
}
