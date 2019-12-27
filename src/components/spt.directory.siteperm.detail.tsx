import * as React from "react";
import { SPListItem, SPData, SPView, SPUser, SPGroup, SPWeb } from "../sharepoint/spt.sharepoint.entities";
import { LogAx } from "../spt.logax";
import { Constants } from "../spt.constants";
import { CommonUI } from "../commonui/commonui";
import { SP } from "../sharepoint/spt.sharepoint";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";

export interface ISitePermDetailProps {
    url: string;
    web: SPWeb;
    user: SPUser;
}

export interface ISitePermDetailState {
    loading: boolean;
    permissions: IPermissions;
}

interface IPermissionBase {
    High: number;
    Low: number;
}

interface IPermissionItem {
    Id: number;
    Title: string;
    Permissions: IPermissionBase;
}

interface IPermissionList {
    Id: string;
    Title: string;
    Permissions: IPermissionBase;
    Items?: IPermissionItem[];
}

interface IPermissionWeb {
    Id: string;
    Title: string;
    Permissions: IPermissionBase;
    Lists?: IPermissionList[];
    SubWebs?: IPermissionWeb[];
}

interface IPermissions {
    UserId: number;
    Admin: boolean;
    Web: IPermissionWeb;
}

export class SitePermissionsDetail extends React.Component<ISitePermDetailProps, ISitePermDetailState> {
    constructor(props: ISitePermDetailProps) {
        super(props);
        this.state = {
            loading: false,
            permissions: null
        };
    }

    render() {
        return <div id="userPermissionDetails">
            {
                !this.state.loading &&
                <div>TODO: Virtualized table with columns Site, List, Item, Read, Write, Delete.
                     Load async subwebs, async lists, async items that have broken permissions from parent</div>
            }
            {
                this.state.loading &&
                <div className="waiting">
                    <img src="icons/ajax-loader.gif" />&nbsp;
                    <span>{Constants.getLiteral("generalCargando")}</span>
                </div>
            }
        </div>;
    }

    public componentDidUpdate(prevProps: ISitePermDetailProps, prevState: ISitePermDetailState) {
        if (prevProps.user.Email !== this.props.user.Email) {
            this.loadWeb();
        }
    }

    public componentDidMount(): void {
        this.loadWeb();
    }

    private loadWeb() {
        let qry: string = SPRest.queryWebPermissionsForUser(this.props.url, this.props.user.Id);
        SPRest.restQuery(qry, RestQueryType.ODataJSON, 1).then((p: any) => {
            LogAx.trace("SPT.directory.SitePermissionsDetail load() result: " + JSON.stringify(p))
            this.setState({
                loading: false,
                permissions: {
                    UserId: this.props.user.Id,
                    Admin: false,
                    Web: {
                        Id: this.props.web.ID,
                        Title: this.props.web.Name,
                        Permissions: {
                            High: p.High,
                            Low: p.Low
                        },
                        Lists: [],
                        SubWebs: []
                    }
                }
            });
        }, (e) => {
            LogAx.trace("SPT.directory.SitePermissionsDetail load() error: " + e)
            this.setState({
                loading: false,
                permissions: null
            });
        });
    }
}
