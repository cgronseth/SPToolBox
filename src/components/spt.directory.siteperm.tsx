import * as React from "react";
import { SPList, SPListItem, SPUser, SPFolder, SPView } from "../sharepoint/spt.sharepoint.entities";
import { LogAx } from "../spt.logax";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { Constants } from "../spt.constants";
import { SP } from "../sharepoint/spt.sharepoint";
import { Dates } from "../spt.dates";
import { SPOps } from "../sharepoint/spt.sharepoint.operations";
import { SitePermissionsUserList } from "./spt.directory.siteperm.users";

export interface IDirectorySitePermissionsProps {
    Url: string;
    Title: string;
}

export interface IDirectorySitePermissionsState {
    userSearch: string;
    userSearchValidationClass: string;
    selectedUser: SPUser;
}

export class SitePermissions extends React.Component<IDirectorySitePermissionsProps, IDirectorySitePermissionsState> {
    constructor(props: IDirectorySitePermissionsProps) {
        super(props);
        this.state = {
            userSearch: null,
            userSearchValidationClass: "inherit",
            selectedUser: null
        };
        this.handleUserSearch = this.handleUserSearch.bind(this);
    }

    render() {
        return <div>
            <div className="title">{Constants.getLiteral("directoryPermissionsSite")}</div>
            <div className="subtitle">{Constants.getLiteral("directoryTitleFiltro")}</div>

            <label>{Constants.getLiteral("directorySearchUser")} </label>
            <input type="text" style={{ width: "400px", color: this.state.userSearchValidationClass }}
                value={this.state.userSearch}
                placeholder={Constants.getLiteral("directorySearchPlaceholder")}
                onChange={this.handleUserSearch} />
            {
                this.state.userSearch &&
                <div>
                    <br />
                    <div className="subtitle">{Constants.getLiteral("directorySearchResults")}</div>
                    <SitePermissionsUserList
                        url={this.props.Url}
                        userFilter={this.state.userSearch}
                        onSelected={(user) => this.setState({ selectedUser: user })} />
                </div>
            }
            {
                this.state.selectedUser &&
                <div>
                    <br />
                    <div className="subtitle">{Constants.getLiteral("directoryTitlePermissions")}</div>
                    Selected: {this.state.selectedUser.DisplayName}
                </div>
            }
        </div>;
    }

    public componentDidUpdate(prevProps: IDirectorySitePermissionsProps, prevState: IDirectorySitePermissionsState) {
        if (prevProps.Url !== this.props.Url) {
            this.setState({
                userSearch: null
            }, () => {

            });
        }
    }

    public componentDidMount(): void {
    }

    public componentWillUnmount(): void {
    }

    private handleUserSearch(event: React.FormEvent<HTMLInputElement>) {
        let pattern: RegExp = /^[0-9A-zÀ-úÄ-ü,\s_-]+$/;
        let searchstring: string = event.currentTarget.value;
        if (searchstring === '') {
            this.setState({ userSearch: null, userSearchValidationClass: "inherit" });
        } else {
            let matches = searchstring.match(pattern);
            if (matches && matches.length) {
                this.setState({ userSearch: matches[0], userSearchValidationClass: "inherit" });
            } else {
                this.setState({ userSearch: null, userSearchValidationClass: "red" });
            }
        }
    }
}