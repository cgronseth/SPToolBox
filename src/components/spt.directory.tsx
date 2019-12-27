import * as React from "react";
import { Constants } from "../spt.constants"
import { Strings } from "../spt.strings";
import { SitePermissions } from "./spt.directory.siteperm";
import { SPSite, SPWeb } from "../sharepoint/spt.sharepoint.entities";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { LogAx } from "../spt.logax";

export interface IDirectoryState {
    showPermissionsSite: boolean;
    site: SPSite;
    web: SPWeb;
}

export class Directory extends React.Component<{}, IDirectoryState> {
    constructor(props: Readonly<{}>) {
        super(props);
        this.state = {
            showPermissionsSite: false,
            site: null,
            web: null
        };
    }

    private currentUrl: string;
    private currentTitle: string;

    render() {
        return <div id="SPT.Directory">
            <div className="panelIzquierdo">
                <div className="tituloSeccion">
                    <span>{Constants.getLiteral("directory")}</span>
                </div>
                <a href="#" onClick={this.clickPermisosSitio}><i className="fas fa-sticky-note"></i> {Constants.getLiteral("directoryPermissionsSite")}</a>
            </div>
            <div className="panelDerecho">
                {
                    this.state.showPermissionsSite && this.state.site && this.state.web &&
                    <SitePermissions
                        Url={this.currentUrl}
                        Title={this.currentTitle}
                        web={this.state.web} />
                }
            </div>
        </div>;
    }

    public componentDidUpdate(prevProps: any, prevState: any) {
    }

    public componentDidMount(): void {
        this.currentUrl = Strings.parseQueryString()["u"];
        this.currentTitle = Strings.parseQueryString()["t"];

        let qryWeb: string = SPRest.queryWeb(this.currentUrl);
        SPRest.restQuery(qryWeb, RestQueryType.ODataJSON, 1).then((w: any) => {
            LogAx.trace("SPT.directory.Directory web result: " + JSON.stringify(w))

            let qrySite: string = SPRest.querySiteInfo(this.currentUrl);
            SPRest.restQuery(qrySite, RestQueryType.ODataJSON, 1).then((s: any) => {
                LogAx.trace("SPT.directory.Directory site result: " + JSON.stringify(s))

                this.setState({
                    web: {
                        ID: w.Id,
                        Name: w.Title,
                        Description: w.Description,
                        Url: w.Url,
                        ServerRelativeUrl: w.ServerRelativeUrl,
                        Created: w.Created
                    },
                    site: {
                        ID: s.Id,
                        Url: s.Url
                    }
                });
            }, (e) => {
                LogAx.trace("SPT.directory.Directory componentDidMount() site error: " + e)
                this.setState({
                    site: null
                });
            });

        }, (e) => {
            LogAx.trace("SPT.directory.Directory componentDidMount() web error: " + e)
            this.setState({
                web: null
            });
        });
    }

    private clickPermisosSitio = (e: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {
        this.setState({
            showPermissionsSite: true
        });
        e.preventDefault();
    };

}
