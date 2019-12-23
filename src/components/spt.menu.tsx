import * as React from "react";
import { Constants } from "../spt.constants"
import { LogAx } from "../spt.logax";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { Strings } from "../spt.strings";

export interface IMenuState {
    analizando: boolean;
    conectado: boolean;
    sharepointData: any;
}
export class Menu extends React.Component<{}, IMenuState> {
    constructor(props: Readonly<{}>) {
        super(props);
        this.state = {
            analizando: false,
            conectado: false,
            sharepointData: null
        };
        this.clickExplorador = this.clickExplorador.bind(this);
        this.clickDirectory = this.clickDirectory.bind(this);
    }

    private currentUrl: string;
    private isCancelled: boolean;

    render() {
        return <div id="SPT.Menu">
            <div>
                {
                    this.state.analizando &&
                    <div>
                        <img src="icons/ajax-loader.gif" width="12px" /> {Constants.getLiteral("menuAnalizandoSitio")}
                    </div>
                }
                {
                    !this.state.analizando && this.state.conectado &&
                    <div>
                        <div className='SPTMenuDiagnosticsConnected'>{Constants.getLiteral("menuConectado")}</div>
                        <div className='SPTMenuDiagnosticsDetail'>&middot;&nbsp;{Constants.getLiteral("menuTitulo")}: {this.state.sharepointData["Title"]}</div>
                        <div className='SPTMenuDiagnosticsDetail'>&middot;&nbsp;{Constants.getLiteral("menuVersion")}: {this.state.sharepointData["UIVersion"]}</div>
                        <div className='SPTMenuDiagnosticsDetail'>&middot;&nbsp;{Constants.getLiteral("menuLCID")}: {this.state.sharepointData["Language"]}</div>
                    </div>
                }
                {
                    !this.state.analizando && !this.state.conectado &&
                    <div>
                        {Constants.getLiteral("menuNoSharePoint")}
                    </div>
                }
            </div>
            {
                this.state.conectado &&
                <div>
                    <hr />
                    <div className="SPTMenuItem" onClick={this.clickExplorador}>{Constants.getLiteral("menuExplorador")}</div>
                    <div className="SPTMenuItem" onClick={this.clickDirectory}>{Constants.getLiteral("menuDirectory")}</div>
                </div>
            }
            <hr />
            <div className="SPTMenuFooter">SharePoint Toolbox V{browser.runtime.getManifest().version}</div>
        </div>;
    }

    public componentDidUpdate(prevProps: any, prevState: any) {
        if (prevState.analizando != this.state.analizando) {
            let qry = SPRest.querySiteInfo(this.currentUrl);
            LogAx.trace("Query:" + qry);

            SPRest.restQuery(qry, RestQueryType.ODataJSON, 0).then((result: any) => {
                if (this.isCancelled) return;
                try {
                    this.setState({
                        sharepointData: result,
                        analizando: false,
                        conectado: true
                    });
                } catch (e) {
                    this.setState({
                        sharepointData: null,
                        analizando: false,
                        conectado: false
                    });
                }
            }, (e) => {
                // Error mostly because this is not a SHP Site (or no permission)
                if (this.isCancelled) return;
                this.setState({
                    sharepointData: null,
                    analizando: false,
                    conectado: false
                });
            });
        }
    }

    public componentDidMount(): void {
        //Launch extension menú. Detects current tab to obtain possible SHP site
        browser.tabs.query({ active: true, windowId: browser.windows.WINDOW_ID_CURRENT })
            .then(tabs => browser.tabs.get(tabs[0].id))
            .then(tab => {
                console.info(tab);
                this.currentUrl = Strings.getWebUrlFromAbsolute(tab.url);
                this.setState({
                    analizando: true
                });
            });
    }

    public componentWillUnmount(): void {
        this.isCancelled = true;
    }

    private clickExplorador(e: React.MouseEvent<HTMLDivElement, MouseEvent>): void {
        let createData: any = {
            type: "panel",
            titlePreface: "SharePoint Toolbox - ",
            url: "spt.explorer.html?u=" + encodeURIComponent(this.currentUrl)
                + "&i=" + encodeURIComponent(this.state.sharepointData["Id"])
                + "&t=" + encodeURIComponent(this.state.sharepointData["Title"])
                + "&v=" + encodeURIComponent(this.state.sharepointData["menuVersion"])
                + "&l=" + encodeURIComponent(this.state.sharepointData["menuLCID"])
        };
        browser.windows.create(createData)
            .then((window) => {
                console.log("Panel " + window.id + " 'SPT.Explorer' created");
            });

        e.preventDefault();
    }

    private clickDirectory(e: React.MouseEvent<HTMLDivElement, MouseEvent>): void {
        let createData: any = {
            type: "panel",
            titlePreface: "SharePoint Toolbox - ",
            url: "spt.directory.html?u=" + encodeURIComponent(this.currentUrl)
                + "&i=" + encodeURIComponent(this.state.sharepointData["Id"])
                + "&t=" + encodeURIComponent(this.state.sharepointData["Title"])
                + "&v=" + encodeURIComponent(this.state.sharepointData["menuVersion"])
                + "&l=" + encodeURIComponent(this.state.sharepointData["menuLCID"])
        };
        browser.windows.create(createData)
            .then((window) => {
                console.log("Panel " + window.id + " 'SPT.Directory' created");
            });

        e.preventDefault();
    }
}
