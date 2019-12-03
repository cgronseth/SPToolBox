import * as React from "react";
import { SPWeb } from "../sharepoint/spt.sharepoint.entities";
import { LogAx } from "../spt.logax";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { Constants } from "../spt.constants";
import { Cache } from "../spt.cache";

export interface IExplorerWebDetailsProps {
    key: string;
    Url: string;
    Nivel: number;
}

export interface IExplorerWebDetailsState {
    web: SPWeb;
    unauthorized: boolean;
}

export class WebDetails extends React.Component<IExplorerWebDetailsProps, IExplorerWebDetailsState> {
    constructor(props: IExplorerWebDetailsProps) {
        super(props);
        this.state = {
            web: null,
            unauthorized: false
        };
    }

    private readonly cacheCargaExploradorWebs: string = "spt.explorer.webdetails.cargaWebs_";
    private isCancelled: boolean;

    render() {
        let detalleWebRoot = this.props.Nivel === 0 ? "detalleWebRoot" : "detalleWeb";

        return <div>
            {
                this.state.web &&
                <div>
                    <div>
                        <div className={detalleWebRoot} onClick={(e) => this.clickSubsitio(e, this.state.web, this.props.Nivel)}>
                            <div className="detalleWebName"> {this.state.web.Name}</div>
                            {
                                this.state.web.Description &&
                                <div>
                                    <div className="detalleWebLabel">{Constants.getLiteral("generalDescripcion")}</div>
                                    <div className="detalleWebValue">{this.state.web.Description}</div>
                                    <div className="clear"></div>
                                </div>
                            }
                            <div className="detalleWebLabel">{Constants.getLiteral("generalCreado")}</div>
                            <div className="detalleWebValue">{this.state.web.Created}</div>
                            <div className="clear"></div>
                        </div>
                    </div>
                    {
                        this.props.Nivel < 99 && this.state.web.Webs &&
                        <div className="detalleWebNivelSubsitio">
                            {
                                this.state.web.Webs.map((web) =>
                                    <WebDetails key={web.ID} Url={web.Url} Nivel={this.props.Nivel + 1} />
                                )
                            }
                        </div>
                    }
                    {
                        this.props.Nivel == 99 &&
                        <div>{Constants.getLiteral("webViewerDetailsLevelLimit")}</div>
                    }
                </div>
            }
            {
                this.state.unauthorized &&
                <div className="detalleWebUnauthorized">
                    <div className="detalleWebName">{Constants.getLiteral("webViewerDetailsUnauzorized")}</div>
                    <div>
                        <div className="detalleWebLabel">{Constants.getLiteral("gernarlSitio")}</div>
                        <div className="detalleWebValue">{this.props.Url}</div>
                        <div className="clear"></div>
                    </div>
                </div>
            }
            {
                !this.state.web && !this.state.unauthorized &&
                <div>
                    <img src="icons/ajax-loader.gif" width="12px" /> {Constants.getLiteral("webViewerCargando")}
                </div>
            }
        </div>;
    }

    public componentDidUpdate(prevProps: IExplorerWebDetailsProps, prevState: IExplorerWebDetailsState) {
        if (prevProps.Url !== this.props.Url) {
            LogAx.trace("WebDetails DU");
            this.setState({
                web: null
            });
            this.loadWeb();
        }
    }

    public componentDidMount(): void {
        LogAx.trace("WebDetails DM " + this.props.Url);
        this.loadWeb();
    }

    public componentWillUnmount(): void {
        this.isCancelled = true;
    }

    private loadWeb() {
        let cacheKey: string = this.cacheCargaExploradorWebs + this.props.Url;

        if (Cache.Has(cacheKey)) {
            this.setState({
                web: Cache.Get<SPWeb>(cacheKey)
            });
        } else {
            let qry: string = SPRest.queryWeb(this.props.Url);
            LogAx.trace("Query Web:" + qry);
            SPRest.restQuery(qry, RestQueryType.ODataJSON).then((resultWeb: any) => {
                try {
                    if (this.isCancelled) return;

                    qry = SPRest.queryWebs(resultWeb["Url"]);
                    LogAx.trace("Query Webs:" + qry);

                    SPRest.restQuery(qry, RestQueryType.ODataJSON).then((r: any) => {
                        if (this.isCancelled) return;

                        try {
                            let subsitios: SPWeb[] = [];
                            if (r.value) {
                                subsitios = r.value.map((web: any) => ({
                                    ID: web.Id,
                                    Name: web.Title,
                                    Url: web.Url,
                                    ServerRelativeUrl: web.ServerRelativeUrl
                                }) as SPWeb);
                            }

                            let oWeb: SPWeb = {
                                ID: resultWeb["Id"],
                                Name: resultWeb["Title"],
                                Description: resultWeb["Description"],
                                Url: resultWeb["Url"],
                                ServerRelativeUrl: resultWeb["ServerRelativeUrl"],
                                Created: resultWeb["Created"],
                                Webs: subsitios
                            }
                            Cache.PutShort<SPWeb>(cacheKey, oWeb);
                            this.setState({
                                web: oWeb
                            });
                        } catch (e) {
                            LogAx.trace("SPT.WebViewer Query Webs ODataAx.restQuery exception: " + e);
                        }
                    }, (e) => {
                        LogAx.trace("SPT.WebViewer Query Webs ODataAx.restQuery error: " + e);
                    });
                } catch (e) {
                    LogAx.trace("SPT.WebViewer Query Site ODataAx.restQuery exception: " + e);
                }
            }, (e) => {
                LogAx.trace("SPT.WebViewer Query Site ODataAx.restQuery error: " + e);
                this.setState({
                    unauthorized: true
                });
            });
        }
    }

    private clickSubsitio = (e: React.MouseEvent<HTMLDivElement, MouseEvent>, web: SPWeb, nivel: number) => {
        if (nivel > 0) {
            let createData: BrowserCreateData = {
                type: "panel",
                titlePreface: "SharePoint Toolbox - Explorer",
                url: "spt.explorer.html?u=" + encodeURIComponent(web.Url)
                    + "&i=" + encodeURIComponent(web.ID)
                    + "&t=" + encodeURIComponent(web.Name)
                    + "&v=xxx"
                    + "&l=xxx"
            };
            browser.windows.create(createData).then((window) => {
                console.log("Panel " + window.id + " 'Explorer' created");
            }, (e) => {
                console.log("Error creating Panel for '" + createData.titlePreface + "': " + e);
            });
        }
        e.preventDefault();
    };
}

/**
 * Interface should be included in WebExtensions Typings, but meanwhile here is my minimal implementation
 */
export interface BrowserCreateData {
    type: browser.windows.CreateType;
    titlePreface: string;
    url: string;
}