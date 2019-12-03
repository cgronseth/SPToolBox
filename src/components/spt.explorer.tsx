import * as React from "react";
import { Constants } from "../spt.constants"
import { LogAx } from "../spt.logax";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { SPList, SPWeb } from "../sharepoint/spt.sharepoint.entities";
import { ListViewer } from "./spt.explorer.listviewer";
import { LibraryViewer } from "./spt.explorer.libraryviewer";
import { WebViewer } from "./spt.explorer.webviewer";
import { Strings } from "../spt.strings";

export interface IExplorerState {
    loading: boolean;
    resultadosListas: SPList[];
    resultadosBibliotecas: SPList[];
    idListaSeleccionada: string;
    idBibliotecaSeleccionada: string;
    subsitioSeleccionado: boolean;
    showHiddenLibraries: boolean;
    showHiddenLists: boolean;
}

export class Explorer extends React.Component<{}, IExplorerState> {
    constructor(props: Readonly<{}>) {
        super(props);
        this.state = {
            loading: true,
            resultadosListas: null, resultadosBibliotecas: null,
            idListaSeleccionada: null, idBibliotecaSeleccionada: null, subsitioSeleccionado: false,
            showHiddenLibraries: false, showHiddenLists: false
        };
    }

    private currentUrl: string;
    private currentTitle: string;

    render() {
        let lockListIcon: string = this.state.showHiddenLists ? "fas fa-lock fa-sm" : "fas fa-lock-open fa-sm";
        let lockLibraryIcon: string = this.state.showHiddenLibraries ? "fas fa-lock fa-sm" : "fas fa-lock-open fa-sm";

        return <div id="SPT.Explorer">
            <div className="panelIzquierdo">
                <div className="tituloSeccion">
                    <span>{Constants.getLiteral("explorerListas")}</span>
                    <span className="action" onClick={() => this.clickHideLists()}><i className={lockListIcon}></i></span>
                </div>
                {
                    this.state.resultadosListas &&
                    this.state.resultadosListas.map((lista) =>
                        <a href="#" key={lista.ID} id={lista.ID} onClick={this.clickLista}><i className="fas fa-sticky-note"></i> {lista.Title}</a>
                    )
                }
                <br />
                <div className="tituloSeccion">
                    <span>{Constants.getLiteral("explorerBibliotecas")}</span>
                    <span className="action" onClick={() => this.clickHideLibraries()}><i className={lockLibraryIcon}></i></span>
                </div>
                {
                    this.state.resultadosBibliotecas &&
                    this.state.resultadosBibliotecas.map((biblioteca) =>
                        <a href="#" key={biblioteca.ID} id={biblioteca.ID} onClick={this.clickBiblioteca}><i className="fas fa-book"></i> {biblioteca.Title}</a>
                    )
                }
                <br />
                <div className="tituloSeccion">{Constants.getLiteral("explorerSubsitios")}</div>
                <a href="#" onClick={(e) => this.clickSubsitio(e)}><i className="fas fa-door-open"></i> {Constants.getLiteral("explorerSubsitiosEstructura")}</a>
            </div>
            <div className="panelDerecho">
                {
                    !this.state.subsitioSeleccionado &&
                    <div id="siteTitle">{this.currentTitle}</div>
                }
                <div id="panelDatos">
                    {
                        this.state.idListaSeleccionada &&
                        <div>
                            <ListViewer ID={this.state.idListaSeleccionada} Url={this.currentUrl} Nivel={0} />
                        </div>
                    }
                    {
                        this.state.idBibliotecaSeleccionada &&
                        <div>
                            <LibraryViewer ID={this.state.idBibliotecaSeleccionada} Url={this.currentUrl} Nivel={0} />
                        </div>
                    }
                    {
                        this.state.subsitioSeleccionado &&
                        <div>
                            <WebViewer Url={this.currentUrl} />
                        </div>
                    }
                    {
                        !this.state.idListaSeleccionada && !this.state.idBibliotecaSeleccionada && !this.state.subsitioSeleccionado &&
                        <p>{Constants.getLiteral("explorerHelp")}</p>
                    }
                </div>
            </div>
        </div>;
    }

    public componentDidUpdate(prevProps: any, prevState: any) {
        if (this.state.showHiddenLibraries != prevState.showHiddenLibraries) {
            this.loadLibraries();
        }
        if (this.state.showHiddenLists != prevState.showHiddenLists) {
            this.loadLists();
        }
    }

    public componentDidMount(): void {
        this.currentUrl = Strings.parseQueryString()["u"];
        this.currentTitle = Strings.parseQueryString()["t"];

        this.loadLists();
        this.loadLibraries();
    }

    private loadLists() {
        let qry: string = SPRest.queryLists(this.currentUrl, 0, this.state.showHiddenLists);
        LogAx.trace("Query Listas:" + qry);

        SPRest.restQuery(qry, RestQueryType.ODataJSON).then((result: any) => {
            try {
                this.setState({
                    resultadosListas: result.value.map((list: any) => ({
                        ID: list.Id,
                        Title: list.Title
                    }))
                });
            } catch (e) {
                LogAx.trace("SPT.Explorer ODataAx.restQuery Lists exception: " + e);
            }
        }, (e) => {
            LogAx.trace("SPT.Explorer ODataAx.restQuery Lists error: " + e);
        });
    }

    private loadLibraries() {
        let qry: string = SPRest.queryLists(this.currentUrl, 1, this.state.showHiddenLibraries);
        LogAx.trace("Query Bibliotecas:" + qry);

        SPRest.restQuery(qry, RestQueryType.ODataJSON).then((result: any) => {
            try {
                this.setState({
                    resultadosBibliotecas: result.value.map((library: any) => ({
                        ID: library.Id,
                        Title: library.Title
                    }))
                });
            } catch (e) {
                LogAx.trace("SPT.Explorer ODataAx.restQuery Library exception: " + e);
            }
        }, (e) => {
            LogAx.trace("SPT.Explorer ODataAx.restQuery Library error: " + e);
        });
    }

    private clickLista = (e: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {
        this.setState({
            idListaSeleccionada: e.currentTarget.id,
            idBibliotecaSeleccionada: null,
            subsitioSeleccionado: false
        });
        e.preventDefault();
    };

    private clickBiblioteca = (e: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {
        this.setState({
            idBibliotecaSeleccionada: e.currentTarget.id,
            idListaSeleccionada: null,
            subsitioSeleccionado: false
        });
        e.preventDefault();
    };

    private clickSubsitio = (e: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {
        this.setState({
            subsitioSeleccionado: true,
            idListaSeleccionada: null,
            idBibliotecaSeleccionada: null
        });
        e.preventDefault();
    };

    private clickHideLibraries() {
        this.setState({
            showHiddenLibraries: !this.state.showHiddenLibraries
        });
    }

    private clickHideLists() {
        this.setState({
            showHiddenLists: !this.state.showHiddenLists
        });
    }
}
