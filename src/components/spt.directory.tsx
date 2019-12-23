import * as React from "react";
import { Constants } from "../spt.constants"
import { LogAx } from "../spt.logax";
import { SPRest, RestQueryType } from "../sharepoint/spt.sharepoint.rest";
import { SPList, SPWeb } from "../sharepoint/spt.sharepoint.entities";
import { ListViewer } from "./spt.explorer.listviewer";
import { LibraryViewer } from "./spt.explorer.libraryviewer";
import { WebViewer } from "./spt.explorer.webviewer";
import { Strings } from "../spt.strings";

export interface IDirectoryState {
    showPermissionsSite: boolean;

}

export class Directory extends React.Component<{}, IDirectoryState> {
    constructor(props: Readonly<{}>) {
        super(props);
        this.state = {
            showPermissionsSite: false
        };
    }

    private currentUrl: string;
    private currentTitle: string;

    render() {
        return <div id="SPT.Directory">
            <div className="panelIzquierdo">
                <div className="tituloSeccion">
                    <span>{Constants.getLiteral("explorerListas")}</span>
                </div>
                <a href="#" onClick={this.clickPermisosSitio}><i className="fas fa-sticky-note"></i> {Constants.getLiteral("directoryPermissionsSite")}</a>
            </div>
            <div className="panelDerecho">
                {
                    this.state.showPermissionsSite &&
                    <div>Oh my permisssions!</div>
                }
                <div id="panelDatos">

                </div>
            </div>
        </div>;
    }

    public componentDidUpdate(prevProps: any, prevState: any) {

    }

    public componentDidMount(): void {
        this.currentUrl = Strings.parseQueryString()["u"];
        this.currentTitle = Strings.parseQueryString()["t"];

    }

    private clickPermisosSitio = (e: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {
        this.setState({
            showPermissionsSite: true
        });
        e.preventDefault();
    };

}
