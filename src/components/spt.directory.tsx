import * as React from "react";
import { Constants } from "../spt.constants"
import { Strings } from "../spt.strings";
import { SitePermissions } from "./spt.directory.siteperm";

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
                    <span>{Constants.getLiteral("directory")}</span>
                </div>
                <a href="#" onClick={this.clickPermisosSitio}><i className="fas fa-sticky-note"></i> {Constants.getLiteral("directoryPermissionsSite")}</a>
            </div>
            <div className="panelDerecho">
                {
                    this.state.showPermissionsSite &&
                    <SitePermissions Url={this.currentUrl} Title={this.currentTitle} />
                }
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
