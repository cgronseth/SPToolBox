import * as React from "react";
import { Constants } from "../spt.constants";
import { WebDetails } from "./spt.explorer.webdetails";

export interface IExplorerWebProps {
    Url: string;
}

export class WebViewer extends React.Component<IExplorerWebProps, {}> {
    constructor(props: IExplorerWebProps) {
        super(props);
        this.state = {};
    }

    render() {
        return <div>
            {
                this.props.Url &&
                <div>
                    <WebDetails key={"0"} Url={this.props.Url} Nivel={0} />
                </div>
            }
        </div>;
    }

}