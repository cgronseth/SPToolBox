import * as React from "react";
import { Constants } from "../../spt.constants"

export enum SPTButtonType {
    MenuItem
}

export interface ISPTButtonProps {
    show: boolean;
    enabled: boolean;
    type: SPTButtonType;
    icon: string;
    textId: string;
    onClick: () => void;
}

interface ISPTButtonState {
}

export class SPTButton extends React.Component<ISPTButtonProps, ISPTButtonState> {
    constructor(props: ISPTButtonProps) {
        super(props);
        this.state = {
        };
    }

    render() {
        let renderedButton = null;

        switch (this.props.type) {
            case SPTButtonType.MenuItem:
                let renderedIcon = (!this.props.icon) ? null : <i className={this.props.icon}></i>
                let renderedText = (!this.props.textId) ? null : <span> {Constants.getLiteral(this.props.textId)}</span>

                if (this.props.show && this.props.enabled) {
                    renderedButton =
                        <div className="menuItem" onClick={() => this.props.onClick()}>
                            {renderedIcon}
                            {renderedText}
                        </div>;
                } else if (this.props.show && !this.props.enabled) {
                    renderedButton =
                        <div className="menuItemDisabled">
                            {renderedIcon}
                            {renderedText}
                        </div>;
                }
                break;
            default:
                renderedButton = <div>Dummy Button</div>;
                break;
        }
        return renderedButton;
    }

    public componentDidMount(): void {
    }

    public componentWillUnmount(): void {
    }
}
