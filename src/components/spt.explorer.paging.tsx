import * as React from "react";
import { Constants } from "../spt.constants"

export interface IListPagingProps {
    AvailableHeight: number;
    TotalItems: number;
    LoadedItems: number;
    onLoadFullList: () => void;
}

interface IListPagingState {
    fullListLoading: boolean,
    fullListLoaded: boolean
}

export class ListPaging extends React.Component<IListPagingProps, IListPagingState> {
    constructor(props: IListPagingProps) {
        super(props);
        this.state = {
            fullListLoading: false,
            fullListLoaded: false
        };
        this.loadFullList = this.loadFullList.bind(this);
    }

    render() {
        return <div id="ListPaging">
            {
                this.props.TotalItems >= 0 &&
                <div className="infoItem">
                    <span>{Constants.getLiteral("libraryViewerTotalItems")}: {this.props.TotalItems}</span>
                </div>
            }
            {
                this.props.LoadedItems > 0 &&
                <div className="infoItem">
                    <span>{Constants.getLiteral("libraryViewerItemsCargados")}: {this.props.LoadedItems}</span>
                </div>
            }
            {
                this.isLoading() &&
                <div className="infoItem">
                    <img src="icons/ajax-loader.gif" />&nbsp;
                    <span>{Constants.getLiteral("libraryViewerCargandoItems")}</span>
                </div>
            }
            {
                this.props.TotalItems > 0 && !this.state.fullListLoading && !this.state.fullListLoaded &&
                <div className="buttonItem" onClick={this.loadFullList}>
                    <i className="fas fa-weight-hanging fa-lg"></i>&nbsp;
                    <span>{Constants.getLiteral("libraryViewerCargarListaCompleta")}</span>
                </div>
            }

        </div>;
    }

    public componentDidUpdate(prevProps: IListPagingProps, prevState: IListPagingState) {
        if (prevProps.LoadedItems != this.props.LoadedItems) {
            if (this.props.LoadedItems == this.props.TotalItems) {
                this.setState({
                    fullListLoaded: true,
                    fullListLoading: false
                });
            }
        }
    }

    public componentDidMount(): void {
    }

    public componentWillUnmount(): void {
    }

    private loadFullList() {
        this.setState({
            fullListLoading: true
        });
        this.props.onLoadFullList();
    }

    private isLoading(): boolean {
        if (this.props.LoadedItems === 0 && this.props.LoadedItems < this.props.TotalItems) {
            return true;
        }
        if (this.props.LoadedItems > 0 && this.props.LoadedItems < this.props.TotalItems) {
            if (this.state.fullListLoading && !this.state.fullListLoaded) {
                return true;
            }
        }
        return false;
    }

}
