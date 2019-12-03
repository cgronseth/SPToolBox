import * as React from "react";
import * as Modal from 'react-modal';
import { Constants } from "../spt.constants";
import { Strings } from "../spt.strings";
import { CSVDelimiters } from "../spt.csv";
import { read } from "fs";

export interface ISPModalState {
    clauseAccepted: boolean;
}

export interface ISPTModalDownloadZipProps {
    Open: boolean;
    DownloadTotalCount: number;
    DownloadOkCount: number;
    DownloadTotalBytes: number;
    DownloadCurrentBytes: number;
    DownloadErrorCount: number;
    CompressPercentage: number;
    onDownload: () => void;
    onClose: () => void;
    onCancel: () => void;
}

export class SPTModalDownloadZip extends React.Component<ISPTModalDownloadZipProps, ISPModalState> {
    constructor(props: ISPTModalDownloadZipProps) {
        super(props);
        this.downloadModal = this.downloadModal.bind(this);
        this.closeModal = this.closeModal.bind(this);
        this.cancelModal = this.cancelModal.bind(this);
    }

    render() {
        let title: string = Constants.getLiteral("explorerModalDescargarTitulo");
        let mensajeDescarga: string = Constants.getLiteral("explorerModalDescargar");
        let mensajeDescargaErrores: string = Constants.getLiteral("explorerModalDescargarErrores");
        let mensajeCompresion: string = Constants.getLiteral("explorerModalComprimir");
        let downloadButton: string = Constants.getLiteral("generalBotonDescargar");
        let cancelButton: string = Constants.getLiteral("generalBotonCancelar");
        let closeButton: string = Constants.getLiteral("generalBotonCerrar");

        let mostrarDescargar: boolean = this.props.CompressPercentage === 100;

        return <Modal
            isOpen={this.props.Open}
            contentLabel={title}
            className="modalDialogWindow"
            ariaHideApp={false}>

            <h3>{title}</h3>
            <div>
                {mensajeDescarga}: {this.props.DownloadOkCount}/{this.props.DownloadTotalCount}
                <span> </span>
                ({Strings.closestByteMetric(this.props.DownloadCurrentBytes)}/{Strings.closestByteMetric(this.props.DownloadTotalBytes)})
            </div>
            {
                this.props.DownloadErrorCount > 0 &&
                <div>{mensajeDescargaErrores}: {this.props.DownloadErrorCount}</div>
            }
            <br />
            <div>{mensajeCompresion}: {this.props.CompressPercentage.toFixed(2)}%</div><br />
            {
                mostrarDescargar &&
                <div>
                    <button onClick={this.downloadModal}>{downloadButton}</button>
                    <button onClick={this.closeModal}>{closeButton}</button>
                </div>
            }
            {
                !mostrarDescargar &&
                <button onClick={this.cancelModal}>{cancelButton}</button>
            }
        </Modal>;
    }

    private downloadModal() {
        this.props.onDownload();
    }

    private closeModal() {
        this.props.onClose();
    }

    private cancelModal() {
        this.props.onCancel();
    }
}

export interface ISPTModalDownloadExcelProps {
    Processing: boolean;
    onDownload: (delimiter: CSVDelimiters) => void;
    onClose: () => void;
}

export interface ISPModalDownloadExcelState {
    delimiter: CSVDelimiters;
}

export class SPTModalDownloadExcel extends React.Component<ISPTModalDownloadExcelProps, ISPModalDownloadExcelState> {
    constructor(props: ISPTModalDownloadExcelProps) {
        super(props);

        let defaultSelectValue: CSVDelimiters;
        switch (Constants.getLCID()) {
            case "es":
                defaultSelectValue = CSVDelimiters.PuntoComma;
                break;
            default:
                defaultSelectValue = CSVDelimiters.Comma;
                break;
        }

        this.state = {
            delimiter: defaultSelectValue
        };
        this.downloadModal = this.downloadModal.bind(this);
        this.closeModal = this.closeModal.bind(this);
        this.handleOptionDelimiterChange = this.handleOptionDelimiterChange.bind(this);
    }

    render() {
        let title: string = Constants.getLiteral("explorerModalExcelTitle");
        let downloadButton: string = Constants.getLiteral("generalBotonDescargar");
        let closeButton: string = Constants.getLiteral("generalBotonCerrar");

        return <Modal
            isOpen={this.props.Processing}
            contentLabel={title}
            className="modalDialogWindow"
            ariaHideApp={false}>

            <h3>{title}</h3>
            <div>{Constants.getLiteral("explorerModalExcelMessage")}</div>
            <br /><br />
            <small>{Constants.getLiteral("explorerModalExcelMessage2")}</small>
            <br /><br />
            <small>
                {Constants.getLiteral("explorerModalExcelMessage3")}:
                <select value={this.state.delimiter} onChange={this.handleOptionDelimiterChange}>
                    <option value={CSVDelimiters.Comma}>{Constants.getLiteral("explorerModalExcelOption1")}</option>
                    <option value={CSVDelimiters.PuntoComma}>{Constants.getLiteral("explorerModalExcelOption2")}</option>
                </select>
            </small>
            <br /><br /><br />
            <div>
                <button onClick={this.downloadModal}>{downloadButton}</button>
                <button onClick={this.closeModal}>{closeButton}</button>
            </div>
        </Modal>;
    }

    private handleOptionDelimiterChange(event: React.ChangeEvent<HTMLSelectElement>) {
        this.setState({
            delimiter: +event.target.value as CSVDelimiters
        });
    }

    private downloadModal() {
        this.props.onDownload(this.state.delimiter);
    }

    private closeModal() {
        this.props.onClose();
    }
}

export interface ISPTModalImportExcelProps {
    Processing: boolean;
    Analizing: boolean;
    Analized: boolean;
    AnalyzedMessage: string[];
    Uploading: boolean;
    Uploaded: boolean;
    ReadingFileProgress: number;
    UploadingRowOK: number;
    UploadingRowError: number;
    UploadingRowTotal: number;
    onAnalyze: (delimiter: CSVDelimiters, file: File) => void;
    onUpload: () => void;
    onClose: () => void;
}

export interface ISPModalImportExcelState {
    delimiter: CSVDelimiters;
    selectedFile: File;
    errorFile: string;
}

export class SPTModalImportExcel extends React.Component<ISPTModalImportExcelProps, ISPModalImportExcelState> {
    constructor(props: ISPTModalImportExcelProps) {
        super(props);

        let defaultSelectValue: CSVDelimiters;
        switch (Constants.getLCID()) {
            case "es":
                defaultSelectValue = CSVDelimiters.PuntoComma;
                break;
            default:
                defaultSelectValue = CSVDelimiters.Comma;
                break;
        }

        this.state = {
            delimiter: defaultSelectValue,
            selectedFile: null,
            errorFile: null
        };
        this.closeModal = this.closeModal.bind(this);
        this.handleOptionDelimiterChange = this.handleOptionDelimiterChange.bind(this);
        this.handleFileSelection = this.handleFileSelection.bind(this);
    }

    render() {
        let title: string = Constants.getLiteral("explorerModalImportExcelTitle");
        let analyzeButton: string = Constants.getLiteral("explorerModalImportExcelAnalyze");
        let uploadButton: string = Constants.getLiteral("generalBotonImportar");
        let closeButton: string;
        if (!this.props.Analized) {
            closeButton = Constants.getLiteral("generalBotonCerrar");
        } else if (this.props.Uploaded) {
            closeButton = Constants.getLiteral("generalBotonAceptar");
        } else {
            closeButton = Constants.getLiteral("generalBotonCancelar");
        }

        let uploadPercentOK: number = Math.round((this.props.UploadingRowOK * 100) / this.props.UploadingRowTotal);
        let uploadPercentError: number = Math.round((this.props.UploadingRowError * 100) / this.props.UploadingRowTotal);
        let uploadPercentLeft: number = 100 - uploadPercentOK - uploadPercentError;

        return <Modal
            isOpen={this.props.Processing}
            contentLabel={title}
            className="modalDialogWindow"
            ariaHideApp={false}>

            <h3>{title}</h3>
            {
                !this.props.Analized &&
                <div>
                    <div>{Constants.getLiteral("explorerModalImportExcelMessage")}</div>
                    <br />
                    <small>{Constants.getLiteral("explorerModalImportExcelMessage2")}</small>
                    <br /><br />
                    <small>
                        {Constants.getLiteral("explorerModalExcelMessage3")}:
                        <select value={this.state.delimiter} onChange={this.handleOptionDelimiterChange}>
                            <option value={CSVDelimiters.Comma}>{Constants.getLiteral("explorerModalExcelOption1")}</option>
                            <option value={CSVDelimiters.PuntoComma}>{Constants.getLiteral("explorerModalExcelOption2")}</option>
                        </select>
                    </small>
                    <br /><br />
                </div>
            }
            {
                !this.props.Analizing && !this.props.Analized &&
                <input type="file" onChange={this.handleFileSelection}></input>
            }
            {
                this.props.Analizing && this.state.errorFile &&
                <div>
                    <br />
                    {Constants.getLiteral("explorerModalImportExcelError")}: {this.state.errorFile}
                    <br />
                </div>
            }
            {
                this.props.Analizing &&
                <div>
                    <div className="progressBar">
                        <div className="percent" style={{ float: 'left', width: this.props.ReadingFileProgress + "%", backgroundColor: "#75aa75" }}>&nbsp;</div>
                        <div className="percent" style={{ float: 'left', width: (100 - this.props.ReadingFileProgress) + "%" }}>&nbsp;</div>
                        <div className="percentText">
                            {Constants.getLiteral("explorerModalImportExcelMessageReading")}: {this.props.ReadingFileProgress}%
                        </div>
                    </div>
                </div>
            }
            {
                this.props.Analized && !this.props.Uploading &&
                <div>
                    <div>{Constants.getLiteral("explorerModalImportExcelMessage3")}</div>
                    <br />
                    <small>{Constants.getLiteral("explorerModalImportExcelMessage4")}</small>
                    <br /><br />
                    {
                        this.props.AnalyzedMessage.length > 0 &&
                        <div className="details">{Constants.getLiteral("explorerModalImportExcelMessageErrorsFound")}<br /><br />
                            {
                                this.props.AnalyzedMessage.map((message, idx) =>
                                    <span key={idx}>{message}<br /></span>
                                )
                            }
                        </div>
                    }
                </div>
            }
            {
                this.props.Uploading && !this.props.Uploaded &&
                <div>
                    <div>{Constants.getLiteral("explorerModalImportExcelMessage5")}</div>
                    <br />
                    <small>{Constants.getLiteral("explorerModalImportExcelMessage6")}</small>
                    <br /><br />
                </div>
            }
            {
                this.props.Uploading &&
                <div className="progressBar">
                    <div className="percent" style={{ float: 'left', width: uploadPercentOK + "%", backgroundColor: "#75aa75" }}>&nbsp;</div>
                    <div className="percent" style={{ float: 'left', width: uploadPercentError + "%", backgroundColor: "#aa7575" }}>&nbsp;</div>
                    <div className="percent" style={{ float: 'left', width: uploadPercentLeft + "%" }}>&nbsp;</div>
                    <div className="percentText">
                        {Constants.getLiteral("explorerModalImportExcelMessageUploading")}: [OK {uploadPercentOK}%] [Error {uploadPercentError}%]
                    </div>
                </div>
            }
            {
                this.props.Uploaded &&
                <div>
                    <div>{Constants.getLiteral("explorerModalImportExcelMessage7")}</div>
                    <br />
                </div>
            }
            <br /><br />
            <div>
                {
                    !this.props.Analizing &&
                    <button onClick={() => this.analyzeModal()}>{analyzeButton}</button>
                }
                {
                    this.props.Analized && !this.props.Uploading &&
                    <button onClick={() => this.props.onUpload()}>{uploadButton}</button>
                }
                <button onClick={this.closeModal}>{closeButton}</button>
            </div>
        </Modal >;
    }

    private handleOptionDelimiterChange(event: React.ChangeEvent<HTMLSelectElement>) {
        this.setState({
            delimiter: +event.target.value as CSVDelimiters
        });
    }

    private handleFileSelection(event: React.ChangeEvent<HTMLInputElement>) {
        this.setState({
            selectedFile: event.target.files[0]
        });
    }

    private analyzeModal() {
        let type = this.state.selectedFile.type.toLowerCase();

        if (type !== 'text/plain' && type !== 'application/vnd.ms-excel') {
            this.setState({
                errorFile: Constants.getLiteral("explorerModalImportExcelError01") + ": [" + type + "]"
            });
        } else {
            this.props.onAnalyze(this.state.delimiter, this.state.selectedFile);
        }
    }

    private closeModal() {
        this.setState({
            errorFile: null,
            selectedFile: null
        });
        this.props.onClose();
    }
}

export interface ISPTModalCopyLibraryProps {
    Open: boolean;
    FilesOK: number;
    FilesError: number;
    FilesTotal: number;
    CopyInFolder: boolean;
    onAcceptClause: () => void;
    onCloseModal: () => void;
}

export class SPTModalCopyLibrary extends React.Component<ISPTModalCopyLibraryProps, ISPModalState> {
    constructor(props: ISPTModalCopyLibraryProps) {
        super(props);
        this.state = {
            clauseAccepted: false
        };

        this.closeModal = this.closeModal.bind(this);
        this.acceptClause = this.acceptClause.bind(this);
    }

    render() {
        let title: string = Constants.getLiteral("explorerModalCopyTitle");
        let acceptButton: string = Constants.getLiteral("generalBotonAceptar");
        let cancelButton: string = Constants.getLiteral("generalBotonCancelar");

        let mensaje: string;
        if (!this.state.clauseAccepted) {
            if (this.props.CopyInFolder) {
                mensaje = Constants.getLiteral("explorerModalCopyMessageFolderConfirm");
            } else {
                mensaje = Constants.getLiteral("explorerModalCopyMessageConfirm");
            }
        } else {
            if ((this.props.FilesOK + this.props.FilesError) < this.props.FilesTotal) {
                mensaje = Constants.getLiteral("explorerModalCopyMessage");
            } else {
                mensaje = Constants.getLiteral("explorerModalCopyMessageFinished");
            }
            mensaje = mensaje.replace("{%1}", this.props.FilesOK.toString());
            mensaje = mensaje.replace("{%2}", this.props.FilesTotal.toString());
            mensaje = mensaje.replace("{%3}", this.props.FilesError.toString());
        }

        return <Modal
            isOpen={this.props.Open}
            contentLabel={title}
            className="modalDialogWindow"
            ariaHideApp={false}
        >
            <h3>{title}</h3>
            <div>{mensaje}</div><br />
            {
                !this.state.clauseAccepted &&
                <div>
                    <button onClick={this.acceptClause}>{acceptButton}</button>
                    <button onClick={this.closeModal}>{cancelButton}</button>
                </div>
            }
            {
                this.state.clauseAccepted &&
                <div>
                    {
                        (this.props.FilesOK + this.props.FilesError) === this.props.FilesTotal &&
                        <button onClick={this.closeModal}>{acceptButton}</button>
                    }
                    {
                        (this.props.FilesOK + this.props.FilesError) < this.props.FilesTotal &&
                        <button onClick={this.closeModal}>{cancelButton}</button>
                    }
                </div>
            }
        </Modal>;
    }

    private acceptClause() {
        this.props.onAcceptClause();
        this.setState({
            clauseAccepted: true,
        });
    }

    private closeModal() {
        this.setState({
            clauseAccepted: false
        });
        this.props.onCloseModal();
    }
}

export interface ISPTModalCopyListProps {
    Open: boolean;
    Analysing: boolean;
    AnalisisMessages: string[];
    RowsOK: number;
    RowsError: number;
    RowsTotal: number;
    onAcceptClause: () => void;
    onCloseModal: () => void;
}

export class SPTModalCopyList extends React.Component<ISPTModalCopyListProps, ISPModalState> {
    constructor(props: ISPTModalCopyListProps) {
        super(props);
        this.state = {
            clauseAccepted: false
        };

        this.closeModal = this.closeModal.bind(this);
        this.acceptClause = this.acceptClause.bind(this);
    }

    render() {
        let title: string = Constants.getLiteral("explorerModalCopyTitle");
        let acceptButton: string = Constants.getLiteral("generalBotonAceptar");
        let cancelButton: string = Constants.getLiteral("generalBotonCancelar");

        let mensaje: string;
        if (this.props.Analysing) {
            mensaje = Constants.getLiteral("explorerModalCopyMessageAnalysing");
        } else {
            if (!this.state.clauseAccepted) {
                mensaje = Constants.getLiteral("explorerModalCopyMessageConfirmList");
            } else {
                if (this.props.RowsOK + this.props.RowsError < this.props.RowsTotal) {
                    mensaje = Constants.getLiteral("explorerModalCopyMessageList");
                } else {
                    mensaje = Constants.getLiteral("explorerModalCopyMessageFinished");
                }
                mensaje = mensaje.replace("{%1}", this.props.RowsOK.toString());
                mensaje = mensaje.replace("{%2}", this.props.RowsTotal.toString());
                mensaje = mensaje.replace("{%3}", this.props.RowsError.toString());
            }
        }

        return <Modal
            isOpen={this.props.Open}
            contentLabel={title}
            className="modalDialogWindow"
            ariaHideApp={false}
        >
            <h3>{title}</h3>
            <div>{mensaje}</div>
            {
                this.props.AnalisisMessages.length > 0 &&
                <div className="details">{Constants.getLiteral("analisysError")}<br /><br />
                    {
                        this.props.AnalisisMessages.map((message, idx) =>
                            <span key={idx}>{message}<br /></span>
                        )
                    }
                </div>
            }
            {
                this.props.Analysing &&
                <div>
                    <br />
                    <button onClick={this.closeModal}>{cancelButton}</button>
                </div>
            }
            {
                !this.props.Analysing && !this.state.clauseAccepted &&
                <div>
                    <br />
                    <button onClick={this.acceptClause}>{acceptButton}</button>
                    <button onClick={this.closeModal}>{cancelButton}</button>
                </div>
            }
            {
                this.state.clauseAccepted &&
                <div>
                    <br />
                    {
                        (this.props.RowsOK + this.props.RowsError) === this.props.RowsTotal &&
                        <button onClick={this.closeModal}>{acceptButton}</button>
                    }
                    {
                        (this.props.RowsOK + this.props.RowsError) < this.props.RowsTotal &&
                        <button onClick={this.closeModal}>{cancelButton}</button>
                    }
                </div>
            }
        </Modal>;
    }

    private acceptClause() {
        this.props.onAcceptClause();
        this.setState({
            clauseAccepted: true,
        });
    }

    private closeModal() {
        this.setState({
            clauseAccepted: false
        });
        this.props.onCloseModal();
    }
}

export interface ISPTModalDeleteLibraryProps {
    Open: boolean;
    FilesOK: number;
    FilesError: number;
    FilesTotal: number;
    onAcceptClause: () => void;
    onCloseModal: () => void;
}

export class SPTModalDeleteLibrary extends React.Component<ISPTModalDeleteLibraryProps, ISPModalState> {
    constructor(props: ISPTModalDeleteLibraryProps) {
        super(props);
        this.state = {
            clauseAccepted: false
        };

        this.closeModal = this.closeModal.bind(this);
        this.acceptClause = this.acceptClause.bind(this);
    }

    render() {
        let title: string = Constants.getLiteral("explorerModalDeleteTitle");
        let acceptButton: string = Constants.getLiteral("generalBotonAceptar");
        let cancelButton: string = Constants.getLiteral("generalBotonCancelar");

        let mensaje: string;
        if (!this.state.clauseAccepted) {
            mensaje = Constants.getLiteral("explorerModalDeleteMessageConfirm");
        } else {
            if ((this.props.FilesOK + this.props.FilesError) < this.props.FilesTotal) {
                mensaje = Constants.getLiteral("explorerModalDeleteMessage");
            } else {
                mensaje = Constants.getLiteral("explorerModalDeleteMessageFinished");
            }
            mensaje = mensaje.replace("{%1}", this.props.FilesOK.toString());
            mensaje = mensaje.replace("{%2}", this.props.FilesTotal.toString());
            mensaje = mensaje.replace("{%3}", this.props.FilesError.toString());
        }

        return <Modal
            isOpen={this.props.Open}
            contentLabel={title}
            className="modalDialogWindow"
            ariaHideApp={false}
        >
            <h3>{title}</h3>
            <div>{mensaje}</div><br />
            {
                !this.state.clauseAccepted &&
                <div>
                    <button onClick={this.acceptClause}>{acceptButton}</button>
                    <button onClick={this.closeModal}>{cancelButton}</button>
                </div>
            }
            {
                this.state.clauseAccepted &&
                <div>
                    {
                        (this.props.FilesOK + this.props.FilesError) === this.props.FilesTotal &&
                        <button onClick={this.closeModal}>{acceptButton}</button>
                    }
                    {
                        (this.props.FilesOK + this.props.FilesError) < this.props.FilesTotal &&
                        <button onClick={this.closeModal}>{cancelButton}</button>
                    }
                </div>
            }
        </Modal>;
    }

    private acceptClause() {
        this.props.onAcceptClause();
        this.setState({
            clauseAccepted: true,
        });
    }

    private closeModal() {
        this.setState({
            clauseAccepted: false
        });
        this.props.onCloseModal();
    }
}

export interface ISPTModalDeleteListProps {
    Open: boolean;
    RowsOK: number;
    RowsError: number;
    RowsTotal: number;
    onAcceptClause: () => void;
    onCloseModal: () => void;
}

export class SPTModalDeleteList extends React.Component<ISPTModalDeleteListProps, ISPModalState> {
    constructor(props: ISPTModalDeleteListProps) {
        super(props);
        this.state = {
            clauseAccepted: false
        };

        this.closeModal = this.closeModal.bind(this);
        this.acceptClause = this.acceptClause.bind(this);
    }

    render() {
        let title: string = Constants.getLiteral("explorerModalDeleteTitle");
        let acceptButton: string = Constants.getLiteral("generalBotonAceptar");
        let cancelButton: string = Constants.getLiteral("generalBotonCancelar");

        let mensaje: string;
        if (!this.state.clauseAccepted) {
            mensaje = Constants.getLiteral("explorerModalDeleteMessageConfirm");
        } else {
            if ((this.props.RowsOK + this.props.RowsError) < this.props.RowsTotal) {
                mensaje = Constants.getLiteral("explorerModalDeleteMessageList");
            } else {
                mensaje = Constants.getLiteral("explorerModalDeleteMessageFinished");
            }
            mensaje = mensaje.replace("{%1}", this.props.RowsOK.toString());
            mensaje = mensaje.replace("{%2}", this.props.RowsTotal.toString());
            mensaje = mensaje.replace("{%3}", this.props.RowsError.toString());
        }

        return <Modal
            isOpen={this.props.Open}
            contentLabel={title}
            className="modalDialogWindow"
            ariaHideApp={false}
        >
            <h3>{title}</h3>
            <div>{mensaje}</div><br />
            {
                !this.state.clauseAccepted &&
                <div>
                    <button onClick={this.acceptClause}>{acceptButton}</button>
                    <button onClick={this.closeModal}>{cancelButton}</button>
                </div>
            }
            {
                this.state.clauseAccepted &&
                <div>
                    {
                        (this.props.RowsOK + this.props.RowsError) === this.props.RowsTotal &&
                        <button onClick={this.closeModal}>{acceptButton}</button>
                    }
                    {
                        (this.props.RowsOK + this.props.RowsError) < this.props.RowsTotal &&
                        <button onClick={this.closeModal}>{cancelButton}</button>
                    }
                </div>
            }
        </Modal>;
    }

    private acceptClause() {
        this.props.onAcceptClause();
        this.setState({
            clauseAccepted: true,
        });
    }

    private closeModal() {
        this.setState({
            clauseAccepted: false
        });
        this.props.onCloseModal();
    }
}