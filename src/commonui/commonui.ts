import { SPListItem, SPView } from "../sharepoint/spt.sharepoint.entities";

export class CommonUI {

    private static imageCheckBoxEmpty: string = "./icons/checkMarco13.png";
    private static imageCheckBoxDot: string = "./icons/checkMarcoDot13.png";
    private static imageCheckBoxChecked: string = "./icons/check13.png";
    private static imageCheckBoxCheckedClicked: string = "./icons/check13g.png";

    public static checkHeaderImage(allChecked: boolean): string {
        return allChecked ? this.imageCheckBoxChecked : this.imageCheckBoxEmpty
    }

    public static checkRowImage(rowListItem: SPListItem, lastClickedItemID: number): string {
        let imgSrcCheckBox: string;
        if (rowListItem.Checked === 1) {
            imgSrcCheckBox = (rowListItem.ID === lastClickedItemID) ? this.imageCheckBoxCheckedClicked : this.imageCheckBoxChecked;
        } else if (rowListItem.Checked === 0) {
            imgSrcCheckBox = this.imageCheckBoxEmpty;
        } else {
            imgSrcCheckBox = this.imageCheckBoxDot;
        }
        return imgSrcCheckBox;
    }

    /**
    * Get width of scrollbar in Browser. Hack for Virtualized when needed.
    */
    public static getScrollbarWidth(document: HTMLDocument): number {
        let sbwidth: number;
        const scrollDiv: HTMLDivElement = document.createElement('div');
        scrollDiv.style.position = 'absolute';
        scrollDiv.style.top = '-9999px';
        scrollDiv.style.width = '50px';
        scrollDiv.style.height = '50px';
        scrollDiv.style.overflow = 'scroll';

        document.body.appendChild(scrollDiv);
        sbwidth = scrollDiv.getBoundingClientRect().width - scrollDiv.clientWidth;
        document.body.removeChild(scrollDiv);

        return sbwidth;
    }

    public static readonly initialSeparation: number = 12;    //Padding to left of folder structure line
    public static readonly levelSeparation: number = 14;      //Separation of each new folder level line

    public static getColumnTypeWidth(index: number, fixedColumns: number, view: SPView, folderDepth?: number) {
        // Test fixed columns by position in grid
        switch (index) {
            case 0: // CheckBox
                return 20;
            case 1: // Folder structure, only in library view
                if (fixedColumns > 1) {
                    return Math.max(60, CommonUI.initialSeparation + ((folderDepth || 0) * CommonUI.levelSeparation));
                }
        }
        // Test by internal field name
        switch (view.ViewFields[index - fixedColumns].InternalName) {
            case 'ComplianceAssetId':
                return 310;
            case 'ContentTypeId':
                return 300;
            case 'DocIcon':
                return 70;
            case 'EncodedAbsUrl':
                return 350;
            case 'FileLeafRef':
                return 250;
            case 'FileRef':
                return 350;
        }
        // Test by field type
        switch (view.ViewFields[index - fixedColumns].Type) {
            case 20:    //User
                return 180;
            case 4:     //DateTime
                return 150;
            case 8:     //Boolean
                return 80;
            case 5:     //Counter (ID)
                return 50;
            case 1:     //Integer
                return 80;
        }
        return 200; //default width
    }
}