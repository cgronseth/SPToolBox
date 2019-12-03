import { number } from "prop-types";

export class SPSecurableObject {
    HasUniqueRoleAssignments?: boolean;
    RoleAssignments?: string[];
};

export class SPData {
    constructor(internalName: string, value: string, type: number) {
        this.InternalName = internalName;
        this.StringValue = value;
        this.Type = type;
    }
    InternalName: string;
    LookupId?: number;
    StringValue: string;
    Type: number;
}

export class SPItem extends SPSecurableObject {
    ID: number;
    Items: SPData[] = [];
    public getItem(internalName: string): SPData {
        for (let i = 0; i < this.Items.length; i++) {
            let element: SPData = this.Items[i];
            if (element.InternalName === internalName) {
                return element;
            }
        }
        return null;
    }
}

export class SPField {
    ID: string; //GUID
    InternalName: string;
    StaticName: string;
    Title: string;
    Description: string;
    Type: number; //https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ms428806%28v%3doffice.15%29
    Required: boolean;
    Hidden: boolean;
    //Lookup fields (Type=7)
    LookupField?: string;
}

export class SPAttachment {
    LeafName: string;
    FileName: string;
    Data: ArrayBuffer;
}

export class SPFolder {
    UniqueId: string;
    Name: string;
    ParentFolder?: SPFolder;
    ServerRelativeUrl?: string;
    ItemCount?: number;
    // Not SharePoint API members
    Level?: number;                 //Helps calculate folder level
    Expand: boolean;                //Control view of items in folder
}

export class SPListItem extends SPItem {
    Name?: string;
    SPFileSystemObjectType: number;
    Attachments?: SPAttachment[];
    Folder?: SPFolder;
    Created: Date;
    Author: SPUser;
    Modified: Date;
    Editor: SPUser;
    Length: number = 0;
    // Not SharePoint API members
    Hidden: Boolean;                //Hide item
    Checked: number = 0;            //Checked state: 0:Not checked, 1:Checked, -1:Partial checked (not all related items checked)
}

export class SPList extends SPSecurableObject {
    ID: string;
    Title?: string;
    InternalName?: string;
    ListItemEntityTypeFullName?: string;
    ItemCount?: number;
    Items?: SPListItem[];
    Fields?: SPField[];
    Hidden?: boolean;
    RootFolder?: SPFolder;
}

export class SPView {
    ID: string;
    Title: string;
    DefaultView: boolean;
    PersonalView: boolean;
    RowLimit: number;
    ServerRelativeUrl: string;
    ViewFields: SPField[];
}

export class SPGroup {
    Name: string;
}

export class SPWeb extends SPSecurableObject {
    ID: string; //GUID
    Url: string;
    ServerRelativeUrl: string;
    Name: string;
    Description?: string;
    ParentWeb?: SPWeb;
    Groups?: SPGroup[];
    Webs?: SPWeb[];
    Lists?: SPList[];
    Created?: Date;
}

export class SPSite {
    ID: string; //GUID
    Url: string;
}

export class SPUser {
    constructor(displayName: string, email: string) {
        this.DisplayName = displayName;
        this.Email = email;
    }
    DisplayName: string;
    Email: string;
}

export enum PermissionKind {
    emptyMask = 0,
    viewListItems = 1,
    addListItems = 2,
    editListItems = 3,
    deleteListItems = 4,
    approveItems = 5,
    openItems = 6,
    viewVersions = 7,
    deleteVersions = 8,
    cancelCheckout = 9,
    managePersonalViews = 10,
    manageLists = 12
    // Add or remove from SP.PermissionKind in sp.debug.js
}