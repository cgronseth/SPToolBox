import { SPField, SPListItem, SPFolder, SPData, PermissionKind } from "./spt.sharepoint.entities";
import { LogAx } from "../spt.logax";
import { Constants } from "../spt.constants";

export interface FileData {
    FileName: string;
    FileData: ArrayBuffer;
    FileLength: number;
}

export class SP {
    // Order SPListItem array by their fields. Sorting fields are passed in "fields" array.
    // If a field ends with a dot (.) it orders Descending. Otherwise Ascending
    public static orderLibraryDataByFields(data: SPListItem[], fields: string[]): SPListItem[] {
        if (fields && fields.length && data && data.length) {
            data.sort((a, b) => {
                return SP.orderLibraryDataByFieldsRecursive(a, b, fields, 0);
            });
        }
        return data;
    }
    private static orderLibraryDataByFieldsRecursive(a: SPListItem, b: SPListItem, fields: string[], idx: number): number {
        if (fields.length == idx || idx > 5)
            return 0;

        let field: string = fields[idx];
        let direction: number = 1; //1=ASC, -1=DESC
        if (field.endsWith('.')) {
            field = field.slice(0, -1);
            direction = -1;
        }

        //Obtain sortable data from field. Data must include value and type.
        let objA: SPData;
        let objB: SPData;
        switch (field) {
            case "ID":
                objA = new SPData(field, a.ID.toString(), 5);
                objB = new SPData(field, b.ID.toString(), 5);
                break;
            case "Author":
                objA = new SPData(field, a.Author.DisplayName, 255);
                objB = new SPData(field, b.Author.DisplayName, 255);
                break;
            case "Editor":
                objA = new SPData(field, a.Editor.DisplayName, 255);
                objB = new SPData(field, b.Editor.DisplayName, 255);
                break;
            case "Created":
                objA = new SPData(field, a.Created.toISOString(), 4);
                objB = new SPData(field, b.Created.toISOString(), 4);
                break;
            case "Modified":
                objA = new SPData(field, a.Modified.toISOString(), 4);
                objB = new SPData(field, b.Modified.toISOString(), 4);
                break;
            case "Folder":
                // Los ítems normales no tiene Folder, pero sí ParentFolder. 
                // Agregar carácter | a los ítems normales para que ordene siempre debajo de su carpeta padre.
                if (a.Folder.Name) {
                    objA = new SPData(field, a.Folder.ServerRelativeUrl, 255);
                } else {
                    objA = new SPData(field, a.Folder.ParentFolder.ServerRelativeUrl + "|", 255);
                }
                if (b.Folder.Name) {
                    objB = new SPData(field, b.Folder.ServerRelativeUrl, 255);
                } else {
                    objB = new SPData(field, b.Folder.ParentFolder.ServerRelativeUrl + "|", 255);
                }
                break;
            default:
                objA = a.getItem(field);
                objB = b.getItem(field);
                break;
        }

        // Apply sorting algorithm depending on data type. Direction reverses result.
        switch (objA.Type) {
            case 1: //Integer
            case 5: //Counter
            case 9: //NumberField
                let na: number = +objA.StringValue;
                let nb: number = +objB.StringValue;
                if (na < nb)
                    return -1 * direction;
                if (na > nb)
                    return 1 * direction;
                break;
            case 4: //Date
                let da: Date = new Date(objA.StringValue);
                let db: Date = new Date(objB.StringValue);
                if (da < db)
                    return -1 * direction;
                if (da > db)
                    return 1 * direction;
                break;
            case 8: //Boolean
                let ba: boolean = objA.StringValue.toLowerCase() === "true";
                let bb: boolean = objB.StringValue.toLowerCase() === "true";
                if (!ba && nb)
                    return -1 * direction;
                if (ba && !bb)
                    return 1 * direction;
                break;
            case 14://GUID
                if (objA.StringValue.toLowerCase() < objB.StringValue.toLowerCase())
                    return -1 * direction;
                if (objA.StringValue.toLowerCase() > objB.StringValue.toLowerCase())
                    return 1 * direction;
                break;
            default://Strings
                if (objA.StringValue < objB.StringValue)
                    return -1 * direction;
                if (objA.StringValue > objB.StringValue)
                    return 1 * direction;
                break;
        }
        // If execution reaches here they are equal, sort by next field
        return SP.orderLibraryDataByFieldsRecursive(a, b, fields, idx + 1);
    }

    public static findFolderByPath(cacheFolders: Map<string, SPFolder>, fileRef: string): SPFolder {
        if (fileRef) {
            let fileRefLower: string = fileRef.toLowerCase();
            for (let value of cacheFolders.values()) {
                if (value.ServerRelativeUrl && value.ServerRelativeUrl.toLowerCase() === fileRefLower) {
                    return value;
                }
            }
        }
        return null;
    }

    /**
     * From folder obtain path structure.
     * <library name>/<sub folder>/<current path>
     * @param folder
     * @param currentPath 
     */
    public static getFolderPath(folder: SPFolder, currentPath: string): string {
        return folder ? SP.getFolderPath(folder.ParentFolder, folder.Name + "/" + currentPath) : currentPath;
    }

    /**
     * From folder obtain path structure. Omit parent library in path.
     * <sub folder>/<current path>
     * @param folder 
     * @param currentPath 
     */
    public static getFolderPathWithoutParent(folder: SPFolder, currentPath: string): string {
        let path: string = this.getFolderPath(folder, currentPath);
        if (path.indexOf('/') !== -1) {
            path = path.substring(path.indexOf('/') + 1);   //Remove first path and leading slash
            if (path.indexOf('/') === -1) {
                path = "/" + path;  //Add forward slash if path was root without folders and was removed in previous step
            }
        }
        return path;
    }

    /**
     * Get maximum folder depth level
     */
    public static getFolderDepthLevel(items: SPListItem[]): number {
        let levels: number[] = items
            .filter(i => i.Folder !== null && i.Folder.Level !== null)
            .map(i => i.Folder.Level);
        return levels.length ? Math.max(...levels) : 0;
    }

    private static literalYes = Constants.getLiteral("generalSi");
    private static literalNo = Constants.getLiteral("generalNo");

    public static parseItemJsonResult(i: any, field: SPField): SPData {
        let spd: SPData = new SPData(field.InternalName, "", field.Type);
        try {
            let odataInternalName: string = SP.safeOdataField(field.InternalName);

            if (i[odataInternalName] !== null && i[odataInternalName] !== undefined) {
                switch (field.Type) {
                    case 7:     //Lookup
                        spd.StringValue = i[odataInternalName][SP.safeOdataField(field.LookupField)];
                        spd.LookupId = i[odataInternalName]["ID"];
                        break;
                    case 20:    //User
                        spd.StringValue = i[odataInternalName]["EMail"];
                        break;
                    case 8:     //Yes/No
                        spd.StringValue = i[odataInternalName] ? SP.literalYes : SP.literalNo;
                        break;
                    default:
                        spd.StringValue = i[odataInternalName] + "";
                        break;
                }
            }
        } catch {
            LogAx.trace("Error parsing column '" + field.InternalName + "' on item id:" + i.ID);
        }
        return spd;
    }

    /**
     * Apply OData renaming to fields with special characters. Just necessary in Rest queries and results
     * @param field 
     */
    public static safeOdataField(field: string) {
        return (!!field && field.startsWith("_x")) ? "OData_" + field : field;
    }

    /**
     * Adapt display values to SharePoint accepted input values 
     * @param stringValue 
     * @param type 
     * @description Sometimes SharePoint doesn't accept its own values, so they must be converted, like boolean from "Yes"->"true"/"No"->"false",
     * or Lookup values use IDs
     */
    public static assignPostValue(data: SPData): string | boolean | number {
        switch (data.Type) {
            case 8:     //Boolean
                return (data.StringValue.toLowerCase() === "no") ? false : true;
            case 7:     //Lookup
                return data.LookupId;
            default:
                return data.StringValue;
        }
    }

    /**
     * Determine Permission based on EffectiveBasePermissions result from REST
     * Stripped and adapted from actual SP.JS code
     * @param high Result "High" from EffectiveBasePermissions query
     * @param low Result "Low" from EffectiveBasePermissions query
     * @param permission Permission mask to check 
     */
    public static checkEffectivePermission(high: number, low: number, permission: PermissionKind): boolean {
        if (!permission) {
            return true;
        }
        var $v_0 = permission - 1;
        var $v_1 = 1;
        if ($v_0 >= 0 && $v_0 < 32) {
            $v_1 = $v_1 << $v_0;
            return 0 !== (low & $v_1);
        }
        if ($v_0 >= 32 && $v_0 < 64) {
            $v_1 = $v_1 << ($v_0 - 32);
            return 0 !== (high & $v_1);
        }
        return false;
    }
}
