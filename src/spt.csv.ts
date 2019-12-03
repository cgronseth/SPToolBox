import { SPListItem, SPView, SPField, SPData } from "./sharepoint/spt.sharepoint.entities";
import { LogAx } from "./spt.logax";
import { IListItemLight } from "./spt.storage";
import { lookup } from "dns";

export enum CSVDelimiters {
    Comma = 0,
    PuntoComma
}

export class CSV {
    /**
     * Create CSV from list item data, ordered as in the specified view
     * @param items 
     * @param view 
     * @param delimitador 
     */
    public static generateCSV(items: SPListItem[], view: SPView, delimiter: CSVDelimiters): string {
        let csvData: string[] = [];
        try {
            let csvDelimiter: string = delimiter === CSVDelimiters.Comma ? "," : ";";
            let row: string[] = [];

            let hasID: boolean = CSV.hasField(view, "ID");
            let hasAuthor: boolean = CSV.hasField(view, "Author");
            let hasEditor: boolean = CSV.hasField(view, "Editor");
            let hasCreated: boolean = CSV.hasField(view, "Created");
            let hasModified: boolean = CSV.hasField(view, "Modified");

            // Add header. Include common fields if not included in view
            hasID || row.push("ID");
            hasAuthor || row.push("Author");
            hasEditor || row.push("Editor");
            hasCreated || row.push("Created");
            hasModified || row.push("Modified");

            for (let field of view.ViewFields) {
                row.push(CSV.getSafeCSVString(field.InternalName));
            }
            csvData.push(row.join(csvDelimiter));

            // Add rows
            for (let item of items) {
                row = [];
                hasID || row.push(item.ID.toString());

                for (let field of view.ViewFields) {
                    let itemData: SPData = item.getItem(field.InternalName);
                    let itemValue: string = itemData.StringValue;
                    switch (field.Type) {
                        case 7: //Lookup
                            if (itemData.LookupId) {
                                itemValue = itemData.LookupId + ";#" + itemData.StringValue;
                            }
                            break;
                        default:
                            break;
                    }
                    row.push(CSV.getSafeCSVString(itemValue));
                }

                hasAuthor || row.push(item.Author.Email);
                hasEditor || row.push(item.Editor.Email);
                hasCreated || row.push(item.Created.toString());
                hasModified || row.push(item.Modified.toString());

                csvData.push(row.join(csvDelimiter));
            }
        }
        catch (e) {
            LogAx.trace("SPT.CSV.generateCSV Error: " + e);
            csvData = [];
        }
        return csvData.join("\r\n");
    }

    public static parseCSV(data: string, listFields: SPField[], delimiter: CSVDelimiters): IListItemLight[] {
        let output: IListItemLight[] = [];
        try {
            // Parse csv data into array of rows and columns (code found on stackoverflow: https://stackoverflow.com/questions/1293147/javascript-code-to-parse-csv-data)
            let csvDelimiter: string = delimiter === CSVDelimiters.Comma ? "," : ";";
            let objPattern = new RegExp(("(\\" + csvDelimiter + "|\\r?\\n|\\r|^)(?:\"((?:\\\\.|\"\"|[^\\\\\"])*)\"|([^\\" + csvDelimiter + "\"\\r\\n]*))"), "gi");
            let arrMatches = null;
            let arrData: any = [[]];
            while (arrMatches = objPattern.exec(data)) {
                if (arrMatches[1].length && arrMatches[1] !== csvDelimiter) {
                    arrData.push([]);
                }
                arrData[arrData.length - 1].push(arrMatches[2] ? arrMatches[2].replace(new RegExp("[\\\\\"](.)", "g"), '$1') : arrMatches[3]);
            }

            // Determine full fields in order of appearance in first row (column headers required in csv)
            let fields: SPField[] = [];
            for (let colIdx = 0; colIdx < arrData[0].length; colIdx++) {
                let findFields = listFields.filter(f => f.InternalName === arrData[0][colIdx]);
                if (findFields.length !== 0) {
                    fields.push(findFields[0]);
                } else {
                    LogAx.trace("SPT.CSV.parseCSV Field not found: " + arrData[0][colIdx]);
                    //Create faux field with available information
                    fields.push({
                        ID: "",
                        InternalName: arrData[0][colIdx],
                        Title: arrData[0][colIdx],
                        StaticName: arrData[0][colIdx],
                        Description: "",
                        Type: 0,
                        Required: false,
                        Hidden: false
                    });
                }
            }

            // Convert raw array to structured item objects
            for (let rowIdx = 1; rowIdx < arrData.length; rowIdx++) {
                let dataRow = arrData[rowIdx];
                let item: IListItemLight = {
                    ID: null,
                    Author: null,
                    Created: null,
                    Editor: null,
                    Modified: null,
                    ItemData: []
                };

                for (let colIdx = 0; colIdx < dataRow.length; colIdx++) {
                    let field = fields[colIdx];
                    let columnData: string = dataRow[colIdx];
                    if (field.InternalName === "Id" || field.InternalName === "ID") {
                        item.ID = +columnData;
                    } else if (field.InternalName === "Author") {
                        item.Author = columnData;
                    } else if (field.InternalName === "Created") {
                        item.Created = new Date(columnData);
                    } else if (field.InternalName === "Editor") {
                        item.Editor = columnData;
                    } else if (field.InternalName === "Modified") {
                        item.Modified = new Date(columnData);
                    } else {
                        let lookupId: number = null;
                        if (field.Type === 7 && columnData.indexOf(";#") !== -1) {
                            let lookupData: string[] = columnData.split(";#");
                            lookupId = +lookupData[0];
                            columnData = lookupData[1];
                        }
                        item.ItemData.push({
                            InternalName: field.InternalName,
                            Type: field.Type,
                            StringValue: columnData,
                            LookupId: lookupId
                        });
                    }
                }
                output.push(item);
            }
        }
        catch (e) {
            LogAx.trace("SPT.CSV.parseCSV Error: " + e);
            output = [];
        }
        return output;
    }

    private static regexSafeString = new RegExp(/("|,|;|\n)/);
    private static getSafeCSVString(input: string): string {
        if (!input) {
            return "";
        }
        let output = input.replace(/"/g, '""');
        return (this.regexSafeString.test(output)) ? '"' + output + '"' : output;
    }

    private static hasField(view: SPView, internalName: string): boolean {
        return view.ViewFields.find(v => v.InternalName === internalName) !== undefined;
    }

}

