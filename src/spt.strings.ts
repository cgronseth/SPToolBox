export class Strings {

    public static readonly UTFBOMStartCode: string = "\ufeff";

    /**
     * Obtains URL closest to the absolute URL
     * @param url 
     */
    public static getWebUrlFromAbsolute(url: string): string {
        let u: string = url.toLowerCase().trim();
        if (u.indexOf('?') !== -1) {
            u = u.substring(0, u.indexOf('?'));
        }
        if (u.indexOf('/_layouts/') !== -1) {
            u = u.substring(0, u.indexOf('/_layouts/'));
        }
        if (u.indexOf('/forms/') !== -1) {
            u = u.substring(0, u.indexOf('/forms/'));
            u = u.substring(0, u.lastIndexOf('/'));
        }
        if (u.indexOf('/lists/') !== -1) {
            u = u.substring(0, u.indexOf('/lists/'));
        }
        if (u.indexOf('/sitepages/') !== -1) {
            u = u.substring(0, u.indexOf('/sitepages/'));
        }
        if (u.endsWith(".aspx")) {
            u = u.substring(0, u.lastIndexOf('/'));
        }
        return u;
    }

    /**
     * Returns url that always ends with forward slash
     * @param url 
     */
    public static safeURL(url: string) {
        return url.endsWith("/") ? url : url + "/";
    }

    /**
     * Replace incompatible REST characters with OData equivalent
     * @param txt 
     */
    public static replaceSpecialCharacters(txt: string): string {
        return encodeURI(txt)
            .replace(/'/g, "''")
            //.replace(/%/g, "%25")
            .replace(/\+/g, "%2B")
            .replace(/\//g, "%2F")
            .replace(/\?/g, "%3F")
            .replace(/#/g, "%23")
            .replace(/&/g, "%26")
            .replace(/\(/g, "%28")
            .replace(/\)/g, "%29");
    }

    /**
     * Retrieve query string key pairs
     * @param queryString 
     */
    public static parseQueryString(): { [key: string]: string } {
        let returnValues: { [key: string]: string } = {};
        let queries = window.location.search.substring(1).split("&");
        for (let query of queries) {
            let queryPair = query.split("=", 2);
            let queryKey = decodeURIComponent(queryPair[0]);
            let queryValue = decodeURIComponent(queryPair.length === 2 ? queryPair[1] : "");
            returnValues[queryKey] = queryValue;
        }
        return returnValues;
    }

    /**
     * Convert bytes to the closest metric by size. Fixed to first decimal value.
     * 5 -> 5.0B
     * 2.300 -> 2.3KB
     * 1.841.231 -> 1.8MB
     * @param bytes 
     */
    public static closestByteMetric(bytes: number): string {
        let sizes = ['B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];

        if (!bytes) {
            return "0" + sizes[0];
        }
        let order: number = Math.min(sizes.length, Math.floor(Math.log(bytes) / Math.log(1024)));

        return (bytes / Math.pow(1024, order)).toFixed(1) + sizes[order];
    }

    /**
     * Convert Base64 from a UTF+BOM source with special characters
     * @param str 
     * @author https://stackoverflow.com/questions/30106476/using-javascripts-atob-to-decode-base64-doesnt-properly-decode-utf-8-strings
     */
    public static b64DecodeUnicode(base64Data: string) {
        // Going backwards: from bytestream, to percent-encoding, to original string.
        let decodedText: string = decodeURIComponent(atob(base64Data).split('').map((c) => {
            return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));
        if (decodedText.startsWith(Strings.UTFBOMStartCode)) {
            decodedText = decodedText.substring(Strings.UTFBOMStartCode.length);
        }
        return decodedText;
    }
}