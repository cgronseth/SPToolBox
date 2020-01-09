/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/spt.renderMenu.tsx");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/components/spt.menu.tsx":
/*!*************************************!*\
  !*** ./src/components/spt.menu.tsx ***!
  \*************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const React = __webpack_require__(/*! react */ "react");
const spt_constants_1 = __webpack_require__(/*! ../spt.constants */ "./src/spt.constants.ts");
const spt_logax_1 = __webpack_require__(/*! ../spt.logax */ "./src/spt.logax.ts");
const spt_sharepoint_rest_1 = __webpack_require__(/*! ../sharepoint/spt.sharepoint.rest */ "./src/sharepoint/spt.sharepoint.rest.ts");
const spt_strings_1 = __webpack_require__(/*! ../spt.strings */ "./src/spt.strings.ts");
class Menu extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            analizando: false,
            conectado: false,
            sharepointData: null
        };
        this.clickExplorador = this.clickExplorador.bind(this);
        this.clickDirectory = this.clickDirectory.bind(this);
    }
    render() {
        return React.createElement("div", { id: "SPT.Menu" },
            React.createElement("div", null,
                this.state.analizando &&
                    React.createElement("div", null,
                        React.createElement("img", { src: "icons/ajax-loader.gif", width: "12px" }),
                        " ",
                        spt_constants_1.Constants.getLiteral("menuAnalizandoSitio")),
                !this.state.analizando && this.state.conectado &&
                    React.createElement("div", null,
                        React.createElement("div", { className: 'SPTMenuDiagnosticsConnected' }, spt_constants_1.Constants.getLiteral("menuConectado")),
                        React.createElement("div", { className: 'SPTMenuDiagnosticsDetail' },
                            "\u00B7\u00A0",
                            spt_constants_1.Constants.getLiteral("menuTitulo"),
                            ": ",
                            this.state.sharepointData["Title"]),
                        React.createElement("div", { className: 'SPTMenuDiagnosticsDetail' },
                            "\u00B7\u00A0",
                            spt_constants_1.Constants.getLiteral("menuVersion"),
                            ": ",
                            this.state.sharepointData["UIVersion"]),
                        React.createElement("div", { className: 'SPTMenuDiagnosticsDetail' },
                            "\u00B7\u00A0",
                            spt_constants_1.Constants.getLiteral("menuLCID"),
                            ": ",
                            this.state.sharepointData["Language"])),
                !this.state.analizando && !this.state.conectado &&
                    React.createElement("div", null, spt_constants_1.Constants.getLiteral("menuNoSharePoint"))),
            this.state.conectado &&
                React.createElement("div", null,
                    React.createElement("hr", null),
                    React.createElement("div", { className: "SPTMenuItem", onClick: this.clickExplorador }, spt_constants_1.Constants.getLiteral("menuExplorador")),
                    React.createElement("div", { className: "SPTMenuItem", onClick: this.clickDirectory }, spt_constants_1.Constants.getLiteral("menuDirectory"))),
            React.createElement("hr", null),
            React.createElement("div", { className: "SPTMenuFooter" },
                "SharePoint Toolbox V",
                browser.runtime.getManifest().version));
    }
    componentDidUpdate(prevProps, prevState) {
        if (prevState.analizando != this.state.analizando) {
            let qry = spt_sharepoint_rest_1.SPRest.queryWebInfo(this.currentUrl);
            spt_logax_1.LogAx.trace("Query:" + qry);
            spt_sharepoint_rest_1.SPRest.restQuery(qry, spt_sharepoint_rest_1.RestQueryType.ODataJSON, 0).then((result) => {
                if (this.isCancelled)
                    return;
                try {
                    this.setState({
                        sharepointData: result,
                        analizando: false,
                        conectado: true
                    });
                }
                catch (e) {
                    this.setState({
                        sharepointData: null,
                        analizando: false,
                        conectado: false
                    });
                }
            }, (e) => {
                // Error mostly because this is not a SHP Site (or no permission)
                if (this.isCancelled)
                    return;
                this.setState({
                    sharepointData: null,
                    analizando: false,
                    conectado: false
                });
            });
        }
    }
    componentDidMount() {
        //Launch extension menú. Detects current tab to obtain possible SHP site
        browser.tabs.query({ active: true, windowId: browser.windows.WINDOW_ID_CURRENT })
            .then(tabs => browser.tabs.get(tabs[0].id))
            .then(tab => {
            console.info(tab);
            this.currentUrl = spt_strings_1.Strings.getWebUrlFromAbsolute(tab.url);
            this.setState({
                analizando: true
            });
        });
    }
    componentWillUnmount() {
        this.isCancelled = true;
    }
    clickExplorador(e) {
        let createData = {
            type: "panel",
            titlePreface: "SharePoint Toolbox - ",
            url: "spt.explorer.html?u=" + encodeURIComponent(this.currentUrl)
                + "&i=" + encodeURIComponent(this.state.sharepointData["Id"])
                + "&t=" + encodeURIComponent(this.state.sharepointData["Title"])
                + "&v=" + encodeURIComponent(this.state.sharepointData["menuVersion"])
                + "&l=" + encodeURIComponent(this.state.sharepointData["menuLCID"])
        };
        browser.windows.create(createData)
            .then((window) => {
            console.log("Panel " + window.id + " 'SPT.Explorer' created");
        });
        e.preventDefault();
    }
    clickDirectory(e) {
        let createData = {
            type: "panel",
            titlePreface: "SharePoint Toolbox - ",
            url: "spt.directory.html?u=" + encodeURIComponent(this.currentUrl)
                + "&i=" + encodeURIComponent(this.state.sharepointData["Id"])
                + "&t=" + encodeURIComponent(this.state.sharepointData["Title"])
                + "&v=" + encodeURIComponent(this.state.sharepointData["menuVersion"])
                + "&l=" + encodeURIComponent(this.state.sharepointData["menuLCID"])
        };
        browser.windows.create(createData)
            .then((window) => {
            console.log("Panel " + window.id + " 'SPT.Directory' created");
        });
        e.preventDefault();
    }
}
exports.Menu = Menu;


/***/ }),

/***/ "./src/sharepoint/spt.sharepoint.entities.ts":
/*!***************************************************!*\
  !*** ./src/sharepoint/spt.sharepoint.entities.ts ***!
  \***************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
class SPSecurableObject {
}
exports.SPSecurableObject = SPSecurableObject;
;
class SPData {
    constructor(internalName, value, type) {
        this.InternalName = internalName;
        this.StringValue = value;
        this.Type = type;
    }
}
exports.SPData = SPData;
class SPItem extends SPSecurableObject {
    constructor() {
        super(...arguments);
        this.Items = [];
    }
    getItem(internalName) {
        for (let i = 0; i < this.Items.length; i++) {
            let element = this.Items[i];
            if (element.InternalName === internalName) {
                return element;
            }
        }
        return null;
    }
}
exports.SPItem = SPItem;
class SPField {
}
exports.SPField = SPField;
class SPAttachment {
}
exports.SPAttachment = SPAttachment;
class SPFolder {
}
exports.SPFolder = SPFolder;
class SPListItem extends SPItem {
    constructor() {
        super(...arguments);
        this.Length = 0;
        this.Checked = 0; //Checked state: 0:Not checked, 1:Checked, -1:Partial checked (not all related items checked)
    }
}
exports.SPListItem = SPListItem;
class SPList extends SPSecurableObject {
}
exports.SPList = SPList;
class SPView {
}
exports.SPView = SPView;
class SPGroup {
}
exports.SPGroup = SPGroup;
class SPWeb extends SPSecurableObject {
}
exports.SPWeb = SPWeb;
class SPSite {
}
exports.SPSite = SPSite;
class SPUser {
    constructor(displayName, email) {
        this.DisplayName = displayName;
        this.Email = email;
    }
}
exports.SPUser = SPUser;
var PermissionKind;
(function (PermissionKind) {
    PermissionKind[PermissionKind["emptyMask"] = 0] = "emptyMask";
    PermissionKind[PermissionKind["viewListItems"] = 1] = "viewListItems";
    PermissionKind[PermissionKind["addListItems"] = 2] = "addListItems";
    PermissionKind[PermissionKind["editListItems"] = 3] = "editListItems";
    PermissionKind[PermissionKind["deleteListItems"] = 4] = "deleteListItems";
    PermissionKind[PermissionKind["approveItems"] = 5] = "approveItems";
    PermissionKind[PermissionKind["openItems"] = 6] = "openItems";
    PermissionKind[PermissionKind["viewVersions"] = 7] = "viewVersions";
    PermissionKind[PermissionKind["deleteVersions"] = 8] = "deleteVersions";
    PermissionKind[PermissionKind["cancelCheckout"] = 9] = "cancelCheckout";
    PermissionKind[PermissionKind["managePersonalViews"] = 10] = "managePersonalViews";
    PermissionKind[PermissionKind["manageLists"] = 12] = "manageLists";
    // Add or remove from SP.PermissionKind in sp.debug.js
})(PermissionKind = exports.PermissionKind || (exports.PermissionKind = {}));


/***/ }),

/***/ "./src/sharepoint/spt.sharepoint.rest.ts":
/*!***********************************************!*\
  !*** ./src/sharepoint/spt.sharepoint.rest.ts ***!
  \***********************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const spt_logax_1 = __webpack_require__(/*! ../spt.logax */ "./src/spt.logax.ts");
const spt_strings_1 = __webpack_require__(/*! ../spt.strings */ "./src/spt.strings.ts");
const spt_sharepoint_1 = __webpack_require__(/*! ./spt.sharepoint */ "./src/sharepoint/spt.sharepoint.ts");
var RestQueryType;
(function (RestQueryType) {
    RestQueryType[RestQueryType["ODataJSON"] = 0] = "ODataJSON";
    RestQueryType[RestQueryType["ArrayBuffer"] = 1] = "ArrayBuffer";
})(RestQueryType = exports.RestQueryType || (exports.RestQueryType = {}));
class SPRest {
    /**
    * Lanza consulta REST y devuelve los resultados en un Promise
    * @param qry Consulta completa: HTTPS://<servidor>/_api/ProjectServer/Projects('2c7d134c- [...]
    * @param reintentos Número de reintentos de 6s. Por defecto 10
    */
    static restQuery(qry, type, reintentos) {
        let tiempoReintento = 30000;
        if (reintentos === undefined) {
            reintentos = 10;
        }
        return new Promise((resolve, reject) => {
            let xhr = new XMLHttpRequest();
            xhr.open('GET', qry, true);
            if (type === RestQueryType.ODataJSON) {
                xhr.setRequestHeader("Accept", "application/json;odata=nometadata");
            }
            else if (type == RestQueryType.ArrayBuffer) {
                xhr.responseType = "arraybuffer";
            }
            xhr.onreadystatechange = () => {
                if (xhr.readyState === 4) {
                    //Debug. Random error. Descomentar para habilitar:
                    /*
                    let debugError: boolean = (Math.random() < 0.3); //30% error
                    */
                    if ( /*!debugError &&*/type === RestQueryType.ODataJSON && xhr.status === 200 && xhr.getResponseHeader('content-type').startsWith('application/json')) {
                        let result = JSON.parse(xhr.responseText);
                        let coolOffDelay = setTimeout(() => {
                            clearTimeout(coolOffDelay);
                            resolve(result);
                        }, Math.floor(Math.random() * 100) + 50);
                    }
                    else if (type === RestQueryType.ArrayBuffer && xhr.status === 200 && xhr.getResponseHeader('content-type') === "application/octet-stream") {
                        resolve(xhr.response);
                    }
                    else if (xhr.status === 401 || xhr.status === 403) {
                        reject('RESTQuery - Unauthorized/Forbidden');
                    }
                    else {
                        spt_logax_1.LogAx.trace('RESTQuery - Status==' + xhr.status.toString() + ': ' + xhr.responseText);
                        tiempoReintento = Math.floor(Math.random() * 1000) + tiempoReintento;
                        // Adapt query if field fails
                        if (xhr.status === 400) {
                            // Example responseText --> {\"odata.error\":{\"code\":\"-1, Microsoft.SharePoint.SPException\",\"message\":{\"lang\":\"en-US\",\"value\":\"The field or property 'AppAuthor' does not exist.\"}}}"
                            let matches = xhr.responseText.match(/property '(.*)' does not exist/);
                            if (matches.length === 2) {
                                // Erase field in query, be it a simple field or in a expand pair --> [",Author/Title", ",Author/EMail", "Author,"]
                                let rgx = new RegExp('(,' + matches[1] + '\/[a-zA-Z_0-9]+)|(' + matches[1] + ',)|(,' + matches[1] + ')', 'g');
                                qry = qry.replace(rgx, "");
                                // Not a connection problem, so restore retries and send off immediatly
                                tiempoReintento = 100;
                                reintentos++;
                                spt_logax_1.LogAx.trace("RESTQuery - Reintento eliminando el campo problemático '" + matches[1] + "'. Query:" + qry);
                            }
                        }
                        // Recursive retry
                        if (reintentos && --reintentos >= 0) {
                            spt_logax_1.LogAx.trace('RESTQuery - Reintentos restantes: ' + reintentos.toString());
                            let retryDelay = setTimeout(() => {
                                clearTimeout(retryDelay);
                                SPRest.restQuery(qry, type, reintentos).then((recResult) => {
                                    resolve(recResult);
                                }, (e) => {
                                    reject(e);
                                });
                            }, tiempoReintento);
                        }
                        else {
                            reject('RESTQuery - Reintentos agotados');
                        }
                    }
                }
            };
            xhr.onerror = () => {
                reject('RESTQuery - OnError: ' + xhr.responseText);
            };
            xhr.send();
        });
    }
    /**
     * Sends POST request
     * @param qry
     * @param requestDigestValue
     * @param newFolder
     * @param isDelete
     * @param fileData
     * @param listItemEntity
     * @param itemData
     */
    static restPost(qry, requestDigestValue, newFolder, isDelete, fileData, listItemEntity, itemData) {
        return new Promise((resolve, reject) => {
            let jsonBody = null;
            let spRequest = new XMLHttpRequest();
            spRequest.open('POST', qry, true);
            //cabeceras indicados en https://pwmather.wordpress.com/2018/05/21/using-rest-in-javascript-to-update-projectonline-project-custom-fields-ppm-pmot-jquery-office365/
            spRequest.setRequestHeader("Accept", "application/json;odata=verbose");
            if (requestDigestValue) {
                spRequest.setRequestHeader("X-RequestDigest", requestDigestValue);
            }
            if (!fileData) {
                spRequest.setRequestHeader("If-Match", "*");
            }
            if (newFolder) {
                jsonBody = JSON.stringify({ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': newFolder });
                spRequest.setRequestHeader("content-type", "application/json;odata=verbose");
            }
            if (isDelete) {
                spRequest.setRequestHeader("X-HTTP-Method", "DELETE");
            }
            if (itemData && listItemEntity) {
                let body = { '__metadata': { 'type': listItemEntity } };
                for (let spd of itemData) {
                    let internalName = spd.InternalName;
                    if (spd.Type === 7) { //Lookup
                        internalName += "Id";
                    }
                    body[internalName] = spt_sharepoint_1.SP.assignPostValue(spd);
                }
                jsonBody = JSON.stringify(body);
                spRequest.setRequestHeader("content-type", "application/json;odata=verbose");
            }
            spRequest.onreadystatechange = () => {
                if (spRequest.readyState === 4) {
                    if (spRequest.status >= 200 && spRequest.status < 300) {
                        if (spRequest.responseText !== '') {
                            resolve(JSON.parse(spRequest.responseText));
                        }
                        else if (spRequest.statusText.toLowerCase() === "ok") {
                            resolve({});
                        }
                        else {
                            reject("RestPost - Status: " + spRequest.status + ", StatusText: " + spRequest.statusText);
                        }
                    }
                    else {
                        reject('RestPost - Status==' + spRequest.status.toString() + ': ' + spRequest.responseText);
                    }
                }
            };
            spRequest.onerror = () => {
                reject('Resultados: restPost - Error: ' + spRequest.responseText);
            };
            if (fileData) {
                spRequest.send(fileData);
            }
            else if (jsonBody) {
                spRequest.send(jsonBody);
            }
            else {
                spRequest.send();
            }
        });
    }
    /**
     * Sets up POST request for adding Files
     * @param url
     * @param qry
     * @param data
     */
    static restPostAddFile(url, qry, data) {
        return new Promise((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, null, false, data).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }
    /**
     * Sets up POST request for adding Folders
     * @param url
     * @param qry
     * @param serverRelativefolder
     */
    static restPostAddFolder(url, qry, serverRelativefolder) {
        return new Promise((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, serverRelativefolder, false).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }
    /**
     * Sets up POST request for deleting files or folders
     * @param url
     * @param qry
     * @param data
     */
    static restPostDelete(url, qry, data) {
        return new Promise((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, null, true, data).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }
    /**
     * Sets up POST request for deleting items
     * @param url
     * @param qry
     * @param data
     */
    static restPostDeleteItem(url, qry) {
        return new Promise((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, null, true).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }
    /**
     * Sets up POST request for adding Files
     * @param url
     * @param qry
     * @param data
     */
    static restPostAddItem(url, qry, listItemEntity, data) {
        return new Promise((resolve, reject) => {
            this.requestDigestValue(url).then((requestDigestValue) => {
                this.restPost(qry, requestDigestValue, null, false, null, listItemEntity, data).then((postResult) => {
                    resolve(postResult);
                }, (e) => {
                    reject(e);
                });
            }, (e) => {
                reject(e);
            });
        });
    }
    /**
     * Read digest value necessary for POST and PUT updates
     * Since this code reads from outside a SharePoint site, a request is necessary
     * Apps in a SharePoint page can read this value directly from "__REQUESTDIGEST" input element
     * @param url
     */
    static requestDigestValue(url) {
        return new Promise((resolve, reject) => {
            if (this.cacheRequestDigest && this.cacheRequestDigestExpire > new Date().getTime()) {
                resolve(this.cacheRequestDigest);
            }
            let qry = SPRest.queryDigestValue(url);
            this.restPost(qry).then((resultadoJSON) => {
                let ctxWebInfo = resultadoJSON.d.GetContextWebInformation;
                let expireSeconds = +ctxWebInfo.FormDigestTimeoutSeconds; // 1800 seconds could be the limiting factor for most big uploads, before 2GB size limit. Just hope it doesn't expire in middle of operation
                this.cacheRequestDigest = ctxWebInfo.FormDigestValue;
                this.cacheRequestDigestExpire = new Date().setSeconds(expireSeconds - 60); // save in cache the expiration date (less a minute just in case)
                resolve(this.cacheRequestDigest);
            }, (e) => {
                reject("RequestDigestValue read Error: " + e);
            });
        });
    }
    /* Querys GET */
    static queryWebInfo(url) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web?$Select=Id,Title,UIVersion,Language";
    }
    static querySiteInfo(url) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Site?$Select=Id,Url";
    }
    static queryLists(url, baseType, showHidden) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists";
        //Select fields
        qry += "?$Select=Id,Title,Description,Hidden,EntityTypeName,ItemCount,Created,LastItemModifiedDate";
        //Where
        qry += "&$filter=BaseType eq " + baseType.toString() + " and Hidden eq " + (showHidden ? "true" : "false");
        //Order
        return qry + "&$OrderBy=Title asc";
    }
    static queryListsLight(url) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists?$Select=Id,Title,ItemCount,HasUniqueRoleAssignments&$filter=Hidden eq false";
    }
    static queryList(url, id) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + id + "')";
        //Select fields
        qry += "?$Select=Id,Title,Description,Hidden,EntityTypeName,ItemCount,Created,LastItemModifiedDate,EntityTypeName,ListItemEntityTypeFullName,RootFolder/Name,RootFolder/UniqueID,RootFolder/ItemCount,RootFolder/ServerRelativeUrl";
        qry += "&$Expand=RootFolder";
        return qry;
    }
    static queryListFields(url, idList) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Fields";
        //Select fields
        qry += "?$Select=Id,Title,InternalName,StaticName,Description,Hidden,Required,FieldTypeKind,LookupField";
        //Where
        qry += "&$filter=Hidden eq false";
        return qry;
    }
    static queryLibraryItemsWithView(url, idList, view) {
        let expands = ["Author", "Editor", "Folder", "Folder/ParentFolder", "File"];
        let fixedFields = ["Id", "ID", "EncodedAbsUrl", "FileRef", "FileLeafRef", "Folder", "File", "Author", "Created", "Editor", "Modified"];
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items";
        //Select fields. Long queries get cut off, so use wildcard selector for normal columns      
        if (view.ViewFields.length > 32) {
            qry += "?$Select=*,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Folder/Name,Folder/UniqueId,Folder/ItemCount,Folder/ServerRelativeUrl,Folder/ParentFolder/UniqueId,File/Length";
        }
        else {
            qry += "?$Select=Id,EncodedAbsUrl,FileRef,FileLeafRef,Created,Modified,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Folder/Name,Folder/UniqueId,Folder/ItemCount,Folder/ServerRelativeUrl,Folder/ParentFolder/UniqueId,File/Length";
        }
        for (let field of view.ViewFields) {
            if (fixedFields.indexOf(field.InternalName) !== -1)
                continue;
            let odataInternalName = spt_sharepoint_1.SP.safeOdataField(field.InternalName);
            switch (field.Type) {
                case 7: //Lookup
                    qry += "," + odataInternalName + "/" + spt_sharepoint_1.SP.safeOdataField(field.LookupField);
                    expands.push(odataInternalName);
                    break;
                case 20: //User
                    qry += "," + odataInternalName + "/Title," + odataInternalName + "/EMail";
                    expands.push(odataInternalName);
                    break;
                default:
                    if (view.ViewFields.length <= 32) { // not necessary with wildcard
                        qry += "," + odataInternalName;
                    }
                    break;
            }
        }
        //Select - Expand
        if (expands.length) {
            qry += "&$Expand=" + expands.join(",");
        }
        //Limit
        qry += "&$Top=" + ((view.RowLimit) ? view.RowLimit : 500);
        return qry;
    }
    static queryListItems(url, idList, rowLimit) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items";
        //Select fields
        qry += "?$Select=*,EncodedAbsUrl,FileRef,FileLeafRef,Author/Title,Author/EMail,Editor/Title,Editor/EMail";
        //Select - Expand
        qry += "&$Expand=Author,Editor";
        //Limit
        qry += "&$Top=" + ((rowLimit) ? rowLimit : 100);
        return qry;
    }
    static queryListItemsLight(url, idList) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items";
        //Select fields
        qry += "?$Select=Id,FileRef,HasUniqueRoleAssignments";
        //Limit
        qry += "&$Top=2000";
        return qry;
    }
    static queryListItemsWithView(url, idList, view) {
        let expands = ["Author", "Editor"];
        let fixedFields = ["Id", "ID", "Title", "Author", "Created", "Editor", "Modified"];
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items";
        //Select fields. Long queries get cut off, so use wildcard selector for normal columns      
        if (view.ViewFields.length > 32) {
            qry += "?$Select=*,Author/Title,Author/EMail,Editor/Title,Editor/EMail";
        }
        else {
            qry += "?$Select=Id,Title,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Created,Modified";
        }
        for (let field of view.ViewFields) {
            if (fixedFields.indexOf(field.InternalName) !== -1)
                continue;
            let odataInternalName = field.InternalName.startsWith("_x") ? "OData_" + field.InternalName : field.InternalName;
            switch (field.Type) {
                case 7: //Lookup
                    let odataLookupField = field.LookupField.startsWith("_x") ? "OData_" + field.LookupField : field.LookupField;
                    qry += "," + odataInternalName + "/" + odataLookupField + "," + odataInternalName + "/ID";
                    expands.push(odataInternalName);
                    break;
                case 20: //User
                    qry += "," + odataInternalName + "/Title," + odataInternalName + "/EMail";
                    expands.push(odataInternalName);
                    break;
                default:
                    if (view.ViewFields.length <= 32) { // not necessary with wildcard
                        qry += "," + odataInternalName;
                    }
                    break;
            }
        }
        //Select - Expand
        if (expands.length) {
            qry += "&$Expand=" + expands.join(",");
        }
        //Limit
        qry += "&$Top=" + ((view.RowLimit) ? view.RowLimit : 100);
        return qry;
    }
    static queryFileData(url, idList, idItem) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Items(" + idItem + ")/File/$value";
    }
    static queryWebs(url) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Webs";
        //Select fields
        qry += "?$Select=Id,Title,Url,ServerRelativeUrl";
        //Order
        return qry + "&$OrderBy=Title asc";
    }
    static queryWeb(url) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web";
        //Select fields
        qry += "?$Select=Id,Title,Description,Url,ServerRelativeUrl,Created";
        //Order
        return qry;
    }
    static queryDigestValue(url) {
        return spt_strings_1.Strings.safeURL(url) + "_api/contextinfo";
    }
    static queryFolder(url, path) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/GetFolderByServerRelativeUrl('" + path + "')/Folders?$Select=Name";
    }
    static queryViews(url, idList) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/Views";
        //Select fields
        qry += "?$Select=Id,Title,DefaultView,Paged,PersonalView,RowLimit,ServerRelativeUrl,ViewFields/Items";
        //Expando to other api methods
        qry += "&$Expand=ViewFields";
        //Filter
        qry += "&$filter=Hidden eq false";
        return qry;
    }
    static querySiteUsers(url) {
        return spt_strings_1.Strings.safeURL(url) + "_api/web/SiteUsers?$Select=Id,Title,Email,IsSiteAdmin,UserPrincipalName";
    }
    static querySiteGroupsUsers(url) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/SiteGroups?$Select=Id,Title,PrincipalType,Users/Id,Users/Title,Users/Email,Users/IsSiteAdmin&$Expand=Users";
    }
    static queryWebPermissions(url) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/RoleAssignments?$Select=PrincipalId,RoleDefinitionBindings/BasePermissions&$Expand=RoleDefinitionBindings";
    }
    static queryWebPermissionsForUser(url, idUser) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/RoleAssignments/GetByPrincipalId(" + idUser + ")/RoleDefinitionBindings?$Select=BasePermissions";
    }
    static queryListPermissionsForUser(url, listId, idUser) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + listId + "')/RoleAssignments/GetByPrincipalId(" + idUser + ")/RoleDefinitionBindings?$Select=BasePermissions";
    }
    static queryItemPermissionsForUser(url, listId, itemId, idUser) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + listId + "')/Items(" + itemId + ")/RoleAssignments/GetByPrincipalId(" + idUser + ")/RoleDefinitionBindings?$Select=BasePermissions";
    }
    static queryPermissionsList(url, listId) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + listId + "')/EffectiveBasePermissions";
    }
    /* Querys POST */
    static queryPostFolders(url) {
        return spt_strings_1.Strings.safeURL(url) + "_api/web/folders";
    }
    static queryPostAddFile(url, path, name) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/GetFolderByServerRelativeUrl('" + path + "')/Files/Add(overwrite=true, url='" + name + "')";
    }
    static queryPostDeleteFolder(url, pathAndName, recycle) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/GetFolderByServerRelativeUrl('" + pathAndName + "')";
        if (recycle) {
            qry += "/Recycle()";
        }
        return qry;
    }
    static queryPostDeleteFile(url, pathAndName, recycle) {
        let qry = spt_strings_1.Strings.safeURL(url) + "_api/Web/GetFileByServerRelativeUrl('" + pathAndName + "')";
        if (recycle) {
            qry += "/Recycle()";
        }
        return qry;
    }
    static queryPostDeleteItem(url, idList, idItem) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/items(" + idItem.toString() + ")";
    }
    static queryPostAddItem(url, idList) {
        return spt_strings_1.Strings.safeURL(url) + "_api/Web/Lists(guid'" + idList + "')/items";
    }
}
exports.SPRest = SPRest;
/* No utiliza de momento
public static restPut(qry: string, data?: string): Promise<any> {
    return new Promise<void>((resolve, reject) => {
        var spRequest = new XMLHttpRequest();
        spRequest.open('PUT', qry, true);
        spRequest.setRequestHeader("Accept", "application/json");
        spRequest.onreadystatechange = () => {
            if (spRequest.readyState === 4 && spRequest.status === 200) {
                var result = JSON.parse(spRequest.responseText);
                resolve(result);
            } else if (spRequest.readyState === 4 && spRequest.status !== 200) {
                LogAx.trace('Resultados: restPut - Status==' + spRequest.status.toString() + ': ' + spRequest.responseText);
                reject();
            }
        };
        spRequest.onerror = () => {
            LogAx.trace('Resultados: restPut - Error: ' + spRequest.responseText);
            reject();
        };
        if (data) {
            spRequest.send(JSON.stringify(data));
        } else {
            spRequest.send();
        }
    });
}*/
/**
 * DigestValue cache: avoid sending a request if already available and not expired
 */
SPRest.cacheRequestDigest = null;
SPRest.cacheRequestDigestExpire = null;


/***/ }),

/***/ "./src/sharepoint/spt.sharepoint.ts":
/*!******************************************!*\
  !*** ./src/sharepoint/spt.sharepoint.ts ***!
  \******************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const spt_sharepoint_entities_1 = __webpack_require__(/*! ./spt.sharepoint.entities */ "./src/sharepoint/spt.sharepoint.entities.ts");
const spt_logax_1 = __webpack_require__(/*! ../spt.logax */ "./src/spt.logax.ts");
const spt_constants_1 = __webpack_require__(/*! ../spt.constants */ "./src/spt.constants.ts");
class SP {
    // Order SPListItem array by their fields. Sorting fields are passed in "fields" array.
    // If a field ends with a dot (.) it orders Descending. Otherwise Ascending
    static orderLibraryDataByFields(data, fields) {
        if (fields && fields.length && data && data.length) {
            data.sort((a, b) => {
                return SP.orderLibraryDataByFieldsRecursive(a, b, fields, 0);
            });
        }
        return data;
    }
    static orderLibraryDataByFieldsRecursive(a, b, fields, idx) {
        if (fields.length == idx || idx > 5)
            return 0;
        let field = fields[idx];
        let direction = 1; //1=ASC, -1=DESC
        if (field.endsWith('.')) {
            field = field.slice(0, -1);
            direction = -1;
        }
        //Obtain sortable data from field. Data must include value and type.
        let objA;
        let objB;
        switch (field) {
            case "ID":
                objA = new spt_sharepoint_entities_1.SPData(field, a.ID.toString(), 5);
                objB = new spt_sharepoint_entities_1.SPData(field, b.ID.toString(), 5);
                break;
            case "Author":
                objA = new spt_sharepoint_entities_1.SPData(field, a.Author.DisplayName, 255);
                objB = new spt_sharepoint_entities_1.SPData(field, b.Author.DisplayName, 255);
                break;
            case "Editor":
                objA = new spt_sharepoint_entities_1.SPData(field, a.Editor.DisplayName, 255);
                objB = new spt_sharepoint_entities_1.SPData(field, b.Editor.DisplayName, 255);
                break;
            case "Created":
                objA = new spt_sharepoint_entities_1.SPData(field, a.Created.toISOString(), 4);
                objB = new spt_sharepoint_entities_1.SPData(field, b.Created.toISOString(), 4);
                break;
            case "Modified":
                objA = new spt_sharepoint_entities_1.SPData(field, a.Modified.toISOString(), 4);
                objB = new spt_sharepoint_entities_1.SPData(field, b.Modified.toISOString(), 4);
                break;
            case "Folder":
                // Los ítems normales no tiene Folder, pero sí ParentFolder. 
                // Agregar carácter | a los ítems normales para que ordene siempre debajo de su carpeta padre.
                if (a.Folder.Name) {
                    objA = new spt_sharepoint_entities_1.SPData(field, a.Folder.ServerRelativeUrl, 255);
                }
                else {
                    objA = new spt_sharepoint_entities_1.SPData(field, a.Folder.ParentFolder.ServerRelativeUrl + "|", 255);
                }
                if (b.Folder.Name) {
                    objB = new spt_sharepoint_entities_1.SPData(field, b.Folder.ServerRelativeUrl, 255);
                }
                else {
                    objB = new spt_sharepoint_entities_1.SPData(field, b.Folder.ParentFolder.ServerRelativeUrl + "|", 255);
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
                let na = +objA.StringValue;
                let nb = +objB.StringValue;
                if (na < nb)
                    return -1 * direction;
                if (na > nb)
                    return 1 * direction;
                break;
            case 4: //Date
                let da = new Date(objA.StringValue);
                let db = new Date(objB.StringValue);
                if (da < db)
                    return -1 * direction;
                if (da > db)
                    return 1 * direction;
                break;
            case 8: //Boolean
                let ba = objA.StringValue.toLowerCase() === "true";
                let bb = objB.StringValue.toLowerCase() === "true";
                if (!ba && nb)
                    return -1 * direction;
                if (ba && !bb)
                    return 1 * direction;
                break;
            case 14: //GUID
                if (objA.StringValue.toLowerCase() < objB.StringValue.toLowerCase())
                    return -1 * direction;
                if (objA.StringValue.toLowerCase() > objB.StringValue.toLowerCase())
                    return 1 * direction;
                break;
            default: //Strings
                if (objA.StringValue < objB.StringValue)
                    return -1 * direction;
                if (objA.StringValue > objB.StringValue)
                    return 1 * direction;
                break;
        }
        // If execution reaches here they are equal, sort by next field
        return SP.orderLibraryDataByFieldsRecursive(a, b, fields, idx + 1);
    }
    static findFolderByPath(cacheFolders, fileRef) {
        if (fileRef) {
            let fileRefLower = fileRef.toLowerCase();
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
    static getFolderPath(folder, currentPath) {
        return folder ? SP.getFolderPath(folder.ParentFolder, folder.Name + "/" + currentPath) : currentPath;
    }
    /**
     * From folder obtain path structure. Omit parent library in path.
     * <sub folder>/<current path>
     * @param folder
     * @param currentPath
     */
    static getFolderPathWithoutParent(folder, currentPath) {
        let path = this.getFolderPath(folder, currentPath);
        if (path.indexOf('/') !== -1) {
            path = path.substring(path.indexOf('/') + 1); //Remove first path and leading slash
            if (path.indexOf('/') === -1) {
                path = "/" + path; //Add forward slash if path was root without folders and was removed in previous step
            }
        }
        return path;
    }
    /**
     * Get maximum folder depth level
     */
    static getFolderDepthLevel(items) {
        let levels = items
            .filter(i => i.Folder !== null && i.Folder.Level !== null)
            .map(i => i.Folder.Level);
        return levels.length ? Math.max(...levels) : 0;
    }
    static parseItemJsonResult(i, field) {
        let spd = new spt_sharepoint_entities_1.SPData(field.InternalName, "", field.Type);
        try {
            let odataInternalName = SP.safeOdataField(field.InternalName);
            if (i[odataInternalName] !== null && i[odataInternalName] !== undefined) {
                switch (field.Type) {
                    case 7: //Lookup
                        spd.StringValue = i[odataInternalName][SP.safeOdataField(field.LookupField)];
                        spd.LookupId = i[odataInternalName]["ID"];
                        break;
                    case 20: //User
                        spd.StringValue = i[odataInternalName]["EMail"];
                        break;
                    case 8: //Yes/No
                        spd.StringValue = i[odataInternalName] ? SP.literalYes : SP.literalNo;
                        break;
                    default:
                        spd.StringValue = i[odataInternalName] + "";
                        break;
                }
            }
        }
        catch (_a) {
            spt_logax_1.LogAx.trace("Error parsing column '" + field.InternalName + "' on item id:" + i.ID);
        }
        return spd;
    }
    /**
     * Apply OData renaming to fields with special characters. Just necessary in Rest queries and results
     * @param field
     */
    static safeOdataField(field) {
        return (!!field && field.startsWith("_x")) ? "OData_" + field : field;
    }
    /**
     * Adapt display values to SharePoint accepted input values
     * @param stringValue
     * @param type
     * @description Sometimes SharePoint doesn't accept its own values, so they must be converted, like boolean from "Yes"->"true"/"No"->"false",
     * or Lookup values use IDs
     */
    static assignPostValue(data) {
        switch (data.Type) {
            case 8: //Boolean
                return (data.StringValue.toLowerCase() === "no") ? false : true;
            case 7: //Lookup
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
    static checkEffectivePermission(high, low, permission) {
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
exports.SP = SP;
SP.literalYes = spt_constants_1.Constants.getLiteral("generalSi");
SP.literalNo = spt_constants_1.Constants.getLiteral("generalNo");


/***/ }),

/***/ "./src/spt.constants.ts":
/*!******************************!*\
  !*** ./src/spt.constants.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
class Constants {
    static getLiteral(id) {
        let text;
        if (this.currentLCID === null) {
            this.currentLCID = Constants.getLCID();
        }
        if (this.currentLCID === "es") {
            text = Constants.ES[id];
        }
        else {
            text = Constants.EN[id];
        }
        return (!text) ? id : text; //if undefined (doesn't exist) just throw out the id string, easier to spot than a blank space
    }
    static getLCID() {
        let lcid = browser.i18n.getUILanguage();
        if (lcid.length > 2) {
            lcid = lcid.substr(0, 2);
        }
        return lcid;
    }
}
exports.Constants = Constants;
Constants.ES = {
    analisysError: "Se han encontrado incompatibilidades. Seguir con la copia puede producir inconsistencias.",
    analisysErrorInternalName: "[Error] Nombre interno no encontrado: %1",
    analisysErrorInternalNameAndType: "[Advertencia] Campo '%1' no tiene el mismo tipo de dato",
    analisysErrorRequired: "[Error] Campo '%1' es requerido en destino",
    directory: "Directorio SharePoint",
    directoryGroups: "Grupos asignados",
    directoryNoGroups: "No asignado a ningún grupo",
    directoryPermissionsSite: "Permisos del sitio",
    directorySearchPlaceholder: "Email, nombre completo o parcial",
    directorySearchResults: "Resultados de la búsqueda",
    directorySearchUser: "Usuario",
    directoryTableWebColumn: "Web",
    directoryTableListColumn: "Lista",
    directoryTableItemColumn: "Elemento",
    directoryTableReadColumn: "Leer",
    directoryTableWriteColumn: "Escribir",
    directoryTableDeleteColumn: "Borrar",
    directoryTitleFiltro: "Búsqueda",
    directoryTitlePermissions: "Permisos",
    explorerBibliotecaTitulo: "Biblioteca",
    explorerBibliotecaInternal: "Nombre interno",
    explorerBibliotecaEntityType: "Nombre entidad",
    explorerBibliotecas: "Bibliotecas",
    explorerCargando: "Cargando",
    explorerListas: "Listas",
    explorerNoRows: "Lista vacía",
    explorerSubsitios: "Sub-Sitios",
    explorerSubsitiosEstructura: "Mostrar estructura",
    explorerHelp: "Seleccione un recurso del menú lateral",
    explorerMenuBotonCopiar: "Copiar",
    explorerMenuBotonDescargar: "Descargar",
    explorerMenuBotonEliminar: "Eliminar",
    explorerMenuBotonPegar: "Pegar",
    explorerMenuBotonPegarFolder: "Pegar en carpeta",
    explorerMenuBotonRefrescar: "Refrescar",
    explorerMenuBotonSubirCSV: "Importar CSV",
    explorerMenuDesplegableVista: "Vista",
    explorerModalComprimir: "Compresión de archivo",
    explorerModalCopyTitle: "Copiar y Pegar",
    explorerModalCopyMessage: "Copiando archivo {%1} / {%2} ({%3} Errores)",
    explorerModalCopyMessageList: "Copiando ítem {%1} / {%2} ({%3} Errores)",
    explorerModalCopyMessageFinished: "Copia finalizada {%1} / {%2} ({%3} Errores)",
    explorerModalCopyMessageAnalysing: "Analizando lista destino...",
    explorerModalCopyMessageConfirm: "ADVERTENCIA: Se va a pegar el contenido seleccionado creando las carpetas y sobreescribiendo archivos existentes. ¿Aceptar y continuar?",
    explorerModalCopyMessageConfirmList: "ADVERTENCIA: Se va a pegar el contenido seleccionado en esta lista. ¿Aceptar y continuar?",
    explorerModalCopyMessageFolderConfirm: "ADVERTENCIA: Se va a pegar el contenido seleccionado ignorando la estructura de carpetas y sobreescribiendo archivos existentes en la carpeta seleccionada. ¿Aceptar y continuar?",
    explorerModalDeleteTitle: "Eliminación",
    explorerModalDeleteMessage: "Eliminando archivo {%1} / {%2} ({%3} Errores)",
    explorerModalDeleteMessageList: "Eliminando ítem {%1} / {%2} ({%3} Errores)",
    explorerModalDeleteMessageFinished: "Eliminación finalizada {%1} / {%2} ({%3} Errores)",
    explorerModalDeleteMessageConfirm: "ADVERTENCIA: Se va a eliminar el contenido seleccionado. ¿Aceptar y continuar?",
    explorerModalDescargar: "Descarga de archivo",
    explorerModalDescargarErrores: "Errores de descarga",
    explorerModalDescargarTitulo: "Descargando archivos",
    explorerModalDescargarCompletado: "Descargado",
    explorerModalExcelTitle: "Descargar hoja Excel",
    explorerModalExcelMessage: "Todas las filas seleccionadas de la vista actual se descargarán como un fichero separado por coma (csv), compatible con Excel.",
    explorerModalExcelMessage2: "Para una correcta compatiblidad con Excel, debe seleccionar el carácter separador correcto para el idioma de tu sistema.",
    explorerModalExcelMessage3: "Separador",
    explorerModalExcelOption1: "Coma (,)",
    explorerModalExcelOption2: "Punto y Coma (;)",
    explorerModalImportExcelTitle: "Importar hoja Excel",
    explorerModalImportExcelMessage: "Seleccionar el archivo CSV que se desea importar en esta lista.",
    explorerModalImportExcelMessage2: "Para una correcta carga, debe seleccionar el carácter separador utilizado en el archivo.",
    explorerModalImportExcelMessage3: "Comenzar la importación de datos en SharePoint.",
    explorerModalImportExcelMessage4: "La información leída del fichero será importada en la lista.",
    explorerModalImportExcelMessageErrorsFound: "¡Atención! Errores encontrados. Pueden causar corrupción en la información. Se recomienda corregir estas incidencias antes de continuar.",
    explorerModalImportExcelMessage5: "Importando información a SharePoint.",
    explorerModalImportExcelMessage6: "Esto puede tardar un momento. Puedes cancelar la operación en cualquier momento.",
    explorerModalImportExcelMessage7: "Importación finalizada.",
    explorerModalImportExcelAnalyze: "Cargar fichero",
    explorerModalImportExcelMessageProcessing: "Procesando",
    explorerModalImportExcelMessageReading: "Fichero leído",
    explorerModalImportExcelMessageUploading: "Importado en la lista",
    explorerModalImportExcelError: "Error de fichero",
    explorerModalImportExcelError01: "No es del tipo esperado",
    libraryViewerCargando: "Cargando bilblioteca",
    libraryViewerCargandoItems: "Cargando datos de la biblioteca",
    libraryViewerCargarListaCompleta: "Cargar lista completa",
    libraryViewerFolder: "Carpeta",
    libraryViewerItemsCargados: "Items cargados",
    libraryViewerTotalItems: "Items en total",
    menuAccesoDenegado: "Access Denied",
    menuAnalizandoSitio: "Analizando Sitio",
    menuConectado: "Conectado",
    menuDirectory: "Directorio",
    menuExplorador: "Explorador",
    menuLCID: "LCID",
    menuNoSharePoint: "Sin Conexión",
    menuTitulo: "Título",
    menuVersion: "Versión UI",
    webViewerCargando: "Cargando web",
    webViewerDetailsUnauzorized: "Web acceso no autorizado",
    webViewerDetailsExploreWeb: "Explorar",
    webViewerDetailsLevelLimit: "Alcanzado límite de 100 niveles de subsitio",
    generalCreado: "Creado",
    generalDescripcion: "Descripción",
    generalSitio: "Sitio",
    generalBotonAceptar: "Aceptar",
    generalBotonCerrar: "Cerrar",
    generalBotonCancelar: "Cancelar",
    generalBotonDescargar: "Descargar",
    generalBotonCargar: "Cargar",
    generalBotonImportar: "Importar",
    generalCargando: "Cargando...",
    generalTrue: "Verdadero",
    generalFalse: "Falso",
    generalSi: "Sí",
    generalNo: "No",
    generalNoData: "No data"
};
Constants.EN = {
    analisysError: "Incompatibilities found. Proceed to copy may cause issues.",
    analisysErrorInternalName: "[Error] Internal name not found: %1",
    analisysErrorInternalNameAndType: "[Warning] The field '%1' is not the same data type",
    analisysErrorRequired: "[Error] Field '%1' is required on target list",
    directory: "SharePoint Directory",
    directoryGroups: "Assigned groups",
    directoryNoGroups: "Not assigned to any group",
    directoryPermissionsSite: "Site permissions",
    directorySearchPlaceholder: "Email, full or partial name",
    directorySearchResults: "Search results",
    directorySearchUser: "Usuario",
    directoryTableWebColumn: "Web",
    directoryTableListColumn: "List",
    directoryTableItemColumn: "Item",
    directoryTableReadColumn: "Read",
    directoryTableWriteColumn: "Write",
    directoryTableDeleteColumn: "Delete",
    directoryTitleFiltro: "Search",
    directoryTitlePermissions: "Permissions",
    explorerBibliotecaTitulo: "Library",
    explorerBibliotecaInternal: "Internal name",
    explorerBibliotecaEntityType: "Full entity",
    explorerBibliotecas: "Libraries",
    explorerCargando: "Loading",
    explorerListas: "Lists",
    explorerNoRows: "Empty list",
    explorerSubsitios: "Sub-Sites",
    explorerSubsitiosEstructura: "Show structure",
    explorerHelp: "Select a resource from the left panel",
    explorerMenuBotonCopiar: "Copy",
    explorerMenuBotonDescargar: "Download",
    explorerMenuBotonEliminar: "Delete",
    explorerMenuBotonPegar: "Paste",
    explorerMenuBotonPegarFolder: "Paste in folder",
    explorerMenuBotonRefrescar: "Refresh",
    explorerMenuBotonSubirCSV: "Import CSV",
    explorerMenuDesplegableVista: "View",
    explorerModalComprimir: "Compressing file",
    explorerModalCopyTitle: "Copy & Paste",
    explorerModalCopyMessage: "Copying file {%1} / {%2} ({%3} Errors)",
    explorerModalCopyMessageList: "Copying item {%1} / {%2} ({%3} Errors)",
    explorerModalCopyMessageFinished: "Copy finished {%1} / {%2} ({%3} Errors)",
    explorerModalCopyMessageAnalysing: "Analyzing target list...",
    explorerModalCopyMessageConfirm: "WARNING: Selected content will be saved creating folders and overwriting existing files. Accept and continue?",
    explorerModalCopyMessageFolderConfirm: "WARNING: Selected content will be saved ignoring folder structure and overwriting existing files in the selected folder. Accept and continue?",
    explorerModalDeleteTitle: "Delete",
    explorerModalDeleteMessage: "Deleting file {%1} / {%2} ({%3} Errors)",
    explorerModalDeleteMessageList: "Deleting item {%1} / {%2} ({%3} Errors)",
    explorerModalDeleteMessageFinished: "Delete finished {%1} / {%2} ({%3} Errors)",
    explorerModalDeleteMessageConfirm: "WARNING: Selected content will be deleted. Accept and continue?",
    explorerModalDescargar: "Downloading file",
    explorerModalDescargarErrores: "Download Errors",
    explorerModalDescargarTitulo: "Downloading files",
    explorerModalDescargarCompletado: "Downloaded",
    explorerModalExcelTitle: "Download Excel sheet",
    explorerModalExcelMessage: "All selected rows in the current view will be downloaded as a comma-separated-value (csv) file compatible with Excel.",
    explorerModalExcelMessage2: "To ensure compatibility, Excel requires you to use the proper delimiter for your systems language.",
    explorerModalExcelMessage3: "Delimiter",
    explorerModalExcelOption1: "Comma (,)",
    explorerModalExcelOption2: "Semicolon (;)",
    explorerModalImportExcelTitle: "Import Excel sheet",
    explorerModalImportExcelMessage: "Select the CSV file you wish to import into this list.",
    explorerModalImportExcelMessage2: "To ensure proper loading, we require you to specify the delimiter used in the file.",
    explorerModalImportExcelMessage3: "Start uploading data to SharePoint.",
    explorerModalImportExcelMessage4: "All data read from the file will be uploaded into the list.",
    explorerModalImportExcelMessageErrorsFound: "Important! Please check these errors found. They could cause data corruption. It is recommended to fix these problems before continuing.",
    explorerModalImportExcelMessage5: "Uploading data to SharePoint.",
    explorerModalImportExcelMessage6: "This make take a while. You may cancel at any time.",
    explorerModalImportExcelMessage7: "Import finished.",
    explorerModalImportExcelAnalyze: "Read file",
    explorerModalImportExcelMessageProcessing: "Processing",
    explorerModalImportExcelMessageReading: "File read",
    explorerModalImportExcelMessageUploading: "Uploaded to the list",
    explorerModalImportExcelError: "File error",
    explorerModalImportExcelError01: "Not of the type expected",
    libraryViewerCargando: "Loading library",
    libraryViewerCargandoItems: "Loading library data",
    libraryViewerCargarListaCompleta: "Load full list",
    libraryViewerFolder: "Folder",
    libraryViewerItemsCargados: "Loaded Items",
    libraryViewerTotalItems: "Total Items",
    menuAccesoDenegado: "Acceso Denegado",
    menuAnalizandoSitio: "Analyzing Site",
    menuConectado: "Connected",
    menuDirectory: "Directory",
    menuExplorador: "Explorer",
    menuLCID: "LCID",
    menuNoSharePoint: "Not Connected",
    menuTitulo: "Title",
    menuVersion: "UI Version",
    webViewerCargando: "Cargando web",
    webViewerDetailsUnauzorized: "Web access not autorized",
    webViewerDetailsExploreWeb: "Explore",
    webViewerDetailsLevelLimit: "Subsite levels limited to 100",
    generalCreado: "Created",
    generalDescripcion: "Description",
    generalSitio: "Site",
    generalBotonAceptar: "Accept",
    generalBotonCerrar: "Close",
    generalBotonCancelar: "Cancel",
    generalBotonDescargar: "Download",
    generalBotonCargar: "Load",
    generalBotonImportar: "Import",
    generalCargando: "Loading...",
    generalTrue: "True",
    generalFalse: "False",
    generalSi: "Yes",
    generalNo: "No",
    generalNoData: "No data"
};
Constants.currentLCID = null;


/***/ }),

/***/ "./src/spt.dates.ts":
/*!**************************!*\
  !*** ./src/spt.dates.ts ***!
  \**************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
class Dates {
    /**
     * Get current time in format "00:00:00.000" used for log
     */
    static getTimestampPrefix() {
        let d = new Date();
        let hours = ("0" + d.getHours()).slice(-2);
        let minutes = ("0" + d.getMinutes()).slice(-2);
        let seconds = ("0" + d.getSeconds()).slice(-2);
        let miliseconds = ("00" + d.getMilliseconds()).slice(-3);
        return hours + ":" + minutes + ":" + seconds + "." + miliseconds + " ";
    }
    /**
     * Get current time in format "yyyy-mm-dd-hh-mm" used for filenames
     */
    static getFileSuffix() {
        let d = new Date();
        let month = ("0" + d.getMonth()).slice(-2);
        let day = ("0" + d.getDate()).slice(-2);
        let hours = ("0" + d.getHours()).slice(-2);
        let minutes = ("0" + d.getMinutes()).slice(-2);
        return d.getFullYear() + "-" + month + "-" + day + "-" + hours + "-" + minutes;
    }
}
exports.Dates = Dates;


/***/ }),

/***/ "./src/spt.logax.ts":
/*!**************************!*\
  !*** ./src/spt.logax.ts ***!
  \**************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const spt_dates_1 = __webpack_require__(/*! ./spt.dates */ "./src/spt.dates.ts");
const appName = "SharePoint Toolbox";
/**
 * Logging helper class. Extend to write to file, database, etc.
 */
class LogAx {
    //Escribe traza verbose a consola y cualquier otro medio futuro
    static trace(txt) {
        if (LogAx.TRACE) {
            //let t = LogAx.groupTexts[appName];
            //LogAx.groupTexts[appName] = (!t) ? Dates.getTimestampPrefix() + txt : t + '\n' + Dates.getTimestampPrefix() + txt;
            console.log("<" + appName + ">" + "[" + spt_dates_1.Dates.getTimestampPrefix() + "]: " + txt);
        }
    }
}
exports.LogAx = LogAx;
//Habilitar trazas informativas. TODO: pasar a configurable/automático
LogAx.TRACE = true;


/***/ }),

/***/ "./src/spt.renderMenu.tsx":
/*!********************************!*\
  !*** ./src/spt.renderMenu.tsx ***!
  \********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const React = __webpack_require__(/*! react */ "react");
const ReactDOM = __webpack_require__(/*! react-dom */ "react-dom");
const spt_menu_1 = __webpack_require__(/*! ./components/spt.menu */ "./src/components/spt.menu.tsx");
ReactDOM.render(React.createElement(spt_menu_1.Menu, null), document.getElementById("spt.menu"));


/***/ }),

/***/ "./src/spt.strings.ts":
/*!****************************!*\
  !*** ./src/spt.strings.ts ***!
  \****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
class Strings {
    /**
     * Obtains URL closest to the absolute URL
     * @param url
     */
    static getWebUrlFromAbsolute(url) {
        let u = url.toLowerCase().trim();
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
    static safeURL(url) {
        return url.endsWith("/") ? url : url + "/";
    }
    /**
     * Replace incompatible REST characters with OData equivalent
     * @param txt
     */
    static replaceSpecialCharacters(txt) {
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
    static parseQueryString() {
        let returnValues = {};
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
    static closestByteMetric(bytes) {
        let sizes = ['B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
        if (!bytes) {
            return "0" + sizes[0];
        }
        let order = Math.min(sizes.length, Math.floor(Math.log(bytes) / Math.log(1024)));
        return (bytes / Math.pow(1024, order)).toFixed(1) + sizes[order];
    }
    /**
     * Convert Base64 from a UTF+BOM source with special characters
     * @param str
     * @author https://stackoverflow.com/questions/30106476/using-javascripts-atob-to-decode-base64-doesnt-properly-decode-utf-8-strings
     */
    static b64DecodeUnicode(base64Data) {
        // Going backwards: from bytestream, to percent-encoding, to original string.
        let decodedText = decodeURIComponent(atob(base64Data).split('').map((c) => {
            return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));
        if (decodedText.startsWith(Strings.UTFBOMStartCode)) {
            decodedText = decodedText.substring(Strings.UTFBOMStartCode.length);
        }
        return decodedText;
    }
}
exports.Strings = Strings;
Strings.UTFBOMStartCode = "\ufeff";


/***/ }),

/***/ "react":
/*!************************!*\
  !*** external "React" ***!
  \************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = React;

/***/ }),

/***/ "react-dom":
/*!***************************!*\
  !*** external "ReactDOM" ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = ReactDOM;

/***/ })

/******/ });
//# sourceMappingURL=spt.renderMenu.js.map