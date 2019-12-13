# SPToolBox

SharePoint Tools for managing Lists and Libraries.

## Project setup

### IDE setup

Apps & versions used for development:

- Visual Studio Code 1.40.2
- Git 2.19.2.windows.1
- NodeJS 12.13.1 / NPM 6.12.1
- Windows 8.1

### Code setup

Clone from Github:

```node
C:\>cd WebExtensionsProjects
C:\WebExtensionsProjects>git clone https://github.com/cgronseth/SPToolBox.git
C:\WebExtensionsProjects>cd SPToolBox
C:\WebExtensionsProjects\SPToolBox>npm i

[... wait install modules ...]

C:\WebExtensionsProjects\SPToolBox>Code .
```

From SPToolbox.zip:

Extract to local folder.

```node
C:\>cd WebExtensionsProjects\SPToolBox
C:\WebExtensionsProjects\SPToolBox>npm i

[... wait install modules ...]

C:\WebExtensionsProjects\SPToolBox>Code .
```

### Build

Continous build run:

```node
npx webpack
```

### Install / Testing

Open FireFox and open URL "about:debugging".

Add temporal extension: browse to code folder and select manifest.json

### Other

Not necessary if cloned from GIT (or have package.json), just for documentation purpose

```node
npm init
npm install webpack --save-dev
npm install webpack-cli --save-dev
npm install --save react react-dom @types/react @types/react-dom
npm install --save-dev typescript awesome-typescript-loader source-map-loader
npm install --save-dev react-modal @types/react-modal
npm install --save-dev web-ext-types
npm install --save-dev jszip @types/jszip
npm install --save react-virtualized @types/react-virtualized
```

### TODO

- En listas corregir [object Object] que aparece en columnas tipo "editar"
- Comprobar funcionamiento general en entornos de producción
- Incorporar sistema de ayuda.
- Incorporar mejoras en el sistema de análisis en copiar-pegar, como algunas restricciones en los datos numéricos, fechas, etc.
- Incorporar filtro por columna
- Incorporar ajuste dinámico anchura columna

## Utilidades

### Debug

- Visualizar storage:
  - En about:debugging entrar a "Debug".
  - En consola introducir el comando:

    ```javascript
    chrome.storage.local.get(null, function(items) {
        console.log(items);
    });
    ```

## Errores

1. Error de columna en biblioteca imagenes. Posiblemente sea por el Expand

Volcado:

<SharePoint Toolbox>[12:26:49.841 ]: https://xxxxx.sharepoint.com/sites/1-it-pruch/_api/Web/Lists(guid'67987483-0e2d-436f-a469-bf151e57fa22')/Items?$Select=Id,EncodedAbsUrl,FileRef,FileLeafRef,Created,Modified,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Folder/Name,Folder/UniqueId,Folder/ItemCount,Folder/ServerRelativeUrl,Folder/ParentFolder/UniqueId,File/Length,DocIcon,LinkFilenameNoMenu,ImageSize,FileSizeDisplay,Created_x0020_Date/TimeCreated,RequiredField,ImageWidth,ImageHeight,NameOrTitle,PreviewOnForm&$Expand=Author,Editor,Folder,Folder/ParentFolder,File,Created_x0020_Date&$Top=500 spt.logax.ts:16:20
<SharePoint Toolbox>[12:26:49.841 ]: RESTQuery - Status==400: {"odata.error":{"code":"-1, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"The field or property 'Created_x0020_Date' does not exist."}}}

2. Error carga Style library. Hay una coma al final del Select antes del Expand

Volcado:

<SharePoint Toolbox>[12:35:04.081 ]: Query Items:https://xxxxx.sharepoint.com/sites/1-it-pruch/_api/Web/Lists(guid'8361eba9-474f-4a3a-b013-37b14e371dfe')/Items?$Select=Id,EncodedAbsUrl,FileRef,FileLeafRef,Created,Modified,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Folder/Name,Folder/UniqueId,Folder/ItemCount,Folder/ServerRelativeUrl,Folder/ParentFolder/UniqueId,File/Length,DocIcon,LinkFilename,CheckoutUser/Title,CheckoutUser/EMail,&$Expand=Author,Editor,Folder,Folder/ParentFolder,File,CheckoutUser&$Top=500 spt.logax.ts:16:20
<SharePoint Toolbox>[12:35:04.229 ]: RESTQuery - Status==400: {"odata.error":{"code":"-1, Microsoft.SharePoint.Client.InvalidClientQueryException","message":{"lang":"en-US","value":"The expression \"Id,EncodedAbsUrl,FileRef,FileLeafRef,Created,Modified,Author/Title,Author/EMail,Editor/Title,Editor/EMail,Folder/Name,Folder/UniqueId,Folder/ItemCount,Folder/ServerRelativeUrl,Folder/ParentFolder/UniqueId,File/Length,DocIcon,LinkFilename,CheckoutUser/Title,CheckoutUser/EMail,\" is not valid."}}}
