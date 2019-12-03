# SPToolBox

SharePoint Tools for managing Lists and other stuff.

## Project setup

### IDE Setup

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

### Build

```node
npx webpack
```

### TODO

#### Importante

- Comprobar funcionamiento general
- Crear guía para extensión y preparar subida nueva versión

#### Sin prisas

- Incorporar sistema de ayuda.
- Incorporar mejoras en el sistema de análisis en copiar-pegar, como algunas restricciones en los datos numéricos, fechas, etc.

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
