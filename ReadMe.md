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

n/a
