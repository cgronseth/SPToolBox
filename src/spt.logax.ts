import { Dates } from "./spt.dates";

const appName: string = "SharePoint Toolbox";

/**
 * Logging helper class. Extend to write to file, database, etc.
 */
export class LogAx {
    //Habilitar trazas informativas. TODO: pasar a configurable/automático
    static readonly TRACE: boolean = true;
    //Escribe traza verbose a consola y cualquier otro medio futuro
    public static trace(txt: string): void {
        if (LogAx.TRACE) {
            //let t = LogAx.groupTexts[appName];
            //LogAx.groupTexts[appName] = (!t) ? Dates.getTimestampPrefix() + txt : t + '\n' + Dates.getTimestampPrefix() + txt;
            console.log("<" + appName + ">" + "[" + Dates.getTimestampPrefix() + "]: " + txt);
        }
    }
    //private static groupTexts: any = {};
}
