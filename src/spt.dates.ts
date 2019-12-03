export class Dates {
    /**
     * Get current time in format "00:00:00.000" used for log
     */
    public static getTimestampPrefix(): string {
        let d = new Date();

        let hours: string = ("0" + d.getHours()).slice(-2);
        let minutes: string = ("0" + d.getMinutes()).slice(-2);
        let seconds: string = ("0" + d.getSeconds()).slice(-2);
        let miliseconds: string = ("00" + d.getMilliseconds()).slice(-3);

        return hours + ":" + minutes + ":" + seconds + "." + miliseconds + " ";
    }

    /**
     * Get current time in format "yyyy-mm-dd-hh-mm" used for filenames
     */
    public static getFileSuffix(): string {
        let d = new Date();

        let month: string = ("0" + d.getMonth()).slice(-2);
        let day: string = ("0" + d.getDate()).slice(-2);
        let hours: string = ("0" + d.getHours()).slice(-2);
        let minutes: string = ("0" + d.getMinutes()).slice(-2);

        return d.getFullYear() + "-" + month + "-" + day + "-" + hours + "-" + minutes;
    }
}