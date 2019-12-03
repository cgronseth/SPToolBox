/*  Simple cache mechanism that stores data in memory for easy fast fetches.

    Use:
        Store data with "Put". It requires a unique key and accepts as value any object. Give it a maximum lifetime in miliseconds.
        Check and retrieve with "Has" and "Get" respectively.

        Example:

        let name: string;
        if (Has("formUserName")) {
            name = Cache.Get<string>("formUserName");
        } else {
            name = GetNameFromSlowDataSource();
            Cache.Put<string>("formUserName", name, 60000);
        }

    Pros:
        - Easy to use and quite fast.
        - Can be extended to save to persistant storage (session html5 storage for example), or add sliding expiration, etc.

    Cons:
        - Data is not automatically purged when expired, only if a "Get" is executed it will remove from memory.
*/

export class Cache {
    private static data: Map<string, any> = new Map<string, any>();
    private static dataTimeout: Map<string, Date> = new Map<string, Date>();

    // Tiempo 5 minutos definido para PutShort
    private static readonly shortTime: number = 300000;

    // Tiempo 1 hora definido para PutLong
    private static readonly longTime: number = 3600000;

    public static Put<T>(key: string, data: T, ms: number) {
        this.data.set(key, data);
        this.dataTimeout.set(key, new Date(Date.now() + ms));
    }

    public static PutShort<T>(key: string, data: T) {
        this.Put(key, data, Cache.shortTime);
    }

    public static PutLong<T>(key: string, data: T) {
        this.Put(key, data, Cache.longTime);
    }

    public static Get<T>(key: string): T {
        if (this.data.has(key)) {
            return this.data.get(key);
        }
        return null;
    }

    public static Has(key: string): boolean {
        if (this.dataTimeout.has(key)) {
            let checkTimeout: Date = this.dataTimeout.get(key);
            if (checkTimeout && checkTimeout.getTime() > Date.now()) {
                return true;
            }
            this.dataTimeout.delete(key);
            this.data.delete(key);
        }
        return false;
    }
}  
