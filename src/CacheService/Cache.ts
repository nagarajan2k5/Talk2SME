import NodeCache = require("node-cache");


class Cache {
    dataCache = new NodeCache();
    constructor() {
        this.dataCache = new NodeCache({ stdTTL: 100, checkperiod: 120 });
    }

    private SetData(key, obj) {
        let success = this.dataCache.set(key, obj);
        return success;
    }

    private GetData(key) {
        let data = this.dataCache.get(key);
        return data;
    }

    private RemoveData(key) {
        let count = this.dataCache.del(key);
        return count;
    }
}


