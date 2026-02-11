class CacheManager {
  constructor() {
    this.cache = new Map();
    this.ttl = 5 * 60 * 1000; //5分钟
  }

  set(key, value) {
    this.cache.set(key, {
      value,
      timestamp: Date.now(),
    });
  }

  get(key) {
    const item = this.cache.get(key);
    if (!item) return null;

    if (Date.now() - item.timestamp > this.ttl) {
      this.cache.delete(key);
      return null;
    }

    return item.value;
  }

  has(key) {
    return this.get(key) !== null;
  }

  clear() {
    this.cache.clear();
  }

  //智能缓存：根据数据量自动调整缓存策略
  setSmart(key, value) {
    const size = this.estimateSize(value);
    if (size > 10 * 1024 * 1024) {
      //10MB
      this.ttl = 1 * 60 * 1000; //大型数据缓存1分钟
    }
    this.set(key, value);
  }

  estimateSize(obj) {
    return JSON.stringify(obj).length;
  }
}
