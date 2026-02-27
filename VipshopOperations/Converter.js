class Converter {
  static _instance = null;

  constructor() {
    if (Converter._instance) {
      return Converter._instance;
    }

    Converter._instance = this;
  }

  // 单例模式获取转化器对象
  static getInstance() {
    if (!Converter._instance) {
      Converter._instance = new Converter();
    }
    return Converter._instance;
  }

  // 字符串——>标准Date对象
  parseDate(dateStr) {
    if (!dateStr) return null;

    const timestamp = Date.parse(String(dateStr));
    if (isNaN(timestamp)) return null;

    return new Date(timestamp);
  }

  // 转换为number
  toNumber(value) {
    if (
      value == null ||
      String(value).trim() === "" ||
      typeof value === "boolean"
    ) {
      return undefined;
    }

    const num = Number(value);
    return isFinite(num) ? num : undefined;
  }

  // 转换为string
  toString(value) {
    if (
      value == null ||
      String(value).trim() === "" ||
      typeof value === "boolean"
    ) {
      return undefined;
    }

    return String(value);
  }

  // 字符串——>去前导符日期字符串YYYY-MM-DD
  toDateStr(value) {
    if (
      value == null ||
      String(value).trim() === "" ||
      typeof value === "boolean"
    ) {
      return undefined;
    }

    const cleanValue = String(value).replace(/^'/, "");

    const date = this.parseDate(cleanValue);
    if (!date) return undefined;

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  // 字符串——>前导符日期字符串'YYYY-MM-DD
  formatDate(value) {
    const date = this.parseDate(value);
    if (!date) return undefined;

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");

    return `'${year}-${month}-${day}`;
  }
}
