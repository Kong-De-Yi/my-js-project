/**
 * 数据转换器 - 负责数据类型的转换和格式化
 *
 * @class Converter
 * @description 作为系统的数据转换核心，提供以下功能：
 * - 字符串与Date对象之间的转换
 * - 数值类型的安全转换（处理空值、布尔值、非法数值）
 * - 日期格式的统一化处理（YYYY-MM-DD）
 * - Excel日期格式的特殊处理（前导符处理）
 *
 * Excel日期处理说明：
 * - Excel中的日期可能带有前导单引号（如 "'2024-01-15"）
 * - 读取时：需要去除前导符后解析
 * - 写入时：需要添加前导符以保持格式
 *
 * 该类采用单例模式，确保全局只有一个转换器实例。
 *
 * @example
 * // 获取转换器实例
 * const converter = Converter.getInstance();
 *
 * // 转换为数字
 * const num = converter.toNumber("123.45"); // 返回 123.45
 *
 * // 转换为日期字符串
 * const dateStr = converter.toDateStr("2024-01-15"); // 返回 "2024-01-15"
 *
 * // 格式化日期（带前导符）
 * const formatted = converter.formatDate(new Date()); // 返回 "'2024-01-15"
 */
class Converter {
  /** @type {Converter} 单例实例 */
  static _instance = null;

  /**
   * 创建转换器实例
   * @private
   */
  constructor() {
    if (Converter._instance) {
      return Converter._instance;
    }

    Converter._instance = this;
  }

  /**
   * 获取转换器的单例实例
   * @static
   * @returns {Converter} 转换器实例
   */
  static getInstance() {
    if (!Converter._instance) {
      Converter._instance = new Converter();
    }
    return Converter._instance;
  }

  /**
   * 将字符串解析为标准的Date对象
   * @param {string|Date} dateStr - 日期字符串或Date对象
   * @returns {Date|null} 解析后的Date对象，解析失败返回null
   * @description
   * 支持的输入格式：
   * - ISO格式： "2024-01-15"
   * - 常见格式： "2024/01/15", "01/15/2024"
   * - Excel序列号： 44975（从1900-01-01开始的天数）
   * - Date对象： 直接返回原对象
   *
   * 注意：使用 Date.parse() 进行解析，不同浏览器对日期格式的支持可能有差异。
   */
  parseDate(dateStr) {
    if (!dateStr) return null;

    const timestamp = Date.parse(String(dateStr));
    if (isNaN(timestamp)) return null;

    return new Date(timestamp);
  }

  /**
   * 将值转换为数字类型
   * @param {*} value - 要转换的值
   * @returns {number|undefined} 转换后的数字，无效值返回undefined
   * @description
   * 转换规则：
   * 1. 空值（null/undefined） => undefined
   * 2. 空字符串或纯空格字符串 => undefined
   * 3. 布尔值 => undefined
   * 4. 其他值尝试转换为数字
   * 5. 转换结果为 Infinity 或 NaN 时返回 undefined
   *
   * @example
   * toNumber("123")      // 返回 123
   * toNumber("123.45")   // 返回 123.45
   * toNumber("")         // 返回 undefined
   * toNumber(true)       // 返回 undefined
   * toNumber("abc")      // 返回 undefined
   */
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

  /**
   * 将值转换为字符串类型
   * @param {*} value - 要转换的值
   * @returns {string|undefined} 转换后的字符串，无效值返回undefined
   * @description
   * 转换规则：
   * 1. 空值（null/undefined） => undefined
   * 2. 空字符串或纯空格字符串 => undefined
   * 3. 布尔值 => undefined
   * 4. 其他值调用 String() 转换
   *
   * @example
   * toString("hello")    // 返回 "hello"
   * toString(123)        // 返回 "123"
   * toString("")         // 返回 undefined
   * toString(true)       // 返回 undefined
   */
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

  /**
   * 将值转换为标准的日期字符串（YYYY-MM-DD）
   * @param {*} value - 要转换的值（日期字符串、Date对象等）
   * @returns {string|undefined} 格式化的日期字符串，无效值返回undefined
   * @description
   * 转换流程：
   * 1. 过滤无效值（空值、空字符串、布尔值）
   * 2. 去除可能存在的 Excel 前导单引号（如 "'2024-01-15" -> "2024-01-15"）
   * 3. 调用 parseDate 解析为 Date 对象
   * 4. 格式化为 YYYY-MM-DD 格式
   *
   * 此方法用于从 Excel 读取数据时的日期转换。
   *
   * @example
   * toDateStr("2024-01-15")        // 返回 "2024-01-15"
   * toDateStr("'2024-01-15")       // 返回 "2024-01-15"（去除前导符）
   * toDateStr(new Date(2024,0,15)) // 返回 "2024-01-15"
   * toDateStr("")                   // 返回 undefined
   */
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

  /**
   * 将日期格式化为带前导符的字符串（'YYYY-MM-DD）
   * @param {*} value - 要格式化的值（日期字符串、Date对象等）
   * @returns {string|undefined} 带前导符的日期字符串，无效值返回undefined
   * @description
   * 转换流程：
   * 1. 调用 parseDate 解析为 Date 对象
   * 2. 格式化为 YYYY-MM-DD 格式
   * 3. 在前面添加单引号
   *
   * Excel 中单引号的作用：
   * - 强制将单元格格式设为文本
   * - 防止日期被自动转换为 Excel 序列号
   * - 保持日期显示格式不变
   *
   * 此方法用于向 Excel 写入数据时的日期格式化。
   *
   * @example
   * formatDate("2024-01-15")        // 返回 "'2024-01-15"
   * formatDate(new Date(2024,0,15)) // 返回 "'2024-01-15"
   * formatDate("")                   // 返回 undefined
   */
  formatDate(value) {
    const date = this.parseDate(value);
    if (!date) return undefined;

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");

    return `'${year}-${month}-${day}`;
  }
}
