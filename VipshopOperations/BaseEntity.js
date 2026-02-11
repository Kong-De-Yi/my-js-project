class BaseEntity {
  constructor() {
    this._validationErrors = new Map();
    this._dynamicProperties = new Map();
    this._originalValues = new Map();
  }

  //动态属性处理
  setDynamicProperty(key, value) {
    this._dynamicProperties.set(key, value);
    //为保持输出顺序，使用Symbol作为属性名
    const sym = Symbol.for(key);
    this[sym] = value;
  }

  getDynamicProperty(key) {
    return this._dynamicProperties.get(key);
  }

  //验证错误处理
  addValidationError(property, message) {
    if (!this._validationErrors.has(property)) {
      this._validationErrors.set(property, []);
    }
    this._validationErrors.get(property).push(message);
  }

  get validationErrors() {
    const errors = {};
    for (const [property, message] of this._validationErrors) {
      errors[property] = message.join(";");
    }
    return errors;
  }

  get hasError() {
    return this._validationErrors.size > 0;
  }

  //输出到工作表时的属性顺序控制
  getOutputProperties(config) {
    const orderedProperties = [];

    //1.固定属性按配置顺序
    if (config.fixedPropertyOrder) {
      orderedProperties.push(...config.fixedPropertyOrder);
    }

    //2.动态属性按时间顺序
    if (this._dynamicProperties.size > 0) {
      const dynamicProps = Array.from(this._dynamicProperties.keys()).sort(
        (a, b) => {
          const dateA = this.extractDateFromPropertyName(a);
          const dateB = this.extractDateFromPropertyName(b);
          return dateA - dateB;
        },
      );
      orderedProperties.push(...dynamicProps);
    }

    //3.其他属性
    const allProperties = Object.keys(this).filter(
      (prop) =>
        !prop.startsWith("_") &&
        prop !== "validationErrors" &&
        prop !== "hasError",
    );

    const remainingProps = allProperties.filter(
      (prop) => !orderedProperties.includes(prop),
    );

    orderedProperties.push(...remainingProps);

    return orderedProperties;
  }

  extractDateFromPropertyName(propName) {
    //从属性名中提取日期信息用于排序
    const dateMatch = propName.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (dateMatch) {
      return new Date(dateMatch[1], dateMatch[2] - 1, dateMatch[3]).getTime();
    }
    return 0;
  }
}
