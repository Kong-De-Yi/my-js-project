/**
 * 验证引擎 - 负责数据的验证和校验
 *
 * @class ValidationEngine
 * @description 作为系统的验证核心，提供以下功能：
 * - 内置丰富的验证规则（必填、枚举、正则、范围、数字、日期等）
 * - 支持字段级的多规则验证
 * - 支持联合主键的唯一性验证
 * - 支持自定义验证规则扩展
 * - 提供详细的验证错误信息格式化
 *
 * 内置验证规则：
 * - required: 必填项验证
 * - enum: 枚举值验证
 * - pattern: 正则表达式验证
 * - range: 数值范围验证
 * - nonNegative: 非负数验证
 * - positive: 正数验证
 * - number: 数字格式验证
 * - date: 日期格式验证
 * - year: 年份范围验证
 * - month: 月份验证（1-12）
 * - week: 周数验证（1-53）
 *
 * 该类采用单例模式，确保全局只有一个验证引擎实例。
 *
 * @example
 * // 获取验证引擎实例
 * const validator = ValidationEngine.getInstance();
 *
 * // 验证单个字段
 * const result = validator.validateField("123", {
 *   validators: [{ type: "number" }, { type: "range", params: { min: 0, max: 100 } }]
 * }, "价格");
 *
 * // 验证整个实体
 * const entityResult = validator.validateEntity(product, productConfig, { allData: products });
 *
 * // 验证实体集合并格式化错误
 * const results = validator.validateAll(products, productConfig);
 * const errorMsg = validator.formatErrors(results, "商品");
 */
class ValidationEngine {
  /** @type {ValidationEngine} 单例实例 */
  static _instance = null;

  /**
   * 创建验证引擎实例
   * @private
   */
  constructor() {
    if (ValidationEngine._instance) {
      return ValidationEngine._instance;
    }

    this._config = DataConfig.getInstance();

    // 默认验证规则
    this._validators = {
      required: (value, params) => ({
        valid: value != undefined && String(value).trim() !== "",
        message: params?.message || "不能为空",
      }),

      enum: (value, params) => ({
        valid:
          value == undefined ||
          String(value).trim() === "" ||
          params.values.includes(value),
        message:
          params?.message || `必须是以下值之一：${params.values.join(", ")}`,
      }),

      pattern: (value, params) => ({
        valid:
          value == undefined ||
          String(value).trim() === "" ||
          params.regex.test(String(value)),
        message:
          params?.message ||
          `格式不正确${params.description ? ":" + params.description : ""}`,
      }),

      range: (value, params) => {
        if (value == undefined || String(value).trim() === "")
          return { valid: true };
        if (typeof value === "boolean")
          return { valid: false, message: "必须是数字" };

        const num = Number(value);
        if (!isFinite(num)) return { valid: false, message: "必须是数字" };

        if (params.min != undefined && num < params.min) {
          return { valid: false, message: `不能小于${params.min}` };
        }
        if (params.max != undefined && num > params.max) {
          return { valid: false, message: `不能大于${params.max}` };
        }
        return { valid: true };
      },

      nonNegative: (value, params) => {
        if (value == undefined || String(value).trim() === "")
          return { valid: true };
        if (typeof value === "boolean")
          return { valid: false, message: "必须是数字" };

        const num = Number(value);
        if (!isFinite(num)) return { valid: false, message: "必须是数字" };
        return { valid: num >= 0, message: params?.message || "不能为负数" };
      },

      positive: (value, params) => {
        if (value == undefined || String(value).trim() === "")
          return { valid: true };
        if (typeof value === "boolean")
          return { valid: false, message: "必须是数字" };

        const num = Number(value);
        if (!isFinite(num)) return { valid: false, message: "必须是数字" };
        return { valid: num > 0, message: params?.message || "必须大于0" };
      },

      number: (value, params) => {
        if (value == undefined || String(value).trim() === "")
          return { valid: true };
        if (typeof value === "boolean")
          return { valid: false, message: "必须是数字" };

        const num = Number(value);
        return {
          valid: isFinite(num),
          message: params?.message || "必须是有效的数字",
        };
      },

      date: (value, params) => {
        if (value == undefined || String(value).trim() === "") {
          return { valid: true };
        }

        const cleanValue = String(value).replace(/^'/, "");
        const timestamp = Date.parse(cleanValue);
        const isValidDate = !isNaN(timestamp);

        return {
          valid: isValidDate,
          message: params?.message || "不是有效的日期格式",
        };
      },

      // 年份验证
      year: (value, params) => {
        if (value == undefined || String(value).trim() === "")
          return { valid: true };
        if (typeof value === "boolean")
          return { valid: false, message: "必须是数字" };

        const num = Number(value);
        if (!isFinite(num)) return { valid: false, message: "必须是数字" };
        const currentYear = new Date().getFullYear();
        const minYear = params?.minYear || currentYear - 10;
        const maxYear = params?.maxYear || currentYear + 5;
        return {
          valid: num >= minYear && num <= maxYear,
          message: `年份必须在${minYear}-${maxYear}之间`,
        };
      },

      // 月份验证
      month: (value, params) => {
        if (value == undefined || String(value).trim() === "")
          return { valid: true };
        if (typeof value === "boolean")
          return { valid: false, message: "必须是数字" };

        const num = Number(value);
        if (!isFinite(num)) return { valid: false, message: "必须是数字" };
        return {
          valid: num >= 1 && num <= 12,
          message: "月份必须在1-12之间",
        };
      },

      // 周数验证
      week: (value, params) => {
        if (value == undefined || String(value).trim() === "")
          return { valid: true };
        if (typeof value === "boolean")
          return { valid: false, message: "必须是数字" };

        const num = Number(value);
        if (!isFinite(num)) return { valid: false, message: "必须是数字" };
        return {
          valid: num >= 1 && num <= 53,
          message: "周数必须在1-53之间",
        };
      },
    };

    ValidationEngine._instance = this;
  }

  /**
   * 获取验证引擎的单例实例
   * @static
   * @returns {ValidationEngine} 验证引擎实例
   */
  static getInstance() {
    if (!ValidationEngine._instance) {
      ValidationEngine._instance = new ValidationEngine();
    }
    return ValidationEngine._instance;
  }

  /**
   * 以字符串形式返回字段的组合键值
   * @private
   * @param {Object} item - 数据项
   * @param {string[]} fields - 字段名数组
   * @returns {string} 用'¦'连接的复合键值，空字段转为空字符串
   * @description
   * 用于生成联合主键的字符串表示，便于比较和去重。
   * 例如：fields = ["itemNumber", "salesDate"] 时，
   * 可能返回 "A001¦2024-01-15"
   */
  _getCompositeKeyValue(item, fields) {
    if (fields.length === 1) {
      const value = item[fields[0]];
      return value != undefined ? String(value) : "";
    }

    return fields
      .map((f) => {
        const value = item[f];
        return value != undefined ? String(value) : "";
      })
      .join("¦");
  }

  /**
   * 获取字段的标题（用于错误信息）
   * @private
   * @param {string} fieldName - 字段名
   * @param {Object} fieldConfig - 字段配置
   * @returns {string} 字段标题，如果没有配置则返回字段名
   */
  _getFieldTitle(fieldName, fieldConfig) {
    return fieldConfig?.title || fieldName;
  }

  /**
   * 验证实体的联合主键完整性和唯一性
   * @private
   * @param {Object} entity - 实体对象
   * @param {Object} entityConfig - 实体配置
   * @param {Object[]} allData - 所有实体数据（用于唯一性验证）
   * @returns {Object} 验证结果
   * @returns {boolean} return.valid - 是否通过验证
   * @returns {string[]} return.errors - 错误信息数组
   * @description
   * 验证内容：
   * 1. 主键字段是否为空（完整性验证）
   * 2. 主键值是否在数据集中已存在（唯一性验证）
   *
   * 注意：验证时会排除自身（item === entity）以避免自冲突。
   */
  _validateCompositeKey(entity, entityConfig, allData) {
    if (!entityConfig.uniqueKey) {
      return { valid: true, errors: [] };
    }

    const uniqueKeyConfig = this._config.parseUniqueKey(entityConfig.uniqueKey);
    const fields = uniqueKeyConfig.fields; // 获取业务实体的主键数组

    if (fields.length === 0) {
      return { valid: true, errors: [] };
    }

    // 验证实体对象主键字段是否为空
    const missingFields = [];
    fields.forEach((field) => {
      const value = entity[field];
      if (value == undefined || String(value).trim() === "") {
        missingFields.push(field);
      }
    });

    if (missingFields.length > 0) {
      const missingTitles = missingFields
        .map((f) => this._getFieldTitle(f, entityConfig.fields[f]))
        .join("、");
      return {
        valid: false,
        errors: [`联合主键字段【${missingTitles}】不能为空`],
      };
    }

    const currentKey = this._getCompositeKeyValue(entity, fields);

    const duplicates = allData.filter((item) => {
      // 排除实体对象本身
      if (item === entity) return false;

      return this._getCompositeKeyValue(item, fields) === currentKey;
    });

    if (duplicates.length > 0) {
      const fieldNames = fields
        .map((f) => this._getFieldTitle(f, entityConfig.fields[f]))
        .join("、");

      const valueParts = fields.map((f) => {
        const value = entity[f];
        const title = this._getFieldTitle(f, entityConfig.fields[f]);
        return `${title}:${value}`;
      });

      const errorMessage =
        uniqueKeyConfig.message ||
        `【${fieldNames}】的组合值(${valueParts.join(" ")})已存在`;

      return {
        valid: false,
        errors: [errorMessage],
      };
    }

    return { valid: true, errors: [] };
  }

  /**
   * 注册自定义验证规则
   * @param {string} name - 规则名称
   * @param {Function} validatorFn - 验证函数
   * @returns {void}
   * @description
   * 验证函数签名：(value, params) => { valid: boolean, message: string }
   * - value: 要验证的值
   * - params: 验证参数（来自字段配置）
   * - 返回：包含 valid 和 message 的对象
   *
   * @example
   * validator.register("customRule", (value, params) => {
   *   if (value === "特殊值") {
   *     return { valid: false, message: "不能是特殊值" };
   *   }
   *   return { valid: true };
   * });
   */
  register(name, validatorFn) {
    this._validators[name] = validatorFn;
  }

  /**
   * 验证单个值
   * @param {*} value - 要验证的值
   * @param {Object} validatorConfig - 验证器配置
   * @param {string} validatorConfig.type - 验证规则类型
   * @param {Object} validatorConfig.params - 验证参数
   * @param {string} [fieldTitle] - 字段标题（用于错误信息）
   * @returns {Object} 验证结果
   * @returns {boolean} return.valid - 是否通过验证
   * @returns {string} return.message - 错误信息（验证通过时为空）
   */
  validateValue(value, validatorConfig, fieldTitle = "") {
    const validator = this._validators[validatorConfig.type];
    if (!validator) {
      return { valid: true };
    }

    const result = validator(value, validatorConfig.params);
    return {
      valid: result.valid,
      message: result.valid ? "" : `【${fieldTitle}】${result.message}`,
    };
  }

  /**
   * 验证单个字段的所有验证规则
   * @param {*} value - 字段值
   * @param {Object} fieldConfig - 字段配置
   * @param {string} fieldName - 字段名
   * @returns {Object} 验证结果
   * @returns {boolean} return.valid - 是否通过验证
   * @returns {string[]} return.errors - 错误信息数组
   * @description
   * 验证逻辑：
   * 1. 如果字段没有验证规则，直接通过
   * 2. 遍历字段配置的所有验证规则
   * 3. 对于非 required 规则，如果值为空则跳过验证
   * 4. 收集所有验证失败的错误信息
   */
  validateField(value, fieldConfig, fieldName) {
    if (!fieldConfig.validators || fieldConfig.validators.length === 0) {
      return { valid: true, errors: [] };
    }

    const errors = [];
    const fieldTitle = fieldConfig.title || fieldName;

    // 遍历验证字段的所有验证规则
    for (const validator of fieldConfig.validators) {
      // 非必须字段的空值跳过验证
      if (
        validator.type !== "required" &&
        (value == undefined || String(value).trim() === "")
      ) {
        continue;
      }

      const result = this.validateValue(value, validator, fieldTitle);
      if (!result.valid) {
        errors.push(result.message);
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }

  /**
   * 验证单个实体对象
   * @param {Object} entity - 实体对象
   * @param {Object} entityConfig - 实体配置
   * @param {Object} [context] - 验证上下文
   * @param {Object[]} [context.allData] - 所有实体数据（用于唯一性验证）
   * @returns {Object} 验证结果
   * @returns {boolean} return.valid - 实体是否有效
   * @returns {Object} return.errors - 错误信息映射
   * @returns {string[]} return.errors[fieldName] - 字段的错误信息数组
   * @returns {string[]} return.errors._composite - 联合主键的错误信息
   * @returns {number} return.rowNumber - 实体在Excel中的行号
   * @description
   * 验证流程：
   * 1. 遍历实体配置中的所有字段（跳过计算字段）
   * 2. 对每个字段调用 validateField 进行验证
   * 3. 如果提供了 allData，调用 _validateCompositeKey 验证主键唯一性
   * 4. 收集所有验证错误
   */
  validateEntity(entity, entityConfig, context = {}) {
    const errors = {};

    // 验证各个字段
    Object.entries(entityConfig.fields).forEach(([fieldName, fieldConfig]) => {
      // 跳过计算字段的验证
      if (fieldConfig.type === "computed") return;

      // 验证单个字段的所有验证规则
      const value = entity[fieldName];
      const result = this.validateField(value, fieldConfig, fieldName);

      if (!result.valid) {
        errors[fieldName] = result.errors;
      }
    });

    // 验证主键唯一性
    if (context.allData) {
      const compositeResult = this._validateCompositeKey(
        entity,
        entityConfig,
        context.allData,
      );
      if (!compositeResult.valid) {
        errors._composite = compositeResult.errors;
      }
    }

    return {
      valid: Object.keys(errors).length === 0,
      errors,
      rowNumber: entity._rowNumber,
    };
  }

  /**
   * 验证整个实体集
   * @param {Object[]} entities - 实体对象数组
   * @param {Object} entityConfig - 实体配置
   * @returns {Object} 验证结果
   * @returns {boolean} return.valid - 是否所有实体都有效
   * @returns {Array} return.items - 每个实体的验证结果
   * @returns {Object} return.summary - 验证统计
   * @returns {number} return.summary.total - 总实体数
   * @returns {number} return.summary.valid - 有效实体数
   * @returns {number} return.summary.invalid - 无效实体数
   * @description
   * 验证流程：
   * 1. 遍历所有实体
   * 2. 对每个实体调用 validateEntity（传入整个数据集用于唯一性验证）
   * 3. 收集统计信息
   */
  validateAll(entities, entityConfig) {
    const results = {
      valid: true,
      items: [],
      summary: {
        total: entities.length,
        valid: 0,
        invalid: 0,
      },
    };

    entities.forEach((entity, index) => {
      const result = this.validateEntity(entity, entityConfig, {
        allData: entities,
      });

      results.items.push({
        ...result,
        data: entity,
        index,
      });

      if (result.valid) {
        results.summary.valid++;
      } else {
        results.summary.invalid++;
        results.valid = false;
      }
    });

    return results;
  }

  /**
   * 格式化实体集的验证结果为可读的错误信息
   * @param {Object} validationResult - validateAll 的返回结果
   * @param {string} entityName - 实体名称（用于错误信息标题）
   * @returns {string|null} 格式化的错误信息，验证通过时返回null
   * @description
   * 格式化格式：
   * 【实体名称】数据验证失败：
   *
   * 第X行：
   *   【字段标题】错误信息1
   *   【字段标题】错误信息2
   *
   * 第Y行：
   *   ...
   *
   * @example
   * 【商品】数据验证失败：
   *
   * 第5行：
   *   【货号】不能为空
   *   【价格】必须大于0
   *
   * 第8行：
   *   【货号】格式不正确
   */
  formatErrors(validationResult, entityName) {
    if (validationResult.valid) {
      return null;
    }

    let message = `【${entityName}】数据验证失败：\n`;

    validationResult.items.forEach((item) => {
      if (!item.valid) {
        message += `\n第${item.rowNumber || "?"}行：\n`;

        Object.values(item.errors).forEach((fieldErrors) => {
          fieldErrors.forEach((err) => {
            message += `  ${err}\n`;
          });
        });
      }
    });

    return message;
  }
}
