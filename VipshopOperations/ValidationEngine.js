// ============================================================================
// ValidationEngine.js - 数据验证引擎（增强版）
// 功能：支持计算字段的验证
// ============================================================================

class ValidationEngine {
  constructor() {
    this._validators = {
      required: (value, params) => ({
        valid: value !== undefined && value !== null && value !== "",
        message: params?.message || "不能为空",
      }),

      enum: (value, params) => ({
        valid: value === undefined || params.values.includes(value),
        message:
          params?.message || `必须是以下值之一：${params.values.join(", ")}`,
      }),

      pattern: (value, params) => ({
        valid: value === undefined || params.regex.test(String(value)),
        message:
          params?.message ||
          `格式不正确${params.description ? "：" + params.description : ""}`,
      }),

      range: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        if (isNaN(num)) return { valid: false, message: "必须是数字" };
        if (params.min !== undefined && num < params.min) {
          return { valid: false, message: `不能小于${params.min}` };
        }
        if (params.max !== undefined && num > params.max) {
          return { valid: false, message: `不能大于${params.max}` };
        }
        return { valid: true };
      },

      nonNegative: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        if (isNaN(num)) return { valid: false, message: "必须是数字" };
        return { valid: num >= 0, message: params?.message || "不能为负数" };
      },

      positive: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        if (isNaN(num)) return { valid: false, message: "必须是数字" };
        return { valid: num > 0, message: params?.message || "必须大于0" };
      },

      number: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        return {
          valid: !isNaN(num) && isFinite(num),
          message: params?.message || "必须是有效的数字",
        };
      },

      date: (value, params) => {
        if (value === undefined || value === null || value === "") {
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

      // 新增：年份验证
      year: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        if (isNaN(num)) return { valid: false, message: "必须是数字" };
        const currentYear = new Date().getFullYear();
        const minYear = params?.minYear || currentYear - 10;
        const maxYear = params?.maxYear || currentYear + 5;
        return {
          valid: num >= minYear && num <= maxYear,
          message: `年份必须在${minYear}-${maxYear}之间`,
        };
      },

      // 新增：月份验证
      month: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        if (isNaN(num)) return { valid: false, message: "必须是数字" };
        return {
          valid: num >= 1 && num <= 12,
          message: "月份必须在1-12之间",
        };
      },

      // 新增：周数验证
      week: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        if (isNaN(num)) return { valid: false, message: "必须是数字" };
        return {
          valid: num >= 1 && num <= 53,
          message: "周数必须在1-53之间",
        };
      },
    };
  }

  _getCompositeKeyValue(item, fields) {
    if (fields.length === 1) {
      const value = item[fields[0]];
      return value !== undefined && value !== null ? String(value) : "";
    }

    return fields
      .map((f) => {
        const value = item[f];
        return value !== undefined && value !== null ? String(value) : "";
      })
      .join("¦");
  }

  _getFieldTitle(fieldName, fieldConfig) {
    return fieldConfig?.title || fieldName;
  }

  _validateCompositeKey(entity, entityConfig, allData) {
    if (!entityConfig.uniqueKey) {
      return { valid: true, errors: [] };
    }

    const uniqueKeyConfig = dataConfig.parseUniqueKey(entityConfig.uniqueKey);
    const fields = uniqueKeyConfig.fields;

    if (fields.length === 0) {
      return { valid: true, errors: [] };
    }

    const missingFields = [];
    fields.forEach((field) => {
      const value = entity[field];
      if (value === undefined || value === null || value === "") {
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
      if (item === entity) return false;
      const itemKey = this._getCompositeKeyValue(item, fields);
      return itemKey === currentKey && itemKey !== "";
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

  register(name, validatorFn) {
    this._validators[name] = validatorFn;
  }

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

  validateField(value, fieldConfig, fieldName) {
    if (!fieldConfig.validators || fieldConfig.validators.length === 0) {
      return { valid: true, errors: [] };
    }

    const errors = [];
    const fieldTitle = fieldConfig.title || fieldName;

    for (const validator of fieldConfig.validators) {
      if (
        validator.type !== "required" &&
        (value === undefined || value === null || value === "")
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

  validateEntity(entity, entityConfig, context = {}) {
    const errors = {};

    // 验证各个字段（包括计算字段？通常不验证计算字段）
    Object.entries(entityConfig.fields).forEach(([fieldName, fieldConfig]) => {
      // 跳过计算字段的验证（因为它们是由系统自动生成的）
      if (fieldConfig.type === "computed") return;

      const value = entity[fieldName];
      const result = this.validateField(value, fieldConfig, fieldName);

      if (!result.valid) {
        errors[fieldName] = result.errors;
      }
    });

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

  formatErrors(validationResult, entityName) {
    if (validationResult.valid) {
      return null;
    }

    let message = `【${entityName}】数据验证失败：\n`;

    validationResult.items.forEach((item) => {
      if (!item.valid) {
        message += `\n第${item.rowNumber || "?"}行：\n`;

        Object.entries(item.errors).forEach(([field, fieldErrors]) => {
          if (field === "_composite") {
            fieldErrors.forEach((err) => {
              message += `  ${err}\n`;
            });
          } else {
            fieldErrors.forEach((err) => {
              message += `  ${err}\n`;
            });
          }
        });
      }
    });

    return message;
  }

  parseDate(dateStr) {
    if (!dateStr) return null;

    const cleanValue = String(dateStr).replace(/^'/, "");
    const timestamp = Date.parse(cleanValue);
    if (isNaN(timestamp)) return null;

    return new Date(timestamp);
  }

  formatDate(date) {
    if (!date) return "";

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  /**
   * 获取年份
   * @param {Date} date - 日期
   * @returns {number} 年份
   */
  getYear(date) {
    return date ? date.getFullYear() : new Date().getFullYear();
  }

  /**
   * 获取月份
   * @param {Date} date - 日期
   * @returns {number} 月份（1-12）
   */
  getMonth(date) {
    return date ? date.getMonth() + 1 : new Date().getMonth() + 1;
  }

  /**
   * 获取ISO周数
   * @param {Date} date - 日期
   * @returns {number} 周数（1-53）
   */
  getISOWeekNumber(date) {
    if (!date) date = new Date();
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
    const week1 = new Date(d.getFullYear(), 0, 4);
    return (
      1 +
      Math.round(((d - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7)
    );
  }
}

const validationEngine = new ValidationEngine();
