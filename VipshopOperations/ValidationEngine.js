class ValidationEngine {
  static _instance = null;

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

  // 以字符串形式返回字段的组合键值，多字段用 ¦ 连接
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

  // 返回字段的标题
  _getFieldTitle(fieldName, fieldConfig) {
    return fieldConfig?.title || fieldName;
  }

  // 验证一个实体对象的主键完整性和唯一性,返回 {valid,errors:[]}
  _validateCompositeKey(entity, entityConfig, allData) {
    if (!entityConfig.uniqueKey) {
      return { valid: true, errors: [] };
    }

    const uniqueKeyConfig = this._config.parseUniqueKey(entityConfig.uniqueKey);
    const fields = uniqueKeyConfig.fields; // 获取业务实体的主键数组

    if (fields.length === 0) {
      return { valid: true, errors: [] };
    }

    // 验证实体对象主键的字段完整性
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

  // 单例模式
  static getInstance() {
    if (!ValidationEngine._instance) {
      ValidationEngine._instance = new ValidationEngine();
    }
    return ValidationEngine._instance;
  }

  // 注册自定义验证规则
  register(name, validatorFn) {
    this._validators[name] = validatorFn;
  }

  // 验证单个规则,返回{valid,message}
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

  // 验证单字段的所有验证规则,返回{valid,errors:[]}
  validateField(value, fieldConfig, fieldName) {
    if (!fieldConfig.validators || fieldConfig.validators.length === 0) {
      return { valid: true, errors: [] };
    }

    const errors = [];
    const fieldTitle = fieldConfig.title || fieldName;

    // 遍历验证字段的所有验证规则
    for (const validator of fieldConfig.validators) {
      // 非必须字段跳过验证
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

  // 验证一个实体对象(包括主键唯一性)，返回{valid,errors:{fieldName:errors,_composite:errors},rowNumber}
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

  // 验证整个实体集
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

  // 格式化整个实体集的验证结果
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
