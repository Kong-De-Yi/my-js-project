// ============================================================================
// 数据验证引擎
// 功能：独立的数据校验器，与业务逻辑完全解耦
// 特点：可扩展、批量验证、精确定位错误行号
// ============================================================================

class ValidationEngine {
  constructor() {
    this._validators = {
      // 必填验证
      required: (value, params) => ({
        valid: value !== undefined && value !== null && value !== "",
        message: params?.message || "不能为空",
      }),

      // 枚举值验证
      enum: (value, params) => ({
        valid: value === undefined || params.values.includes(value),
        message:
          params?.message || `必须是以下值之一：${params.values.join(", ")}`,
      }),

      // 正则验证
      pattern: (value, params) => ({
        valid: value === undefined || params.regex.test(String(value)),
        message:
          params?.message ||
          `格式不正确${params.description ? "：" + params.description : ""}`,
      }),

      // 数字范围
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

      // 非负数
      nonNegative: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        if (isNaN(num)) return { valid: false, message: "必须是数字" };
        return { valid: num >= 0, message: params?.message || "不能为负数" };
      },

      // 正数
      positive: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        if (isNaN(num)) return { valid: false, message: "必须是数字" };
        return { valid: num > 0, message: params?.message || "必须大于0" };
      },

      // 数字类型
      number: (value, params) => {
        if (value === undefined) return { valid: true };
        const num = Number(value);
        return {
          valid: !isNaN(num) && isFinite(num),
          message: params?.message || "必须是有效的数字",
        };
      },

      //时间
      date: (value, params) => {
        if (value === undefined || value === null || value === "") {
          return { valid: true };
        }

        // 移除可能的前导引号
        const cleanValue = String(value).replace(/^'/, "");

        // 尝试解析日期
        const timestamp = Date.parse(cleanValue);
        const isValidDate = !isNaN(timestamp);

        return {
          valid: isValidDate,
          message: params?.message || "不是有效的日期格式",
        };
      },
    };
  }

  // 注册自定义验证器
  register(name, validatorFn) {
    this._validators[name] = validatorFn;
  }

  // 验证单个值
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

  // 验证单个字段
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

  // 验证整个实体
  validateEntity(entity, entityConfig, context = {}) {
    const errors = {};

    Object.entries(entityConfig.fields).forEach(([fieldName, fieldConfig]) => {
      if (fieldConfig.type === "computed") return;

      const value = entity[fieldName];
      const result = this.validateField(value, fieldConfig, fieldName);

      if (!result.valid) {
        errors[fieldName] = result.errors;
      }
    });

    // 唯一性验证
    if (entityConfig.uniqueKey && context.allData) {
      const key = entityConfig.uniqueKey;
      const value = entity[key];

      if (value) {
        const duplicates = context.allData.filter(
          (item) => item !== entity && item[key] === value,
        );

        if (duplicates.length > 0) {
          const fieldTitle = entityConfig.fields[key]?.title || key;
          errors[key] = errors[key] || [];
          errors[key].push(`【${fieldTitle}】值"${value}"已存在`);
        }
      }
    }

    return {
      valid: Object.keys(errors).length === 0,
      errors,
      rowNumber: entity._rowNumber,
    };
  }

  // 批量验证
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

  // 格式化错误消息，用于显示给用户
  formatErrors(validationResult, entityName) {
    if (validationResult.valid) {
      return null;
    }

    let message = `【${entityName}】数据验证失败：\n`;

    validationResult.items.forEach((item) => {
      if (!item.valid) {
        message += `\n第${item.rowNumber || "?"}行：\n`;
        Object.entries(item.errors).forEach(([field, fieldErrors]) => {
          fieldErrors.forEach((err) => {
            message += `  ${err}\n`;
          });
        });
      }
    });

    return message;
  }
}

// 单例导出
const validationEngine = new ValidationEngine();
