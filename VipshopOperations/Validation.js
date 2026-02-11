class ValidationRule {
  constructor(type, options = {}) {
    this.type = type;
    this.options = options;
  }

  validate(value, entity) {
    switch (this.type) {
      case "required":
        return this.validateRequired(value);
      case "number":
        return this.validateNumber(value);
      case "range":
        return this.validateRange(value);
      case "pattern":
        return this.validatePattern(value);
      case "custom":
        return this.validateCustom(value, entity);
      default:
        return { isValid: true };
    }
  }

  validateRequired(value) {
    const isValid = value != null && value !== "";
    return { isValid, message: isValid ? null : "此项为必填项" };
  }

  validateNumber(value) {
    if (value == null || value === "") return { isValid: true }; //空值不验证

    const num = Number(value);
    const isValid = !isNaN(num) && isFinite(num);

    return { isValid, message: isValid ? null : "请输入有效的数字" };
  }

  //其他验证方法
}
