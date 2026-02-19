// ============================================================================
// 实体识别器
// 功能：通过检查是否包含实体的所有必填字段来识别实体类型
// 特点：简单直接，一个工作表一次只能导入一个实体
// ============================================================================

class EntityIdentifier {
  constructor() {
    this._config = DataConfig.getInstance();
  }

  /**
   * 识别导入数据的实体类型
   * @param {Array} headers - 导入数据的表头行（标题字符串数组）
   * @returns {string|null} 实体名称，如果无法识别返回null
   */
  identify(headers) {
    if (!headers || headers.length === 0) {
      return null;
    }

    // 标准化表头：去除前后空格
    const normalizedHeaders = headers.map((h) => String(h).trim());

    // 按优先级检查各个实体
    const entityPriority = [
      "ProductSales", // 销售数据优先
      "RegularProduct", // 常态商品
      "ComboProduct", // 组合商品
      "Inventory", // 库存
    ];

    for (const entityName of entityPriority) {
      if (this._matchesEntity(entityName, normalizedHeaders)) {
        return entityName;
      }
    }

    return null;
  }

  /**
   * 检查表头是否匹配实体的必填字段
   */
  _matchesEntity(entityName, headers) {
    const entity = this._config.get(entityName);
    if (!entity) return false;

    const requiredFields = entity.requiredFields || [];
    if (requiredFields.length === 0) return false;

    // 获取必填字段的标题
    const requiredTitles = requiredFields.map((field) => {
      const fieldConfig = entity.fields[field];
      return fieldConfig?.title || field;
    });

    // 检查是否包含所有必填字段
    return requiredTitles.every((title) => headers.includes(title));
  }

  /**
   * 验证实体是否支持导入
   */
  canImport(entityName) {
    const importableEntities = [
      "ProductSales",
      "RegularProduct",
      "ComboProduct",
      "Inventory",
    ];
    return importableEntities.includes(entityName);
  }

  /**
   * 获取实体的导入模式
   * - append: 追加模式（不覆盖，如销售数据）
   * - overwrite: 覆盖模式（如常态商品、库存等）
   */
  getImportMode(entityName) {
    const appendEntities = ["ProductSales"]; // 只有销售数据是追加模式
    return appendEntities.includes(entityName) ? "append" : "overwrite";
  }
}
