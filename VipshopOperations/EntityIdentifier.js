// ============================================================================
// 实体识别器
// 功能：通过检查是否包含实体的所有必填字段来识别实体类型
// 特点：简单直接，一个工作表一次只能导入一个实体
// ============================================================================

class EntityIdentifier {
  static _instance = null;

  constructor() {
    if (EntityIdentifier._instance) {
      return EntityIdentifier._instance;
    }

    this._config = DataConfig.getInstance();
    this._importableEntities = [];

    // 获取可导入业务实体
    for (const [key, value] of Object.entries(this._config.getAll())) {
      if (value?.canImport === true) {
        this._importableEntities.push(key);
      }
    }

    EntityIdentifier._instance = this;
  }

  // 检查表头是否匹配实体的必填字段
  _matchesEntity(entityName, headers) {
    const entity = this._config.get(entityName);
    if (!entity) return false;

    const requiredTitles = entity.requiredTitles || [];
    if (requiredTitles.length === 0) return false;

    // 检查是否包含所有必填字段
    return requiredTitles.every((title) => headers.includes(title));
  }

  // 单例模式
  static getInstance() {
    if (!EntityIdentifier._instance) {
      EntityIdentifier._instance = new EntityIdentifier();
    }
    return EntityIdentifier._instance;
  }

  // 识别导入数据的实体类型
  identify(headers) {
    if (!headers || headers.length === 0) {
      return null;
    }

    // 标准化表头：去除前后空格
    const normalizedHeaders = headers.map((h) => String(h).trim());

    for (const entityName of this._importableEntities) {
      if (this._matchesEntity(entityName, normalizedHeaders)) {
        return entityName;
      }
    }

    return null;
  }

  // 返回可以导入的所有实体名称
  getImportableEntities() {
    return this._importableEntities;
  }

  // 验证实体是否支持导入
  canImport(entityName) {
    return this._importableEntities.includes(entityName);
  }

  // 获取实体的导入模式
  getImportMode(entityName) {
    if (!this.canImport(entityName)) return null;

    const entity = this._config.get(entityName);
    return entity?.importMode || null;
  }
}
