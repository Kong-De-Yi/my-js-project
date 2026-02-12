// ============================================================================
// 数据仓库
// 功能：统一数据访问接口，内置缓存、索引、延迟加载
// 特点：纯内存操作，索引加速，自动计算字段
// ============================================================================

class Repository {
  constructor(excelDAO) {
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();

    // 数据缓存
    this._cache = new Map();

    // 索引缓存
    this._indexes = new Map();

    // 计算上下文
    this._context = {
      brandConfig: null,
      profitCalculator: null,
    };
  }

  // 设置计算上下文
  setContext(context) {
    Object.assign(this._context, context);
  }

  // 获取品牌配置映射
  getBrandConfigMap() {
    if (this._context.brandConfig) {
      return this._context.brandConfig;
    }

    try {
      const brands = this.findAll("BrandConfig");
      const brandMap = {};
      brands.forEach((b) => {
        brandMap[b.brandSN] = b;
      });
      this._context.brandConfig = brandMap;
      return brandMap;
    } catch (e) {
      throw new Error(
        "读取品牌配置失败，请确认【品牌配置】工作表存在且数据正确",
      );
    }
  }

  // 刷新品牌配置
  refreshBrandConfig() {
    this._context.brandConfig = null;
    this._cache.delete("BrandConfig");
    this._indexes.delete("BrandConfig");
    return this.getBrandConfigMap();
  }

  // 查找所有（带缓存）
  findAll(entityName) {
    // 检查缓存
    if (this._cache.has(entityName)) {
      return this._cache.get(entityName);
    }

    // 从Excel读取
    const entityConfig = this._config.get(entityName);
    const data = this._excelDAO.read(entityName);

    // 计算计算字段
    this._computeFields(data, entityConfig);

    // 存入缓存
    this._cache.set(entityName, data);

    // 建立索引
    this._buildIndexes(entityName, data, entityConfig);

    return data;
  }

  // 根据索引查询单个
  findOne(entityName, query) {
    const data = this.findAll(entityName);
    const entityConfig = this._config.get(entityName);

    // 尝试使用索引
    if (Object.keys(query).length === 1) {
      const field = Object.keys(query)[0];
      const value = query[field];

      if (entityConfig.uniqueKey === field) {
        const index = this._getIndex(entityName, field);
        if (index && index.has(value)) {
          return index.get(value) || null;
        }
      }
    }

    // 全量查询
    return (
      data.find((item) => {
        return Object.entries(query).every(([key, val]) => item[key] === val);
      }) || null
    );
  }

  // 条件查询
  find(entityName, predicate) {
    const data = this.findAll(entityName);

    if (typeof predicate === "function") {
      return data.filter(predicate);
    }

    if (typeof predicate === "object") {
      return data.filter((item) => {
        return Object.entries(predicate).every(([key, val]) => {
          if (val === undefined) return true;
          return item[key] === val;
        });
      });
    }

    return data;
  }

  // 查询单个货号
  findProductByItemNumber(itemNumber) {
    return this.findOne("Product", { itemNumber });
  }

  // 查询货号列表
  findProducts(query = {}) {
    return this.find("Product", query);
  }

  // 查询商品价格
  findPriceByItemNumber(itemNumber) {
    return this.findOne("ProductPrice", { itemNumber });
  }

  // 查询常态商品
  findRegularProducts(query = {}) {
    return this.find("RegularProduct", query);
  }

  // 查询库存
  findInventory(productCode) {
    return this.findOne("Inventory", { productCode });
  }

  // 查询组合商品
  findComboProducts(productCode) {
    return this.find("ComboProduct", { productCode });
  }

  // 查询销售数据
  findProductSales(query = {}) {
    return this.find("ProductSales", query);
  }

  // 获取系统记录
  getSystemRecord() {
    const records = this.findAll("SystemRecord");
    if (records.length > 0) {
      return records[0];
    }

    // 创建新记录
    const newRecord = {
      updateDateOfLast7Days: undefined,
      updateDateOfProductPrice: undefined,
      updateDateOfRegularProduct: undefined,
      updateDateOfInventory: undefined,
      updateDateOfProductSales: undefined,
      _rowNumber: 1,
    };

    records.push(newRecord);
    this._cache.set("SystemRecord", records);
    return newRecord;
  }

  // 保存数据
  save(entityName, data) {
    // 验证数据
    const entityConfig = this._config.get(entityName);
    const validationResult = validationEngine.validateAll(data, entityConfig);

    if (!validationResult.valid) {
      const errorMsg = validationEngine.formatErrors(
        validationResult,
        entityConfig.worksheet,
      );
      throw new Error(errorMsg);
    }

    // 计算计算字段
    this._computeFields(data, entityConfig);

    // 写入Excel
    this._excelDAO.write(entityName, data);

    // 更新缓存
    this._cache.set(entityName, data);

    // 重建索引
    this._buildIndexes(entityName, data, entityConfig);

    return data;
  }

  // 清空数据
  clear(entityName) {
    this._excelDAO.clear(entityName);
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
  }

  // 刷新缓存
  refresh(entityName) {
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
    return this.findAll(entityName);
  }

  // 批量操作：更新多个实体
  transaction(operations) {
    const results = {};

    try {
      for (const [entityName, data] of Object.entries(operations)) {
        results[entityName] = this.save(entityName, data);
      }
    } catch (e) {
      throw new Error(`批量操作失败：${e.message}`);
    }

    return results;
  }

  // 建立索引
  _buildIndexes(entityName, data, entityConfig) {
    if (!entityConfig.uniqueKey) return;

    if (!this._indexes.has(entityName)) {
      this._indexes.set(entityName, new Map());
    }

    const entityIndexes = this._indexes.get(entityName);
    const field = entityConfig.uniqueKey;

    if (!entityIndexes.has(field)) {
      entityIndexes.set(field, new Map());
    }

    const index = entityIndexes.get(field);
    index.clear();

    data.forEach((item) => {
      const value = item[field];
      if (value !== undefined && value !== null) {
        index.set(value, item);
      }
    });
  }

  // 获取索引
  _getIndex(entityName, field) {
    if (!this._indexes.has(entityName)) {
      return null;
    }

    const entityIndexes = this._indexes.get(entityName);
    return entityIndexes.get(field) || null;
  }

  // 计算计算字段
  _computeFields(data, entityConfig) {
    const computedFields = {};

    Object.entries(entityConfig.fields).forEach(([key, config]) => {
      if (config.type === "computed" && config.compute) {
        computedFields[key] = config.compute;
      }
    });

    if (Object.keys(computedFields).length === 0) {
      return;
    }

    data.forEach((item) => {
      Object.entries(computedFields).forEach(([key, computeFn]) => {
        try {
          item[key] = computeFn(item, this._context);
        } catch (e) {
          // 计算失败时不设置值
        }
      });
    });
  }
}
