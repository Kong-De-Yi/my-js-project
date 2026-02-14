// ============================================================================
// 数据仓库模块
// 功能：统一数据访问接口，支持单字段和多字段联合索引，内置缓存机制
// 特点：所有字段持久化，联合主键验证，高效查询
// ============================================================================

/**
 * 数据仓库类
 * 负责所有实体的数据缓存、索引管理、查询优化
 */
class Repository {
  /**
   * 构造函数
   * @param {ExcelDAO} excelDAO - Excel数据访问对象
   */
  constructor(excelDAO) {
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();

    // ========== 数据缓存 ==========
    // 格式：Map<实体名称, 实体数据数组>
    this._cache = new Map();

    // ========== 索引存储 ==========
    // 格式：Map<实体名称, Map<索引键, Map<组合键, 数据>>>
    // 第一层：按实体分组
    // 第二层：按索引字段组合分组（如 "brandSN|itemStatus"）
    // 第三层：按实际值组合分组（如 "10000708¦商品上线"）
    this._indexes = new Map();

    // ========== 索引配置 ==========
    // 格式：Map<实体名称, 索引配置数组>
    this._indexConfigs = new Map();

    // ========== 计算上下文 ==========
    // 用于计算字段的上下文对象，如品牌配置、利润计算器等
    this._context = {
      brandConfig: null,
      profitCalculator: null,
    };
  }

  // ==================== 索引配置管理 ====================

  /**
   * 注册实体的索引配置
   * @param {string} entityName - 实体名称
   * @param {Array} indexConfigs - 索引配置数组
   * @example
   * repository.registerIndexes('Product', [
   *   { fields: ['itemNumber'], unique: true },
   *   { fields: ['brandSN', 'itemStatus'], unique: false }
   * ]);
   */
  registerIndexes(entityName, indexConfigs) {
    // 存储索引配置
    this._indexConfigs.set(entityName, indexConfigs || []);

    // 如果已经有缓存数据，立即建立索引
    if (this._cache.has(entityName)) {
      const data = this._cache.get(entityName);
      this._buildAllIndexes(entityName, data);
    }
  }

  // ==================== 上下文管理 ====================

  /**
   * 设置计算上下文
   * @param {Object} context - 上下文对象
   */
  setContext(context) {
    Object.assign(this._context, context);
  }

  /**
   * 获取品牌配置映射
   * @returns {Object} 品牌SN到品牌配置的映射
   */
  getBrandConfigMap() {
    if (this._context.brandConfig) {
      return this._context.brandConfig;
    }

    try {
      const brands = this.findAll("BrandConfig");
      const brandMap = {};
      brands.forEach((b) => {
        if (b.brandSN) {
          brandMap[b.brandSN] = b;
        }
      });
      this._context.brandConfig = brandMap;
      return brandMap;
    } catch (e) {
      throw new Error(
        "读取品牌配置失败，请确认【品牌配置】工作表存在且数据正确",
      );
    }
  }

  /**
   * 刷新品牌配置
   * @returns {Object} 刷新后的品牌配置映射
   */
  refreshBrandConfig() {
    this._context.brandConfig = null;
    this._cache.delete("BrandConfig");
    this._indexes.delete("BrandConfig");
    return this.getBrandConfigMap();
  }

  // ==================== 核心数据访问方法 ====================

  /**
   * 查找实体的所有数据（带缓存）
   * @param {string} entityName - 实体名称
   * @returns {Array} 实体数据数组
   */
  findAll(entityName) {
    // 检查缓存
    if (this._cache.has(entityName)) {
      return this._cache.get(entityName);
    }

    // 从Excel读取数据
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    const data = this._excelDAO.read(entityName);

    // 计算计算字段
    this._computeFields(data, entityConfig);

    // 存入缓存
    this._cache.set(entityName, data);

    // 建立所有索引
    this._buildAllIndexes(entityName, data);

    return data;
  }

  /**
   * 根据条件查询数据（自动使用索引）
   * @param {string} entityName - 实体名称
   * @param {Object} condition - 查询条件
   * @returns {Array} 查询结果数组
   */
  find(entityName, condition) {
    const data = this.findAll(entityName);

    // 如果条件为空，返回所有数据
    if (!condition || Object.keys(condition).length === 0) {
      return data;
    }

    // 尝试使用索引查询
    const indexedResult = this._queryByIndex(entityName, condition);
    if (indexedResult !== null) {
      return indexedResult;
    }

    // 回退到全表扫描
    return this._fullScan(data, condition);
  }

  /**
   * 查询单个数据（自动使用索引）
   * @param {string} entityName - 实体名称
   * @param {Object} condition - 查询条件
   * @returns {Object|null} 查询结果，未找到返回null
   */
  findOne(entityName, condition) {
    // 尝试使用唯一索引
    const entityConfig = this._config.get(entityName);
    const uniqueKey = entityConfig?.uniqueKey;

    if (
      uniqueKey &&
      condition[uniqueKey] !== undefined &&
      Object.keys(condition).length === 1
    ) {
      const indexResult = this._queryByIndex(entityName, {
        [uniqueKey]: condition[uniqueKey],
      });
      if (indexResult && indexResult.length > 0) {
        return indexResult[0];
      }
      return null;
    }

    // 普通查询
    const results = this.find(entityName, condition);
    return results.length > 0 ? results[0] : null;
  }

  /**
   * 复杂条件查询（支持操作符）
   * @param {string} entityName - 实体名称
   * @param {Object} options - 查询选项
   * @returns {Array} 查询结果
   */
  query(entityName, options = {}) {
    let results = this.findAll(entityName);

    // 应用过滤条件
    if (options.filter) {
      results = results.filter((item) => {
        return Object.entries(options.filter).every(([key, condition]) => {
          const value = item[key];

          // 函数式过滤
          if (typeof condition === "function") {
            return condition(value);
          }

          // 操作符过滤
          if (condition && typeof condition === "object") {
            // IN 操作
            if (condition.$in && Array.isArray(condition.$in)) {
              return condition.$in.includes(value);
            }
            // BETWEEN 操作
            if (condition.$between && condition.$between.length === 2) {
              const num = Number(value);
              return (
                num >= condition.$between[0] && num <= condition.$between[1]
              );
            }
            // 大于
            if (condition.$gt !== undefined)
              return Number(value) > condition.$gt;
            // 大于等于
            if (condition.$gte !== undefined)
              return Number(value) >= condition.$gte;
            // 小于
            if (condition.$lt !== undefined)
              return Number(value) < condition.$lt;
            // 小于等于
            if (condition.$lte !== undefined)
              return Number(value) <= condition.$lte;
            // 不等于
            if (condition.$ne !== undefined) return value != condition.$ne;
            // 模糊匹配
            if (condition.$like) {
              const pattern = condition.$like.replace(/%/g, ".*");
              return new RegExp(pattern, "i").test(String(value));
            }
          }

          // 精确匹配
          return value == condition;
        });
      });
    }

    // 应用排序
    if (options.sort) {
      const sortFields = Array.isArray(options.sort)
        ? options.sort
        : [options.sort];
      results.sort((a, b) => {
        for (const sort of sortFields) {
          let field,
            order = 1;
          if (typeof sort === "string") {
            field = sort;
          } else {
            field = sort.field;
            order = sort.order === "desc" ? -1 : 1;
          }

          let valA = a[field];
          let valB = b[field];

          // 数字比较
          if (typeof valA === "number" && typeof valB === "number") {
            if (valA !== valB) return (valA - valB) * order;
          }
          // 日期比较
          else if (field.includes("Date") || field.includes("Time")) {
            const dateA = Date.parse(valA) || 0;
            const dateB = Date.parse(valB) || 0;
            if (dateA !== dateB) return (dateA - dateB) * order;
          }
          // 字符串比较
          else {
            valA = String(valA || "");
            valB = String(valB || "");
            const compare = valA.localeCompare(valB);
            if (compare !== 0) return compare * order;
          }
        }
        return 0;
      });
    }

    // 应用分页
    if (options.limit) {
      const start = options.offset || 0;
      results = results.slice(start, start + options.limit);
    }

    return results;
  }

  // ==================== 快捷查询方法 ====================

  /**
   * 按货号查询商品
   * @param {string} itemNumber - 货号
   * @returns {Object|null} 商品对象
   */
  findProductByItemNumber(itemNumber) {
    return this.findOne("Product", { itemNumber });
  }

  /**
   * 查询所有商品（可带条件）
   * @param {Object} query - 查询条件
   * @returns {Array} 商品数组
   */
  findProducts(query = {}) {
    return this.find("Product", query);
  }

  /**
   * 按款号查询商品
   * @param {string} styleNumber - 款号
   * @returns {Array} 商品数组
   */
  findProductsByStyle(styleNumber) {
    return this.find("Product", { styleNumber });
  }

  /**
   * 按品牌和状态查询商品
   * @param {string} brandSN - 品牌SN
   * @param {string} itemStatus - 商品状态
   * @returns {Array} 商品数组
   */
  findProductsByBrandAndStatus(brandSN, itemStatus) {
    return this.find("Product", { brandSN, itemStatus });
  }

  /**
   * 按SPU查询商品
   * @param {string} P_SPU - SPU编码
   * @returns {Array} 商品数组
   */
  findProductsBySPU(P_SPU) {
    return this.find("Product", { P_SPU });
  }

  /**
   * 查询商品价格
   * @param {string} itemNumber - 货号
   * @returns {Object|null} 价格对象
   */
  findPriceByItemNumber(itemNumber) {
    return this.findOne("ProductPrice", { itemNumber });
  }

  /**
   * 查询常态商品
   * @param {Object} query - 查询条件
   * @returns {Array} 常态商品数组
   */
  findRegularProducts(query = {}) {
    return this.find("RegularProduct", query);
  }

  /**
   * 查询库存
   * @param {string} productCode - 商品编码
   * @returns {Object|null} 库存对象
   */
  findInventory(productCode) {
    return this.findOne("Inventory", { productCode });
  }

  /**
   * 查询组合商品
   * @param {string} productCode - 组合商品编码
   * @returns {Array} 组合商品数组
   */
  findComboProducts(productCode) {
    return this.find("ComboProduct", { productCode });
  }

  /**
   * 查询销售数据
   * @param {Object} query - 查询条件
   * @returns {Array} 销售数据数组
   */
  findProductSales(query = {}) {
    return this.find("ProductSales", query);
  }

  /**
   * 获取系统记录
   * @returns {Object} 系统记录对象
   */
  getSystemRecord() {
    const records = this.findAll("SystemRecord");
    if (records.length > 0) {
      return records[0];
    }

    // 创建新记录
    const newRecord = {
      updateDateOfLast7Days: "",
      updateDateOfProductPrice: "",
      updateDateOfRegularProduct: "",
      updateDateOfInventory: "",
      updateDateOfProductSales: "",
      _rowNumber: 1,
    };

    // 保存新记录
    this._cache.set("SystemRecord", [newRecord]);
    return newRecord;
  }

  /**
   * 查询库存预警商品
   * @param {number} threshold - 预警阈值（天数）
   * @returns {Array} 预警商品列表
   */
  findLowStockProducts(threshold = 30) {
    return this.query("Product", {
      filter: {
        sellableDays: { $lt: threshold },
        itemStatus: { $in: ["商品上线", "部分上线"] },
      },
      sort: { field: "sellableDays", order: "asc" },
    });
  }

  /**
   * 查询高利润商品
   * @param {number} minProfit - 最小利润
   * @param {number} minRate - 最小利润率
   * @returns {Array} 高利润商品列表
   */
  findHighProfitProducts(minProfit = 10, minRate = 0.3) {
    return this.query("Product", {
      filter: (item) =>
        (item.profit || 0) >= minProfit && (item.profitRate || 0) >= minRate,
      sort: [
        { field: "profit", order: "desc" },
        { field: "profitRate", order: "desc" },
      ],
    });
  }

  /**
   * 按日期范围查询销售数据
   * @param {string} startDate - 开始日期
   * @param {string} endDate - 结束日期
   * @returns {Array} 销售数据数组
   */
  findSalesByDateRange(startDate, endDate) {
    const start = validationEngine.parseDate(startDate)?.getTime() || 0;
    const end = validationEngine.parseDate(endDate)?.getTime() || Infinity;

    return this.query("ProductSales", {
      filter: (item) => {
        const date = validationEngine.parseDate(item.salesDate)?.getTime() || 0;
        return date >= start && date <= end;
      },
      sort: { field: "salesDate", order: "asc" },
    });
  }

  // ==================== 数据修改方法 ====================

  /**
   * 保存数据（带验证）
   * @param {string} entityName - 实体名称
   * @param {Array} data - 要保存的数据数组
   * @returns {Array} 保存后的数据
   */
  save(entityName, data) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    // 日期字段标准化
    data.forEach((item) => {
      Object.entries(entityConfig.fields).forEach(
        ([fieldName, fieldConfig]) => {
          if (
            fieldConfig.type === "date" ||
            fieldConfig.validators?.some((v) => v.type === "date")
          ) {
            if (item[fieldName]) {
              const date = validationEngine.parseDate(item[fieldName]);
              if (date) {
                item[fieldName] = validationEngine.formatDate(date);
              }
            }
          }
        },
      );
    });

    // 验证数据（使用ValidationEngine）
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

    // 重建所有索引
    this._buildAllIndexes(entityName, data);

    return data;
  }

  /**
   * 清空实体数据
   * @param {string} entityName - 实体名称
   */
  clear(entityName) {
    this._excelDAO.clear(entityName);
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
  }

  /**
   * 刷新实体缓存
   * @param {string} entityName - 实体名称
   * @returns {Array} 刷新后的数据
   */
  refresh(entityName) {
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
    return this.findAll(entityName);
  }

  /**
   * 批量操作（事务）
   * @param {Object} operations - 操作对象 {实体名称: 数据数组}
   * @returns {Object} 操作结果
   */
  transaction(operations) {
    const results = {};
    const errors = [];

    for (const [entityName, data] of Object.entries(operations)) {
      try {
        results[entityName] = this.save(entityName, data);
      } catch (e) {
        errors.push(`【${entityName}】保存失败：${e.message}`);
      }
    }

    if (errors.length > 0) {
      throw new Error(`批量操作失败：\n${errors.join("\n")}`);
    }

    return results;
  }

  // ==================== 私有辅助方法 ====================

  /**
   * 建立实体的所有索引
   * @param {string} entityName - 实体名称
   * @param {Array} data - 实体数据
   * @private
   */
  _buildAllIndexes(entityName, data) {
    // 确保实体有索引容器
    if (!this._indexes.has(entityName)) {
      this._indexes.set(entityName, new Map());
    }

    const entityIndexes = this._indexes.get(entityName);
    const indexConfigs = this._indexConfigs.get(entityName) || [];

    // 添加默认唯一键索引
    const entityConfig = this._config.get(entityName);
    if (entityConfig?.uniqueKey) {
      // 检查是否已存在该索引
      const exists = indexConfigs.some(
        (config) =>
          config.fields.length === 1 &&
          config.fields[0] === entityConfig.uniqueKey,
      );
      if (!exists) {
        indexConfigs.push({
          fields: [entityConfig.uniqueKey],
          unique: true,
        });
      }
    }

    // 添加联合主键索引
    if (entityConfig?.compositeKey) {
      const exists = indexConfigs.some(
        (config) =>
          JSON.stringify(config.fields.sort()) ===
          JSON.stringify(entityConfig.compositeKey.fields.sort()),
      );
      if (!exists) {
        indexConfigs.push({
          fields: entityConfig.compositeKey.fields,
          unique: true,
        });
      }
    }

    // 构建每个索引
    indexConfigs.forEach((config) => {
      // 关键点：对索引字段进行排序，保证索引键标准化
      const sortedFields = [...config.fields].sort();
      const indexKey = sortedFields.join("|");

      if (!entityIndexes.has(indexKey)) {
        entityIndexes.set(indexKey, new Map());
      }

      const index = entityIndexes.get(indexKey);
      index.clear();

      data.forEach((item) => {
        const value = this._getCompositeKey(item, sortedFields);

        if (value !== undefined && value !== null && value !== "") {
          if (config.unique) {
            // 唯一索引：检查重复
            if (index.has(value)) {
              // 记录错误但不中断
              item._indexError = `索引${indexKey}值"${value}"重复`;
            } else {
              index.set(value, item);
            }
          } else {
            // 非唯一索引
            if (!index.has(value)) {
              index.set(value, []);
            }
            index.get(value).push(item);
          }
        }
      });
    });
  }

  /**
   * 根据索引查询
   * @param {string} entityName - 实体名称
   * @param {Object} condition - 查询条件
   * @returns {Array|null} 查询结果，null表示无可用索引
   * @private
   */
  _queryByIndex(entityName, condition) {
    const entityIndexes = this._indexes.get(entityName);
    if (!entityIndexes) return null;

    // 关键点：对条件字段排序，保证能匹配标准化的索引键
    const conditionFields = Object.keys(condition).sort();

    // 1. 先尝试完全匹配（所有字段的联合索引）
    const exactMatchKey = conditionFields.join("|");
    if (entityIndexes.has(exactMatchKey)) {
      const index = entityIndexes.get(exactMatchKey);
      const compositeValue = this._getCompositeKey(condition, conditionFields);

      if (index.has(compositeValue)) {
        const result = index.get(compositeValue);
        return Array.isArray(result) ? result : [result];
      }
      return [];
    }

    // 2. 尝试前缀匹配（使用前N个字段）
    for (let i = conditionFields.length; i > 0; i--) {
      const prefixFields = conditionFields.slice(0, i);
      const prefixKey = prefixFields.join("|");

      if (entityIndexes.has(prefixKey)) {
        const index = entityIndexes.get(prefixKey);
        const results = [];

        // 遍历索引，筛选出匹配前缀的数据
        for (let [compositeValue, items] of index.entries()) {
          const itemCondition = this._parseCompositeKey(
            prefixKey,
            compositeValue,
          );
          let match = true;

          for (let j = 0; j < prefixFields.length; j++) {
            const field = prefixFields[j];
            // 使用!=进行宽松比较（处理数字和字符串的自动转换）
            if (itemCondition[field] != condition[field]) {
              match = false;
              break;
            }
          }

          if (match) {
            if (Array.isArray(items)) {
              results.push(...items);
            } else {
              results.push(items);
            }
          }
        }

        if (results.length > 0) {
          return results;
        }
      }
    }

    return null; // 未找到可用索引
  }

  /**
   * 全表扫描（当没有可用索引时）
   * @param {Array} data - 数据数组
   * @param {Object} condition - 查询条件
   * @returns {Array} 查询结果
   * @private
   */
  _fullScan(data, condition) {
    return data.filter((item) => {
      return Object.entries(condition).every(([key, val]) => {
        if (val === undefined) return true;
        if (Array.isArray(val)) {
          return val.includes(item[key]);
        }
        return item[key] == val;
      });
    });
  }

  /**
   * 获取组合键值
   * @param {Object} item - 数据对象
   * @param {Array} fields - 字段数组（已排序）
   * @returns {string} 组合键值
   * @private
   */
  _getCompositeKey(item, fields) {
    if (fields.length === 1) {
      const value = item[fields[0]];
      return value !== undefined && value !== null ? String(value) : "";
    }

    // 多字段组合：用特殊分隔符连接
    // 使用'¦'（断条符号）作为分隔符，避免与数据中的字符混淆
    return fields
      .map((f) => {
        const value = item[f];
        return value !== undefined && value !== null ? String(value) : "";
      })
      .join("¦");
  }

  /**
   * 解析组合键为条件对象
   * @param {string} indexKey - 索引键（如 "brandSN|itemStatus"）
   * @param {string} compositeValue - 组合值（如 "10000708¦商品上线"）
   * @returns {Object} 条件对象
   * @private
   */
  _parseCompositeKey(indexKey, compositeValue) {
    const fields = indexKey.split("|");
    const values = String(compositeValue).split("¦");

    const condition = {};
    fields.forEach((field, i) => {
      condition[field] = values[i] === "" ? undefined : values[i];
    });

    return condition;
  }

  /**
   * 计算计算字段
   * @param {Array} data - 数据数组
   * @param {Object} entityConfig - 实体配置
   * @private
   */
  _computeFields(data, entityConfig) {
    const computedFields = {};

    // 找出所有计算字段
    Object.entries(entityConfig.fields).forEach(([key, config]) => {
      if (config.type === "computed" && config.compute) {
        computedFields[key] = config.compute;
      }
    });

    if (Object.keys(computedFields).length === 0) {
      return;
    }

    // 为每条数据计算计算字段
    data.forEach((item) => {
      Object.entries(computedFields).forEach(([key, computeFn]) => {
        try {
          item[key] = computeFn(item, this._context);
        } catch (e) {
          // 计算失败时不设置值，保持undefined
        }
      });
    });
  }

  /**
   * 获取索引统计信息（用于调试）
   * @param {string} entityName - 实体名称
   * @returns {Object} 索引统计信息
   */
  getIndexStats(entityName) {
    const entityIndexes = this._indexes.get(entityName);
    if (!entityIndexes) {
      return { hasIndexes: false };
    }

    const stats = {
      hasIndexes: true,
      indexes: [],
    };

    for (let [indexKey, index] of entityIndexes.entries()) {
      stats.indexes.push({
        key: indexKey,
        size: index.size,
        fields: indexKey.split("|"),
      });
    }

    return stats;
  }

  /**
   * 清空所有缓存（用于调试或重置）
   */
  clearAllCache() {
    this._cache.clear();
    this._indexes.clear();
  }
}

// 确保全局可用
if (typeof module !== "undefined" && module.exports) {
  module.exports = Repository;
}
