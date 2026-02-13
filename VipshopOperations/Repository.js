// ============================================================================
// 数据仓库（增强版）
// 功能：支持单字段索引和多字段联合索引，大幅提升复杂查询性能
// 特点：自动维护索引，支持组合查询，O(1)时间复杂度
// ============================================================================

class Repository {
  constructor(excelDAO) {
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();

    // 数据缓存
    this._cache = new Map();

    // 索引存储结构：
    // _indexes.get(entityName) -> Map {
    //   'field1' -> Map(value -> data)                    // 单字段索引
    //   'field1|field2' -> Map(key -> data)               // 联合索引
    //   'field1|field2|field3' -> Map(compositeKey -> data) // 多字段索引
    // }
    this._indexes = new Map();

    // 索引配置
    this._indexConfigs = new Map(); // 实体名 -> 索引配置数组

    // 计算上下文
    this._context = {
      brandConfig: null,
      profitCalculator: null,
    };
  }

  // ========== 索引管理方法 ==========

  /**
   * 注册索引配置
   * @param {string} entityName - 实体名称
   * @param {Array} indexConfigs - 索引配置数组
   * @example
   * registerIndexes('Product', [
   *   { fields: ['brandSN', 'itemStatus'], unique: false },  // 联合索引：品牌+状态
   *   { fields: ['styleNumber', 'color'], unique: true },    // 唯一联合索引：款号+颜色
   *   { fields: ['firstListingTime'], unique: false }        // 单字段索引
   * ]);
   */
  registerIndexes(entityName, indexConfigs) {
    this._indexConfigs.set(entityName, indexConfigs);

    // 如果已经有缓存数据，立即建立索引
    if (this._cache.has(entityName)) {
      const data = this._cache.get(entityName);
      this._buildAllIndexes(entityName, data);
    }
  }

  /**
   * 构建所有索引
   */
  _buildAllIndexes(entityName, data) {
    if (!this._indexes.has(entityName)) {
      this._indexes.set(entityName, new Map());
    }

    const entityIndexes = this._indexes.get(entityName);
    const indexConfigs = this._indexConfigs.get(entityName) || [];

    // 添加默认唯一键索引
    const entityConfig = this._config.get(entityName);
    if (entityConfig.uniqueKey) {
      indexConfigs.push({
        fields: [entityConfig.uniqueKey],
        unique: true,
      });
    }

    // 构建每个索引
    indexConfigs.forEach((config) => {
      const indexKey = config.fields.join("|");

      if (!entityIndexes.has(indexKey)) {
        entityIndexes.set(indexKey, new Map());
      }

      const index = entityIndexes.get(indexKey);
      index.clear();

      data.forEach((item) => {
        const value = this._getCompositeKey(item, config.fields);

        if (value !== undefined && value !== null && value !== "") {
          if (config.unique) {
            // 唯一索引：直接映射到单个数据项
            index.set(value, item);
          } else {
            // 非唯一索引：映射到数据数组
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
   * 获取组合键值
   */
  _getCompositeKey(item, fields) {
    if (fields.length === 1) {
      return item[fields[0]];
    }

    // 多字段组合：用特殊分隔符连接
    return fields.map((f) => item[f] ?? "").join("¦"); // 使用不常用字符作为分隔符
  }

  /**
   * 解析组合键为条件对象
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
   * 根据索引查询
   * @param {string} entityName - 实体名称
   * @param {Object} condition - 查询条件
   * @returns {Array} 查询结果
   */
  _queryByIndex(entityName, condition) {
    const entityIndexes = this._indexes.get(entityName);
    if (!entityIndexes) return null;

    const conditionFields = Object.keys(condition).sort();

    // 尝试匹配最合适的索引
    // 1. 先尝试完全匹配
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
            if (itemCondition[field] != condition[field]) {
              // 使用!=进行宽松比较
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

  // ========== 核心数据访问方法 ==========

  /**
   * 设置计算上下文
   */
  setContext(context) {
    Object.assign(this._context, context);
  }

  /**
   * 获取品牌配置映射
   */
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

  /**
   * 刷新品牌配置
   */
  refreshBrandConfig() {
    this._context.brandConfig = null;
    this._cache.delete("BrandConfig");
    this._indexes.delete("BrandConfig");
    return this.getBrandConfigMap();
  }

  /**
   * 查找所有（带缓存）
   */
  findAll(entityName) {
    if (this._cache.has(entityName)) {
      return this._cache.get(entityName);
    }

    const entityConfig = this._config.get(entityName);
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
   * 根据条件查询（自动使用索引）
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

    // 回退到全量遍历
    return data.filter((item) => {
      return Object.entries(condition).every(([key, val]) => {
        if (val === undefined) return true;
        if (Array.isArray(val)) {
          return val.includes(item[key]);
        }
        return item[key] == val; // 使用==进行宽松比较
      });
    });
  }

  /**
   * 查询单个（自动使用索引）
   */
  findOne(entityName, condition) {
    // 尝试使用唯一索引
    const entityConfig = this._config.get(entityName);
    const uniqueKey = entityConfig.uniqueKey;

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
   * 复杂条件查询（支持多条件组合）
   */
  query(entityName, options = {}) {
    let results = this.findAll(entityName);

    // 应用过滤条件
    if (options.filter) {
      results = results.filter((item) => {
        return Object.entries(options.filter).every(([key, condition]) => {
          const value = item[key];

          if (typeof condition === "function") {
            return condition(value);
          }

          if (condition && typeof condition === "object") {
            if (condition.$in && Array.isArray(condition.$in)) {
              return condition.$in.includes(value);
            }
            if (condition.$between && condition.$between.length === 2) {
              const num = Number(value);
              return (
                num >= condition.$between[0] && num <= condition.$between[1]
              );
            }
            if (condition.$gt !== undefined)
              return Number(value) > condition.$gt;
            if (condition.$gte !== undefined)
              return Number(value) >= condition.$gte;
            if (condition.$lt !== undefined)
              return Number(value) < condition.$lt;
            if (condition.$lte !== undefined)
              return Number(value) <= condition.$lte;
            if (condition.$ne !== undefined) return value != condition.$ne;
            if (condition.$like) {
              const pattern = condition.$like.replace(/%/g, ".*");
              return new RegExp(pattern, "i").test(String(value));
            }
          }

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

          if (typeof valA === "number" && typeof valB === "number") {
            if (valA !== valB) return (valA - valB) * order;
          } else {
            valA = String(valA || "");
            valB = String(valB || "");
            if (valA !== valB) return valA.localeCompare(valB) * order;
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

  // ========== 快捷查询方法 ==========

  /**
   * 按货号查询
   */
  findProductByItemNumber(itemNumber) {
    return this.findOne("Product", { itemNumber });
  }

  /**
   * 按款号查询所有颜色
   */
  findProductsByStyle(styleNumber) {
    return this.find("Product", { styleNumber });
  }

  /**
   * 按品牌+状态查询
   */
  findProductsByBrandAndStatus(brandSN, itemStatus) {
    return this.find("Product", { brandSN, itemStatus });
  }

  /**
   * 按SPU查询
   */
  findProductsBySPU(P_SPU) {
    return this.find("Product", { P_SPU });
  }

  /**
   * 按货号范围查询
   */
  findProductsByItemNumberRange(start, end) {
    return this.query("Product", {
      filter: {
        itemNumber: {
          $between: [start, end],
        },
      },
    });
  }

  /**
   * 按售龄范围查询
   */
  findProductsBySalesAge(minAge, maxAge) {
    return this.query("Product", {
      filter: {
        salesAge: {
          $between: [minAge, maxAge],
        },
      },
    });
  }

  /**
   * 查询库存预警商品（可售天数<阈值）
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

  // ========== 数据修改方法 ==========

  /**
   * 保存数据（带验证）
   */
  save(entityName, data) {
    const entityConfig = this._config.get(entityName);

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

    // 验证数据
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
   * 清空数据
   */
  clear(entityName) {
    this._excelDAO.clear(entityName);
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
  }

  /**
   * 刷新缓存
   */
  refresh(entityName) {
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
    return this.findAll(entityName);
  }

  /**
   * 批量操作
   */
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

  // ========== 辅助方法 ==========

  /**
   * 建立所有索引
   */
  _buildAllIndexes(entityName, data) {
    if (!this._indexes.has(entityName)) {
      this._indexes.set(entityName, new Map());
    }

    const entityIndexes = this._indexes.get(entityName);
    const indexConfigs = this._indexConfigs.get(entityName) || [];

    // 添加默认唯一键索引
    const entityConfig = this._config.get(entityName);
    if (entityConfig.uniqueKey) {
      indexConfigs.push({
        fields: [entityConfig.uniqueKey],
        unique: true,
      });
    }

    // 构建每个索引
    indexConfigs.forEach((config) => {
      const indexKey = config.fields.join("|");

      if (!entityIndexes.has(indexKey)) {
        entityIndexes.set(indexKey, new Map());
      }

      const index = entityIndexes.get(indexKey);
      index.clear();

      data.forEach((item) => {
        const value = this._getCompositeKey(item, config.fields);

        if (value !== undefined && value !== null && value !== "") {
          if (config.unique) {
            // 唯一索引：直接映射到单个数据项
            // 检查重复
            if (index.has(value)) {
              item._indexError = `联合索引${config.fields.join("+")}值重复`;
            }
            index.set(value, item);
          } else {
            // 非唯一索引：映射到数据数组
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
   * 获取组合键值
   */
  _getCompositeKey(item, fields) {
    if (fields.length === 1) {
      return item[fields[0]];
    }

    // 多字段组合：用特殊分隔符连接
    return fields.map((f) => item[f] ?? "").join("¦");
  }

  /**
   * 根据索引查询
   */
  _queryByIndex(entityName, condition) {
    const entityIndexes = this._indexes.get(entityName);
    if (!entityIndexes) return null;

    const conditionFields = Object.keys(condition).sort();

    // 尝试完全匹配
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

    // 尝试前缀匹配
    for (let i = conditionFields.length; i > 0; i--) {
      const prefixFields = conditionFields.slice(0, i);
      const prefixKey = prefixFields.join("|");

      if (entityIndexes.has(prefixKey)) {
        const index = entityIndexes.get(prefixKey);
        const results = [];

        for (let [compositeValue, items] of index.entries()) {
          const itemCondition = this._parseCompositeKey(
            prefixKey,
            compositeValue,
          );
          let match = true;

          for (let j = 0; j < prefixFields.length; j++) {
            const field = prefixFields[j];
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

    return null;
  }

  /**
   * 解析组合键
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
   */
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
