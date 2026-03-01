/**
 * 数据仓库 - 核心数据管理类，提供类似数据库的访问模式
 *
 * @class Repository
 * @description 作为系统的数据访问核心，提供以下功能：
 * - 数据缓存管理，减少 Excel 重复读取
 * - 多字段索引的自动建立和查询优化
 * - 支持精确匹配、前缀匹配的索引查询
 * - 复杂查询（过滤、排序、分页）
 * - 数据验证和计算字段处理
 * - CRUD 操作（增删改查）
 * - 事务性批量操作
 *
 * 该类采用单例模式，确保全局只有一个数据仓库实例。
 *
 * @example
 * // 获取仓库实例
 * const repository = Repository.getInstance(excelDAO);
 *
 * // 查询所有产品
 * const allProducts = repository.findAll("Product");
 *
 * // 条件查询
 * const onlineProducts = repository.find("Product", { itemStatus: "商品上线" });
 *
 * // 复杂查询
 * const results = repository.query("Product", {
 *   filter: { price: { $gt: 100 } },
 *   sort: { field: "profit", order: "desc" },
 *   limit: 20
 * });
 */
class Repository {
  /** @type {Repository} 单例实例 */
  static _instance = null;

  /**
   * 创建仓库实例
   * @param {ExcelDAO} [excelDAO] - Excel 数据访问对象，若不提供则自动获取
   */
  constructor(excelDAO) {
    if (Repository._instance) {
      return Repository._instance;
    }

    this._excelDAO = excelDAO || ExcelDAO.getInstance(); // 可以传入，也可以自动获取;
    this._config = DataConfig.getInstance();
    this._importableEntities = this._config.getImportableEntities();
    this._updatableEntities = this._config.getUpdatableEntities();
    this._validationEngine = ValidationEngine.getInstance();
    this._converter = Converter.getInstance();

    this._cache = new Map();
    this._indexes = new Map();
    this._indexConfigs = new Map();

    this._context = {
      brandConfig: null,
      profitCalculator: null,
    };

    Repository._instance = this;
  }

  /**
   * 获取仓库单例实例
   * @static
   * @param {ExcelDAO} [excelDAO] - Excel 数据访问对象
   * @returns {Repository} 仓库实例
   */
  static getInstance(excelDAO) {
    if (!Repository._instance) {
      Repository._instance = new Repository(excelDAO);
    }
    return Repository._instance;
  }

  /**
   * 构建复合键值字符串
   * @private
   * @param {Object} item - 数据项
   * @param {string[]} fields - 字段名数组
   * @returns {string} 用'¦'连接的复合键值，空字段转为空字符串
   *
   * @example
   * // 返回 "A001¦红色¦M"
   * _getCompositeKey(product, ["itemNumber", "color", "size"])
   */
  _getCompositeKey(item, fields) {
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
   * 构建实体的所有索引
   * @private
   * @param {string} entityName - 实体名称
   * @param {Object[]} data - 实体数据数组
   * @description
   * 根据索引配置建立多字段索引：
   * - 自动添加主键索引（如果配置了uniqueKey）
   * - 支持唯一索引和非唯一索引
   * - 索引键为排序后的字段组合
   */
  _buildAllIndexes(entityName, data) {
    if (!this._indexes.has(entityName)) {
      this._indexes.set(entityName, new Map());
    }

    const entityIndexes = this._indexes.get(entityName);
    let indexConfigs = this._indexConfigs.get(entityName) || [];

    const entityConfig = this._config.get(entityName);

    // 检查索引配置中是否有配置主键索引,没有则添加
    if (entityConfig?.uniqueKey) {
      const uniqueKeyConfig = this._config.parseUniqueKey(
        entityConfig.uniqueKey,
      );
      const fields = uniqueKeyConfig.fields;

      if (fields.length > 0) {
        const exists = indexConfigs.some(
          (config) =>
            JSON.stringify([...config.fields].sort()) ===
            JSON.stringify([...fields].sort()),
        );

        if (!exists) {
          indexConfigs.push({
            fields: fields,
            unique: true,
          });
        }
      }
    }

    // 根据索引配置建立索引
    indexConfigs.forEach((config) => {
      const sortedFields = [...config.fields].sort();
      const indexKey = sortedFields.join("|");

      if (!entityIndexes.has(indexKey)) {
        entityIndexes.set(indexKey, new Map());
      }

      const index = entityIndexes.get(indexKey);
      index.clear();

      data.forEach((item) => {
        // 构建组合索引值
        const value = this._getCompositeKey(item, sortedFields);

        if (value != undefined && String(value).trim() !== "") {
          // 唯一性索引值直接覆盖
          if (config.unique) {
            index.set(value, item);
          } else {
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
   * 通过索引查询数据
   * @private
   * @param {string} entityName - 实体名称
   * @param {Object} condition - 查询条件
   * @returns {Object[]|null} 查询结果数组，无可用索引时返回null
   *
   * @description
   * 索引查询策略（按优先级）：
   * 1. 精确匹配：条件字段完全匹配某个索引
   * 2. 前缀匹配：使用部分字段的索引进行过滤，再对剩余条件扫描
   * 3. 无可用索引：返回null，让上层执行全表扫描
   *
   * @example
   * // 使用品牌+年份索引查询
   * _queryByIndex("Product", { brand: "A", year: 2024 })
   */
  _queryByIndex(entityName, condition) {
    const entityIndexes = this._indexes.get(entityName);
    if (!entityIndexes) return null;

    const conditionFields = Object.keys(condition).sort();

    // 1. 精确匹配
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

    // 2. 前缀匹配
    for (let i = conditionFields.length; i > 0; i--) {
      const prefixFields = conditionFields.slice(0, i);
      const prefixKey = prefixFields.join("|");
      const remainingFields = conditionFields.slice(i); // 剩余字段

      if (entityIndexes.has(prefixKey)) {
        const index = entityIndexes.get(prefixKey);
        const results = [];

        for (let [compositeValue, items] of index.entries()) {
          // 解析前缀值
          const prefixValues = compositeValue.split("¦");

          // 检查前缀是否匹配
          let prefixMatch = true;
          for (let j = 0; j < prefixFields.length; j++) {
            const field = prefixFields[j];
            const prefixVal = prefixValues[j];

            // 如果条件中该字段未定义，视为匹配
            if (condition[field] != undefined) {
              if (String(prefixVal) != String(condition[field])) {
                prefixMatch = false;
                break;
              }
            }
          }

          if (prefixMatch) {
            // 对匹配前缀的记录，检查剩余条件
            const records = Array.isArray(items) ? items : [items];

            const matchedRecords = records.filter((record) => {
              return remainingFields.every((field) => {
                // 如果条件中没有这个字段，视为匹配
                if (condition[field] == undefined) return true;
                return record[field] == condition[field];
              });
            });

            if (matchedRecords.length > 0) {
              results.push(...matchedRecords);
            }
          }
        }

        if (results.length > 0) {
          return results;
        }
      }
    }

    // 3. 没有可用索引，返回 null 让上层做全表扫描
    return null;
  }

  /**
   * 全表扫描查询
   * @private
   * @param {Object[]} data - 数据数组
   * @param {Object} condition - 查询条件
   * @returns {Object[]} 符合条件的记录
   *
   * @description
   * 支持的条件类型：
   * - 精确匹配：{ field: value }
   * - 数组匹配：{ field: [value1, value2] }（满足任一值）
   * - undefined值自动忽略
   */
  _fullScan(data, condition) {
    return data.filter((item) => {
      return Object.entries(condition).every(([key, val]) => {
        if (val == undefined) return true;
        if (Array.isArray(val)) {
          return val.includes(item[key]);
        }
        return item[key] == val;
      });
    });
  }

  /**
   * 计算实体的计算字段
   * @private
   * @param {Object[]} data - 数据数组
   * @param {Object} entityConfig - 实体配置
   * @description
   * 根据字段配置中的compute函数计算字段值
   * 计算函数可访问this._context提供的外部上下文（如品牌配置、利润计算器）
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
          // 为计算函数提供上下文
          item[key] = computeFn(item, this._context);
        } catch (e) {
          // 计算失败时不设置值
        }
      });
    });
  }

  /**
   * 应用字段默认值
   * @private
   * @param {Object} item - 数据项
   * @param {Object} entityConfig - 实体配置
   * @description
   * 为未定义的字段设置默认值：
   * - 默认值可以是静态值或函数
   * - 函数接收当前数据项作为参数
   */
  _applyDefaultValues(item, entityConfig) {
    Object.entries(entityConfig.fields).forEach(([fieldName, fieldConfig]) => {
      // 如果字段没有值且有默认值配置
      if (item[fieldName] === undefined && fieldConfig.default !== undefined) {
        // 如果默认值是函数，则调用它
        if (typeof fieldConfig.default === "function") {
          item[fieldName] = fieldConfig.default(item);
        } else {
          item[fieldName] = fieldConfig.default;
        }
      }
    });
  }

  /**
   * 判断两条记录是否相同（基于唯一键）
   * @private
   * @param {Object} record1 - 第一条记录
   * @param {Object} record2 - 第二条记录
   * @param {Object} entityConfig - 实体配置
   * @returns {boolean} 是否相同
   */
  _isSameRecord(record1, record2, entityConfig) {
    if (entityConfig.uniqueKey) {
      const uniqueKeyConfig = this._config.parseUniqueKey(
        entityConfig.uniqueKey,
      );
      return uniqueKeyConfig.fields.every(
        (field) => record1[field] === record2[field],
      );
    }
  }

  /**
   * 注册实体索引配置
   * @param {string} entityName - 实体名称
   * @param {Array<Object>} indexConfigs - 索引配置数组
   * @param {string[]} indexConfigs[].fields - 索引字段列表
   * @param {boolean} [indexConfigs[].unique] - 是否唯一索引
   *
   * @example
   * repository.registerIndexes("Product", [
   *   { fields: ["itemNumber"], unique: true },  // 主键索引
   *   { fields: ["brand", "year"] }              // 普通复合索引
   * ]);
   */
  registerIndexes(entityName, indexConfigs) {
    this._indexConfigs.set(entityName, indexConfigs || []);

    if (this._cache.has(entityName)) {
      const data = this._cache.get(entityName);
      this._buildAllIndexes(entityName, data);
    }
  }

  /**
   * 设置上下文环境
   * @param {Object} context - 上下文对象
   * @param {Object} [context.brandConfig] - 品牌配置映射
   * @param {Object} [context.profitCalculator] - 利润计算器
   * @description 为计算字段提供外部依赖
   */
  setContext(context) {
    Object.assign(this._context, context);
  }

  /**
   * 获取品牌配置映射
   * @returns {Object.<string, Object>} 品牌SN到品牌配置的映射
   * @throws {Error} 读取品牌配置失败时抛出
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
   * 刷新品牌配置缓存
   * @returns {Object.<string, Object>} 新的品牌配置映射
   */
  refreshBrandConfig() {
    this._context.brandConfig = null;
    this._cache.delete("BrandConfig");
    this._indexes.delete("BrandConfig");
    return this.getBrandConfigMap();
  }

  /**
   * 查询实体的所有数据
   * @param {string} entityName - 实体名称
   * @returns {Object[]} 实体数据数组
   * @throws {Error} 实体不存在时抛出
   * @description
   * 查询策略：
   * - 优先从缓存读取
   * - 缓存不存在时从Excel读取
   * - 读取后自动建立索引和计算字段
   */
  findAll(entityName) {
    if (this._cache.has(entityName)) {
      return this._cache.get(entityName);
    }

    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    const data = this._excelDAO.read(entityName);

    // 计算计算字段
    this._computeFields(data, entityConfig);
    this._cache.set(entityName, data);
    this._buildAllIndexes(entityName, data);

    return data;
  }

  /**
   * 查询符合条件的实体对象
   * @param {string} entityName - 实体名称
   * @param {Object} [condition] - 查询条件
   * @returns {Object[]} 符合条件的记录数组
   * @description
   * 查询策略：
   * - 优先通过索引查询（性能最优）
   * - 无可用索引时执行全表扫描
   *
   * @example
   * // 单条件查询
   * repository.find("Product", { status: "上线" });
   *
   * // 多条件查询（AND关系）
   * repository.find("Product", { brand: "A", year: 2024 });
   *
   * // 无条件返回所有
   * repository.find("Product");
   */
  find(entityName, condition) {
    const data = this.findAll(entityName);

    if (!condition || Object.keys(condition).length === 0) {
      return data;
    }

    // 优先通过索引查询数据
    const indexedResult = this._queryByIndex(entityName, condition);
    if (indexedResult != null) {
      return indexedResult;
    }

    // 索引无数据遍历全表
    return this._fullScan(data, condition);
  }

  /**
   * 复杂查询（支持过滤、排序、分页）
   * @param {string} entityName - 实体名称
   * @param {Object} [options] - 查询选项
   * @param {Object|Function} [options.filter] - 过滤条件
   * @param {Object|Array} [options.sort] - 排序规则
   * @param {number} [options.limit] - 限制返回数量
   * @param {number} [options.offset] - 偏移量（用于分页）
   * @returns {Object[]} 查询结果
   *
   * @description
   * 支持的过滤操作符：
   * - $in: 包含于数组
   * - $between: 介于两值之间
   * - $gt: 大于
   * - $gte: 大于等于
   * - $lt: 小于
   * - $lte: 小于等于
   * - $ne: 不等于
   * - $like: 模糊匹配（支持%通配符）
   * - 函数: 自定义过滤函数
   *
   * @example
   * // 复杂查询示例
   * repository.query("Product", {
   *   filter: {
   *     price: { $gt: 100, $lt: 500 },
   *     status: { $in: ["上线", "预售"] },
   *     name: { $like: "%T恤%" }
   *   },
   *   sort: [
   *     { field: "price", order: "desc" },
   *     { field: "name", order: "asc" }
   *   ],
   *   limit: 20,
   *   offset: 40
   * });
   */
  query(entityName, options = {}) {
    let results = this.findAll(entityName);

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
          } else if (field.includes("Date") || field.includes("Time")) {
            const dateA = Date.parse(valA) || 0;
            const dateB = Date.parse(valB) || 0;
            if (dateA !== dateB) return (dateA - dateB) * order;
          } else {
            valA = String(valA || "");
            valB = String(valB || "");
            const compare = valA.localeCompare(valB);
            if (compare !== 0) return compare * order;
          }
        }
        return 0;
      });
    }

    if (options.limit) {
      const start = options.offset || 0;
      results = results.slice(start, start + options.limit);
    }

    return results;
  }

  /**
   * 保存实体数据到工作表
   * @param {string} entityName - 实体名称
   * @param {Object[]} data - 要保存的数据数组
   * @returns {Object[]} 保存后的数据
   * @throws {Error} 验证失败时抛出详细错误信息
   * @description
   * 保存流程：
   * 1. 执行数据验证
   * 2. 重新计算计算字段
   * 3. 按默认排序规则排序
   * 4. 写入Excel并更新缓存
   * 5. 重建索引
   */
  save(entityName, data) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    // 验证数据
    const validationResult = this._validationEngine.validateAll(
      data,
      entityConfig,
    );

    if (!validationResult.valid) {
      const errorMsg = this._validationEngine.formatErrors(
        validationResult,
        entityConfig.worksheet,
      );
      throw new Error(errorMsg);
    }

    // 计算计算字段
    this._computeFields(data, entityConfig);

    // 按默认排序规则排序
    if (entityConfig?.defaultSort) {
      data.sort(entityConfig.defaultSort);
    }

    this._excelDAO.write(entityName, data);
    this._cache.set(entityName, data);
    this._buildAllIndexes(entityName, data);

    return data;
  }

  /**
   * 清空实体所有数据
   * @param {string} entityName - 实体名称
   */
  clear(entityName) {
    this._excelDAO.clear(entityName);
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
  }

  /**
   * 清空所有缓存和索引
   */
  clearAllCache() {
    this._cache.clear();
    this._indexes.clear();
  }

  /**
   * 刷新实体缓存和索引
   * @param {string} entityName - 实体名称
   * @returns {Object[]} 刷新后的数据
   * @description 强制从Excel重新读取数据，更新缓存和索引
   */
  refresh(entityName) {
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
    return this.findAll(entityName);
  }

  /**
   * 批量保存多个实体数据（事务性）
   * @param {Object.<string, Object[]>} operations - 操作对象，键为实体名，值为数据数组
   * @returns {Object.<string, Object[]>} 各实体保存结果
   * @throws {Error} 任一实体保存失败时抛出汇总错误
   *
   * @example
   * repository.transaction({
   *   Product: [{ name: "产品1" }, { name: "产品2" }],
   *   Inventory: [{ productCode: "001", quantity: 100 }]
   * });
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

  // ==================== 数据修改操作 ====================

  /**
   * 新增单条记录
   * @param {string} entityName - 实体名称
   * @param {Object} item - 要新增的数据项
   * @param {Object} [options] - 选项
   * @param {boolean} [options.validateOnly=false] - 仅验证不保存
   * @returns {Object|boolean} 验证通过返回true，保存返回操作结果 { insert: [item] }
   * @throws {Error} 验证失败时抛出
   * @description
   * 新增流程：
   * - 自动生成行号(_rowNumber)
   * - 应用字段默认值
   * - 执行数据验证（包括唯一性验证）
   */
  add(entityName, item, options = {}) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知业务实体：${entityName}`);
    }

    // 获取当前数据
    const currentData = this.findAll(entityName);

    // 自动生成行号
    const maxRowNumber = currentData.reduce(
      (max, item) => Math.max(max, item._rowNumber || 0),
      0,
    );
    item._rowNumber = maxRowNumber + 1;

    // 处理默认值
    this._applyDefaultValues(item, entityConfig);

    // 验证单条数据
    const validationResult = this._validationEngine.validateEntity(
      item,
      entityConfig,
      { allData: currentData }, // 传入现有数据用于唯一性验证
    );

    if (!validationResult.valid) {
      const errorMsg = this._validationEngine.formatErrors(
        { items: [{ ...validationResult }] },
        entityConfig.worksheet,
      );
      throw new Error(errorMsg);
    }

    if (options.validateOnly) {
      return true;
    }

    // 添加到数据集中
    currentData.push(item);
    // 保存所有数据
    this.save(entityName, currentData);

    return { insert: [item] };
  }

  /**
   * 批量新增多条记录
   * @param {string} entityName - 实体名称
   * @param {Object[]} items - 要新增的数据项数组
   * @param {Object} [options] - 选项
   * @param {boolean} [options.validateOnly=false] - 仅验证不保存
   * @returns {Object|boolean} 验证通过返回true，保存返回操作结果 { insert: items }
   * @throws {Error} 任一记录验证失败时抛出汇总错误
   */
  addMany(entityName, items, options = {}) {
    if (!Array.isArray(items)) {
      throw new Error("items 必须是数组");
    }

    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    const currentData = this.findAll(entityName);
    const maxRowNumber = currentData.reduce(
      (max, item) => Math.max(max, item._rowNumber || 0),
      0,
    );

    const errors = [];

    items.forEach((newItem, index) => {
      try {
        // 生成行号
        newItem._rowNumber = maxRowNumber + index + 1;

        // 处理默认值
        this._applyDefaultValues(newItem, entityConfig);

        // 提前验证数据
        const validationResult = this._validationEngine.validateEntity(
          newItem,
          entityConfig,
          { allData: [...currentData, ...newItems] }, // 包括之前验证通过的新记录
        );

        if (!validationResult.valid) {
          throw new Error(
            this._validationEngine.formatErrors(
              { items: [{ ...validationResult }] },
              entityConfig.worksheet,
            ),
          );
        }
      } catch (e) {
        errors.push(`第${index + 1}条记录：${e.message}`);
      }
    });

    if (errors.length > 0) {
      throw new Error(`批量新增失败：\n${errors.join("\n")}`);
    }

    if (options.validateOnly) {
      return true;
    }

    // 合并并保存所有数据到缓存
    const updatedData = [...currentData, ...items];
    this.save(entityName, updatedData);

    return { insert: items };
  }

  /**
   * 更新符合条件的记录
   * @param {string} entityName - 实体名称
   * @param {Object} condition - 查询条件
   * @param {Object} updates - 更新内容
   * @param {Object} [options] - 选项
   * @param {boolean} [options.upsert=false] - 不存在时是否插入
   * @param {boolean} [options.multi=false] - 是否允许更新多条
   * @param {boolean} [options.validateOnly=false] - 仅验证不保存
   * @returns {Object|boolean} 验证通过返回true，保存返回操作结果 { insert: [], update: [] }
   * @throws {Error} 验证失败或条件不满足时抛出
   *
   * @example
   * // 更新单条
   * repository.update("Product", { itemNumber: "A001" }, { price: 199 });
   *
   * // 更新多条（需显式允许）
   * repository.update("Product", { brand: "A" }, { discount: 0.8 }, { multi: true });
   *
   * // 不存在则插入
   * repository.update("Product", { itemNumber: "B001" }, { name: "新品" }, { upsert: true });
   */
  update(entityName, condition, updates, options = {}) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    const result = { insert: [], update: [] };

    // 查找要更新的记录
    const records = this.find(entityName, condition);

    if (records.length === 0) {
      if (options.upsert) {
        // 如果没找到且 upsert 为 true，则新增
        const newItem = { ...condition, ...updates };

        if (this.add(entityName, newItem, options) === true) {
          return true;
        } else {
          result.insert.push(newItem);
          return result;
        }
      }
      throw new Error(`未找到符合条件的记录`);
    }

    if (records.length > 1 && !options.multi) {
      throw new Error(`找到多条记录，请使用更精确的条件或设置 { multi: true }`);
    }

    const currentData = this.findAll(entityName);
    const updatedRecords = [];
    const errors = [];

    records.forEach((record) => {
      try {
        // 找到原记录在数组中的索引
        const index = currentData.findIndex((item) =>
          this._isSameRecord(item, record, entityConfig),
        );

        if (index === -1) {
          throw new Error("记录状态异常，请刷新后重试");
        }

        // 创建更新后的记录
        const updatedRecord = {
          ...currentData[index],
          ...updates,
          _rowNumber: currentData[index]._rowNumber, // 保持行号不变
        };

        // 验证更新后的数据
        const validationResult = this._validationEngine.validateEntity(
          updatedRecord,
          entityConfig,
          { allData: currentData.filter((_, i) => i !== index) }, // 排除被更新的业务对象自身
        );

        if (!validationResult.valid) {
          throw new Error(
            this._validationEngine.formatErrors(
              { items: [{ ...validationResult }] },
              entityConfig.worksheet,
            ),
          );
        }

        updatedRecords.push({ index: updatedRecord });
      } catch (e) {
        errors.push(e.message);
      }
    });

    if (errors.length > 0) {
      throw new Error(`更新失败：\n${errors.join("\n")}`);
    }

    if (options.validateOnly) {
      return true;
    }

    // 替换并保存所有数据到缓存
    for (const [index, updatedRecord] of Object.entries(updatedRecords)) {
      currentData[Number(index)] = updatedRecord;
      result.update.push(updatedRecord);
    }
    this.save(entityName, currentData);

    return result;
  }

  /**
   * 更新多条符合条件的已存在的记录
   * @param {string} entityName - 实体名称
   * @param {Object} condition - 查询条件
   * @param {Object} updates - 更新内容
   * @returns {Object} 操作结果
   * @see update
   */
  updateMany(entityName, condition, updates) {
    return this.update(entityName, condition, updates, { multi: true });
  }

  /**
   * 插入或更新单条记录（存在则更新，不存在则插入）
   * @param {string} entityName - 实体名称
   * @param {Object} item - 数据项
   * @param {string[]} [uniqueFields] - 用于判断唯一性的字段，默认使用实体唯一键
   * @returns {Object} 操作结果 { insert: [], update: [] }
   * @throws {Error} 唯一字段缺失或找到多条记录时抛出
   *
   * @example
   * // 使用默认唯一键
   * repository.upsert("Product", { itemNumber: "A001", name: "产品A", price: 100 });
   *
   * // 指定唯一字段
   * repository.upsert("Product", { code: "001", name: "产品B" }, ["code"]);
   */
  upsert(entityName, item, uniqueFields) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    // 如果没有指定唯一字段，使用实体的唯一键
    if (!uniqueFields && entityConfig.uniqueKey) {
      const uniqueKeyConfig = this._config.parseUniqueKey(
        entityConfig.uniqueKey,
      );
      uniqueFields = uniqueKeyConfig.fields;
    }

    if (!uniqueFields || uniqueFields.length === 0) {
      throw new Error("请指定用于判断唯一性的字段");
    }

    // 构建查询条件
    const condition = {};
    uniqueFields.forEach((field) => {
      if (item[field] == undefined) {
        throw new Error(`唯一字段 ${field} 在数据中不存在`);
      }
      condition[field] = item[field];
    });

    // 查找是否存在
    const existing = this.find(entityName, condition);

    if (existing.length > 1) {
      throw new Error(`找到多条记录，无法确定更新哪一条`);
    }

    const result = { insert: [], update: [] };

    if (existing.length === 1) {
      // 更新
      this.update(entityName, condition, item);
      result.update.push(item);
    } else {
      // 新增
      this.add(entityName, item);
      result.insert.push(item);
    }

    return result;
  }

  /**
   * 删除符合条件的记录
   * @param {string} entityName - 实体名称
   * @param {Object} condition - 查询条件
   * @param {Object} [options] - 选项
   * @param {boolean} [options.multi=false] - 是否允许删除多条
   * @returns {number} 删除的记录数
   * @throws {Error} 找到多条但未设置multi时抛出
   */
  delete(entityName, condition, options = {}) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    // 查找要删除的记录
    const recordsToDelete = this.find(entityName, condition);

    if (recordsToDelete.length === 0) {
      return 0;
    }

    if (recordsToDelete.length > 1 && !options.multi) {
      throw new Error(`找到多条记录，请使用更精确的条件或设置 { multi: true }`);
    }

    const currentData = this.findAll(entityName);

    // 过滤掉要删除的记录
    const newData = currentData.filter((item) => {
      const shouldDelete = recordsToDelete.some((record) =>
        this._isSameRecord(item, record, entityConfig),
      );
      return !shouldDelete;
    });

    if (newData.length === currentData.length) {
      return 0;
    }

    // 保存所有数据
    this.save(entityName, newData);

    return currentData.length - newData.length;
  }

  // ==================== 快捷查询方法 ====================

  /**
   * 获取所有（符合条件）产品
   * @param {Object} [query] - 查询条件
   * @returns {Object[]} 产品数组
   */
  findProducts(query = {}) {
    return this.find("Product", query);
  }

  /**
   * 通过货号获取产品
   * @param {string} itemNumber - 货号
   * @returns {Object|undefined} 产品对象
   */
  findProductByItemNumber(itemNumber) {
    return this.find("Product", { itemNumber })[0];
  }

  /**
   * 通过款号获取产品
   * @param {string} styleNumber - 款号
   * @returns {Object[]} 产品数组
   */
  findProductsByStyle(styleNumber) {
    return this.find("Product", { styleNumber });
  }

  /**
   * 通过商品状态获取产品
   * @param {string} itemStatus - 商品状态
   * @returns {Object[]} 产品数组
   */
  findProductsByStatus(itemStatus) {
    return this.find("Product", { itemStatus });
  }

  /**
   * 获取低库存产品（可售库存低于阈值且处于上线状态）
   * @param {number} [threshold=30] - 库存阈值
   * @returns {Object[]} 低库存产品数组
   */
  findLowStockProducts(threshold = 30) {
    return this.query("Product", {
      filter: {
        sellableInventory: { $lt: threshold },
        itemStatus: { $in: ["商品上线", "部分上线"] },
      },
      sort: { field: "sellableInventory", order: "asc" },
    });
  }

  /**
   * 获取低毛利产品
   * @param {number} [minProfit=5] - 最小毛利阈值
   * @returns {Object[]} 低毛利产品数组
   */
  findLowProfitProducts(minProfit = 5) {
    return this.query("Product", {
      filter: (item) => (item.profit || 0) < minProfit,
      sort: [{ field: "profit", order: "desc" }],
    });
  }

  /**
   * 获取低毛利率产品
   * @param {number} [minRate=0.35] - 最小毛利率阈值
   * @returns {Object[]} 低毛利率产品数组
   */
  findLowProfitRateProducts(minRate = 0.35) {
    return this.query("Product", {
      filter: (item) => (item.profitRate || 0) < minRate,
      sort: [{ field: "profitRate", order: "desc" }],
    });
  }

  /**
   * 获取所有（符合条件）的商品价格
   * @param {Object} [query] - 查询条件
   * @returns {Object[]} 商品价格数组
   */
  findProductPrices(query = {}) {
    return this.find("ProductPrice", query);
  }

  /**
   * 通过货号获取产品价格
   * @param {string} itemNumber - 货号
   * @returns {Object|undefined} 产品价格对象
   */
  findPriceByItemNumber(itemNumber) {
    return this.find("ProductPrice", { itemNumber })[0];
  }

  /**
   * 获取所有（符合条件）的常态商品
   * @param {Object} [query] - 查询条件
   * @returns {Object[]} 常态商品数组
   */
  findRegularProducts(query = {}) {
    return this.find("RegularProduct", query);
  }

  /**
   * 通过货号获取常态商品
   * @param {string} itemNumber - 货号
   * @returns {Object[]} 常态商品数组
   */
  findRegularProductsByItemNumber(itemNumber) {
    return this.find("RegularProduct", { itemNumber });
  }

  /**
   * 通过条码获取库存
   * @param {string} productCode - 商品条码
   * @returns {Object|undefined} 库存对象
   */
  findInventory(productCode) {
    return this.find("Inventory", { productCode })[0];
  }

  /**
   * 通过条码获取组合商品
   * @param {string} productCode - 商品条码
   * @returns {Object[]} 组合商品数组
   */
  findComboProducts(productCode) {
    return this.find("ComboProduct", { productCode });
  }

  /**
   * 获取指定货号在某个年份的销售记录
   * @param {string} itemNumber - 货号
   * @param {number|string} year - 年份
   * @returns {Object[]} 销售记录数组
   */
  findSalesByItemAndYear(itemNumber, year) {
    return this.find("ProductSales", {
      itemNumber,
      salesYear: year,
    });
  }

  /**
   * 获取指定货号在某个年月的销售记录
   * @param {string} itemNumber - 货号
   * @param {number|string} year - 年份
   * @param {number|string} month - 月份（1-12）
   * @returns {Object[]} 销售记录数组
   */
  findSalesByItemAndYearMonth(itemNumber, year, month) {
    const monthStr = String(month).padStart(2, "0");
    const yearMonth = `${year}-${monthStr}`;

    return this.find("ProductSales", {
      itemNumber,
      yearMonth,
    });
  }

  /**
   * 获取指定货号在某个年周的销售记录
   * @param {string} itemNumber - 货号
   * @param {number|string} year - 年份
   * @param {number|string} week - 周数
   * @returns {Object[]} 销售记录数组
   */
  findSalesByItemAndYearWeek(itemNumber, year, week) {
    const weekStr = String(week).padStart(2, "0");
    const yearWeek = `${year}-${weekStr}`;

    return this.find("ProductSales", {
      itemNumber,
      yearWeek,
    });
  }

  /**
   * 获取指定货号在指定日期的销售记录
   * @param {string} itemNumber - 货号
   * @param {Date|string} date - 日期
   * @returns {Object|undefined} 销售记录对象
   */
  findSalesByItemAndDate(itemNumber, date) {
    const dateStr = this._converter.toDateStr(date);

    return this.find("ProductSales", { itemNumber, salesDate: dateStr })[0];
  }

  /**
   * 获取指定货号最近N天内的商品销售记录
   * @param {string} itemNumber - 货号
   * @param {number} days - 天数
   * @returns {Object[]} 销售记录数组（按销售日期倒序）
   */
  findSalesLastNDays(itemNumber, days) {
    return this.query("ProductSales", {
      filter: {
        itemNumber,
        daysSinceSale: { $lte: days },
      },
      sort: { field: "salesDate", order: "desc" },
    });
  }

  /**
   * 获取指定时间段的商品销售记录
   * @param {Date|string} startDate - 开始日期
   * @param {Date|string} endDate - 结束日期
   * @returns {Object[]} 销售记录数组（按销售日期正序）
   */
  findSalesByDateRange(startDate, endDate) {
    const start = Date.parse(startDate) || 0;
    const end = Date.parse(endDate) || Infinity;

    return this.query("ProductSales", {
      filter: {
        salesDate: (value) => {
          const dateTs = Date.parse(value) || 0;
          return dateTs >= start && dateTs <= end;
        },
      },
      sort: { field: "salesDate", order: "asc" },
    });
  }

  /**
   * 获取系统记录
   * @returns {Object} 系统记录对象（单条）
   * @description 如果不存在系统记录，则创建一个空记录并缓存
   */
  getSystemRecord() {
    const records = this.findAll("SystemRecord");
    if (records.length > 0) {
      return records[0];
    }

    const entityConfig = this._config.get("SystemRecord");

    const fields = entityConfig.fields;
    const newRecord = {};

    Object.keys(fields).forEach((f) => (newRecord[f] = undefined));

    this._cache.set("SystemRecord", [newRecord]);
    return newRecord;
  }

  /**
   * 更新系统记录中的日期字段
   * @param {string} entityName - 实体名称
   * @param {string} dateField - 日期字段类型（'importDate' 或 'updateDate'）
   * @description
   * - 仅对可导入/可更新的实体生效
   * - 根据实体配置更新对应的系统字段为当前时间
   */
  updateSystemRecord(entityName, dateField) {
    if (
      !this._importableEntities.includes(entityName) &&
      dateField === "importDate"
    )
      return;
    if (
      !this._updatableEntities.includes(entityName) &&
      dateField === "updateDate"
    )
      return;

    const entityConfig = this._config.get(entityName);
    const updateField = entityConfig?.[dateField];
    if (!updateField) return;

    const systemRecord = this.getSystemRecord();
    const now = new Date();

    systemRecord[updateField] = now;

    this.save("SystemRecord", [systemRecord]);
  }
}
