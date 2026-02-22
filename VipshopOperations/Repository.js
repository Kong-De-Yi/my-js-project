// ============================================================================
// Repository.js - 数据仓库（增强版）
// 功能：支持索引字段的自动计算和存储
// ============================================================================

class Repository {
  constructor(excelDAO) {
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();
    this._validationEngine = ValidationEngine.getInstance();

    this._cache = new Map();
    this._indexes = new Map();
    this._indexConfigs = new Map();

    this._context = {
      brandConfig: null,
      profitCalculator: null,
    };
  }

  // 根据_indexConfigs建立业务实体的索引
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

    indexConfigs.forEach((config) => {
      const sortedFields = [...config.fields].sort();
      const indexKey = sortedFields.join("|");

      if (!entityIndexes.has(indexKey)) {
        entityIndexes.set(indexKey, new Map());
      }

      const index = entityIndexes.get(indexKey);
      index.clear();

      data.forEach((item) => {
        const value = this._getCompositeKey(item, sortedFields);

        if (value != undefined && String(value).trim() !== "") {
          if (config.unique) {
            if (index.has(value)) {
              item._indexError = `索引${indexKey}值"${value}"重复`;
            } else {
              index.set(value, item);
            }
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

  registerIndexes(entityName, indexConfigs) {
    this._indexConfigs.set(entityName, indexConfigs || []);

    if (this._cache.has(entityName)) {
      const data = this._cache.get(entityName);
      this._buildAllIndexes(entityName, data);
    }
  }

  setContext(context) {
    Object.assign(this._context, context);
  }

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

  refreshBrandConfig() {
    this._context.brandConfig = null;
    this._cache.delete("BrandConfig");
    this._indexes.delete("BrandConfig");
    return this.getBrandConfigMap();
  }

  findAll(entityName) {
    if (this._cache.has(entityName)) {
      return this._cache.get(entityName);
    }

    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    const data = this._excelDAO.read(entityName);

    // 计算计算字段（包括索引字段）
    this._computeFields(data, entityConfig);

    this._cache.set(entityName, data);
    this._buildAllIndexes(entityName, data);

    return data;
  }

  find(entityName, condition) {
    const data = this.findAll(entityName);

    if (!condition || Object.keys(condition).length === 0) {
      return data;
    }

    const indexedResult = this._queryByIndex(entityName, condition);
    if (indexedResult != null) {
      return indexedResult;
    }

    return this._fullScan(data, condition);
  }

  findOne(entityName, condition) {
    const entityConfig = this._config.get(entityName);

    if (entityConfig?.uniqueKey) {
      const uniqueKeyConfig = this._config.parseUniqueKey(
        entityConfig.uniqueKey,
      );
      const fields = uniqueKeyConfig.fields;

      const hasAllFields = fields.every((f) => condition[f] != undefined);

      if (hasAllFields && Object.keys(condition).length === fields.length) {
        const indexResult = this._queryByIndex(entityName, condition);
        if (indexResult && indexResult.length > 0) {
          return indexResult[0];
        }
        return null;
      }
    }

    const results = this.find(entityName, condition);
    return results.length > 0 ? results[0] : null;
  }

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

  // ==================== 快捷查询方法 ====================

  findProductByItemNumber(itemNumber) {
    return this.findOne("Product", { itemNumber });
  }

  findProducts(query = {}) {
    return this.find("Product", query);
  }

  findProductsByStyle(styleNumber) {
    return this.find("Product", { styleNumber });
  }

  findProductsByBrandAndStatus(brandSN, itemStatus) {
    return this.find("Product", { brandSN, itemStatus });
  }

  findProductsBySPU(P_SPU) {
    return this.find("Product", { P_SPU });
  }

  findPriceByItemNumber(itemNumber) {
    return this.findOne("ProductPrice", { itemNumber });
  }

  findRegularProducts(query = {}) {
    return this.find("RegularProduct", query);
  }

  findInventory(productCode) {
    return this.findOne("Inventory", { productCode });
  }

  findComboProducts(productCode) {
    return this.find("ComboProduct", { productCode });
  }

  findProductSales(query = {}) {
    return this.find("ProductSales", query);
  }

  findSalesByYear(year) {
    return this.find("ProductSales", { salesYear: year });
  }

  findSalesByYearMonth(yearMonth) {
    return this.find("ProductSales", { yearMonth: yearMonth });
  }

  findSalesByItemAndYear(itemNumber, year) {
    return this.find("ProductSales", {
      itemNumber: itemNumber,
      salesYear: year,
    });
  }

  findSalesByItemAndYearMonth(itemNumber, yearMonth) {
    return this.find("ProductSales", {
      itemNumber: itemNumber,
      yearMonth: yearMonth,
    });
  }

  findSalesLastNDays(days) {
    return this.query("ProductSales", {
      filter: {
        daysSinceSale: { $lte: days },
      },
      sort: { field: "salesDate", order: "asc" },
    });
  }

  getSystemRecord() {
    const records = this.findAll("SystemRecord");
    if (records.length > 0) {
      return records[0];
    }

    const newRecord = {
      recordId: "SYSTEM_RECORD_1",
      updateDateOfLast7Days: "",
      updateDateOfProductPrice: "",
      updateDateOfRegularProduct: "",
      updateDateOfInventory: "",
      updateDateOfProductSales: "",
      _rowNumber: 1,
    };

    this._cache.set("SystemRecord", [newRecord]);
    return newRecord;
  }

  findLowStockProducts(threshold = 30) {
    return this.query("Product", {
      filter: {
        sellableDays: { $lt: threshold },
        itemStatus: { $in: ["商品上线", "部分上线"] },
      },
      sort: { field: "sellableDays", order: "asc" },
    });
  }

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

    // 计算计算字段（包括索引字段）
    this._computeFields(data, entityConfig);

    this._excelDAO.write(entityName, data);

    this._cache.set(entityName, data);
    this._buildAllIndexes(entityName, data);

    return data;
  }

  clear(entityName) {
    this._excelDAO.clear(entityName);
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
  }

  refresh(entityName) {
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
    return this.findAll(entityName);
  }

  // 批量保存多个实体数据
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

  _parseCompositeKey(indexKey, compositeValue) {
    const fields = indexKey.split("|");
    const values = String(compositeValue).split("¦");

    const condition = {};
    fields.forEach((field, i) => {
      condition[field] = values[i] === "" ? undefined : values[i];
    });

    return condition;
  }

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
   * 计算计算字段（包括索引字段）
   * @param {Array} data - 数据数组
   * @param {Object} entityConfig - 实体配置
   * @private
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

  clearAllCache() {
    this._cache.clear();
    this._indexes.clear();
  }
}
