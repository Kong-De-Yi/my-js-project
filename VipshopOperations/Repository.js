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

  // 以字符串形式返回字段的组合键值，多字段用 ¦ 连接
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

  // 通过索引查询数据
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

  // 遍历全表数据
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

  // 计算业务实体的计算字段
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

  // 应用字段默认值
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

  // 判断两条记录是否相同（基于唯一键）
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

  // 注册实体对象的索引
  registerIndexes(entityName, indexConfigs) {
    this._indexConfigs.set(entityName, indexConfigs || []);

    if (this._cache.has(entityName)) {
      const data = this._cache.get(entityName);
      this._buildAllIndexes(entityName, data);
    }
  }

  // 设置上下文环境
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

  // 从缓存中查询业务实体的所有数据，没有则从工作表读取并建立缓存和索引
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

  // 查询符合条件的实体对象（优先索引，其次遍历全表）
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

  // 复杂查询（遍历全表查询，支持条件过滤，排序，分页）
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

  // 保存业务实体数据到工作表并更新缓存和索引
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

    this._excelDAO.write(entityName, data);
    this._cache.set(entityName, data);
    this._buildAllIndexes(entityName, data);

    return data;
  }

  // 删除整个业务实体数据
  clear(entityName) {
    this._excelDAO.clear(entityName);
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
  }

  // 清空所有缓存和索引数据
  clearAllCache() {
    this._cache.clear();
    this._indexes.clear();
  }

  // 刷新缓存和索引中的业务实体数据
  refresh(entityName) {
    this._cache.delete(entityName);
    this._indexes.delete(entityName);
    return this.findAll(entityName);
  }

  // 批量保存多个不同业务实体数据
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

  // 新增一条记录，options={validateOnly:true}
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
      return item;
    }

    // 添加到数据集中
    currentData.push(item);
    // 保存所有数据
    this.save(entityName, currentData);

    return item;
  }

  // 批量新增多条记录，options={validateOnly:true}
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

    const newItems = [];
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

        newItems.push(newItem);
      } catch (e) {
        errors.push(`第${index + 1}条记录：${e.message}`);
      }
    });

    if (errors.length > 0) {
      throw new Error(`批量新增失败：\n${errors.join("\n")}`);
    }

    if (options.validateOnly) {
      return newItems;
    }

    // 合并并保存
    const updatedData = [...currentData, ...newItems];
    this.save(entityName, updatedData);

    return newItems;
  }

  // 更新一条或多条符合条件的记录，options={upsert:true,validateOnly:true,multi: true}
  update(entityName, condition, updates, options = {}) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    // 查找要更新的记录
    const records = this.find(entityName, condition);

    if (records.length === 0) {
      if (options.upsert) {
        // 如果没找到且 upsert 为 true，则新增
        const newItem = { ...condition, ...updates };
        return this.add(entityName, newItem);
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
              {
                items: [
                  { ...validationResult, rowNumber: updatedRecord._rowNumber },
                ],
              },
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
      return Object.values(updatedRecords);
    }

    // 替换并保存
    for (const [index, updatedRecord] of Object.entries(updatedRecords)) {
      currentData[Number(index)] = updatedRecord;
    }
    this.save(entityName, currentData);

    return Object.values(updatedRecords);
  }

  // 更新多条符合条件的已存在的记录
  updateMany(entityName, condition, updates) {
    return this.update(entityName, condition, updates, { multi: true });
  }

  // 插入或更新一条记录（存在则更新，不存在则插入,通过主键或指定字段确认唯一性）
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

    const result = null;
    if (existing.length === 1) {
      // 更新
      result = this.update(entityName, condition, item);
    } else {
      // 新增
      result = this.add(entityName, item);
    }

    return Array.isArray(result) ? result[0] : result;
  }

  // 删除一条或多条符合条件的记录,options={multi:true}
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

    // 保存
    this.save(entityName, newData);

    return currentData.length - newData.length;
  }

  // ==================== 快捷查询方法 ====================

  // 获取所有（符合条件）产品
  findProducts(query = {}) {
    return this.find("Product", query);
  }
  // 通过货号获取产品
  findProductByItemNumber(itemNumber) {
    return this.find("Product", { itemNumber });
  }
  // 通过款号获取产品
  findProductsByStyle(styleNumber) {
    return this.find("Product", { styleNumber });
  }
  // 通过商品状态获取产品
  findProductsByStatus(itemStatus) {
    return this.find("Product", { itemStatus });
  }
  // 获取上线产品可售库存低于指定数量的产品
  findLowStockProducts(threshold = 30) {
    return this.query("Product", {
      filter: {
        sellableInventory: { $lt: threshold },
        itemStatus: { $in: ["商品上线", "部分上线"] },
      },
      sort: { field: "sellableInventory", order: "asc" },
    });
  }
  // 获取低毛利产品
  findLowProfitProducts(minProfit = 5) {
    return this.query("Product", {
      filter: (item) => (item.profit || 0) < minProfit,
      sort: [{ field: "profit", order: "desc" }],
    });
  }
  // 获取低毛利率产品
  findLowProfitRateProducts(minRate = 0.35) {
    return this.query("Product", {
      filter: (item) => (item.profitRate || 0) < minRate,
      sort: [{ field: "profitRate", order: "desc" }],
    });
  }

  // 通过货号获取产品价格
  findPriceByItemNumber(itemNumber) {
    return this.findOne("ProductPrice", { itemNumber });
  }

  // 获取所有（符合条件）的常态商品
  findRegularProducts(query = {}) {
    return this.find("RegularProduct", query);
  }

  // 通过条码获取库存
  findInventory(productCode) {
    return this.find("Inventory", { productCode });
  }

  // 通过条码获取组合商品
  findComboProducts(productCode) {
    return this.find("ComboProduct", { productCode });
  }

  // 获取所有（符合条件）的商品销售
  findProductSales(query = {}) {
    return this.find("ProductSales", query);
  }
  // 获取指定年份的商品销售
  findSalesByYear(year) {
    return this.find("ProductSales", { salesYear: year });
  }
  // 获取指定年月的商品销售
  findSalesByYearMonth(yearMonth) {
    return this.find("ProductSales", { yearMonth: yearMonth });
  }
  // 获取指定货号在某个年份的销售
  findSalesByItemAndYear(itemNumber, year) {
    return this.find("ProductSales", {
      itemNumber: itemNumber,
      salesYear: year,
    });
  }
  // 获取指定货号在某个年月的销售
  findSalesByItemAndYearMonth(itemNumber, yearMonth) {
    return this.find("ProductSales", {
      itemNumber: itemNumber,
      yearMonth: yearMonth,
    });
  }
  // 获取最近N天的商品销售
  findSalesLastNDays(days) {
    return this.query("ProductSales", {
      filter: {
        daysSinceSale: { $lte: days },
      },
      sort: { field: "salesDate", order: "asc" },
    });
  }
  // 获取指定时间段的商品销售
  findSalesByDateRange(startDate, endDate) {
    const start = this._excelDAO.parseDate(startDate)?.getTime() || 0;
    const end = this._excelDAO.parseDate(endDate)?.getTime() || Infinity;

    return this.query("ProductSales", {
      filter: (item) => {
        const date = this._excelDAO.parseDate(item.salesDate)?.getTime() || 0;
        return date >= start && date <= end;
      },
      sort: { field: "salesDate", order: "asc" },
    });
  }

  // 获取系统记录
  getSystemRecord() {
    const records = this.findAll("SystemRecord");
    if (records.length > 0) {
      return records[0];
    }

    const newRecord = {
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
}
