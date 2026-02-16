// ============================================================================
// 数据导入服务
// 功能：从“导入数据”工作表读取数据，智能识别实体类型并导入
// 特点：销售数据追加模式，其他实体覆盖模式
// ============================================================================

class DataImportService {
  constructor(repository, excelDAO) {
    this._repository = repository;
    this._excelDAO = excelDAO;
    this._config = dataConfig;
    this._identifier = entityIdentifier;
  }

  /**
   * 执行数据导入
   * @returns {Object} 导入结果报告
   */
  import() {
    // 1. 获取导入数据
    const wb = this._excelDAO.getWorkbook();
    const sheet = wb.Sheets("导入数据");

    if (!sheet) {
      throw new Error("未找到【导入数据】工作表");
    }

    const usedRange = sheet.UsedRange;
    if (!usedRange || usedRange.Value2 === null) {
      throw new Error("【导入数据】工作表中没有数据");
    }

    const data = usedRange.Value2;
    if (data.length < 2) {
      throw new Error("【导入数据】工作表中只有标题行，没有数据");
    }

    // 2. 提取表头
    const headers = data[0].map((h) => String(h).trim());

    // 3. 识别实体类型
    const entityName = this._identifier.identify(headers);

    if (!entityName) {
      throw new Error(
        "无法识别导入数据的类型，请确保表头包含以下必填字段之一：\n" +
          "- 销售数据：日期、货号、曝光UV、商详UV、加购UV、客户数、拒退件数、销售量、销售额、首次上架时间\n" +
          "- 常态商品：条码、货号、款号、颜色、尺码、三级品类、品牌SN、尺码状态、商品状态、市场价、唯品价、到手价、可售库存、可售天数、商品ID、P_SPU\n" +
          "- 组合商品：组合商品实体编码、商品编码、数量\n" +
          "- 商品库存：商品编码、数量、进货仓库存、后整车间、超卖车间、备货车间、销退仓库存、采购在途数\n",
      );
    }

    // 4. 验证是否支持导入
    if (!this._identifier.canImport(entityName)) {
      throw new Error(`实体【${entityName}】不支持导入操作`);
    }

    // 5. 获取导入模式
    const mode = this._identifier.getImportMode(entityName);

    // 6. 执行导入
    const entityConfig = this._config.get(entityName);
    const fieldMapping = this._buildFieldMapping(headers, entityConfig);

    // 转换数据
    const items = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const item = { _rowNumber: i + 1 };

      Object.entries(fieldMapping).forEach(([field, colIndex]) => {
        const rawValue = row[colIndex];
        const fieldConfig = entityConfig.fields[field];

        if (fieldConfig?.type === "number") {
          item[field] = this._toNumber(rawValue);
        } else {
          item[field] = this._toString(rawValue);
        }
      });

      items.push(item);
    }

    // 7. 根据模式处理数据
    let result;
    if (mode === "append") {
      result = this._appendData(entityName, items);
    } else {
      result = this._overwriteData(entityName, items);
    }

    // 8. 清空导入数据表
    this._clearImportSheet();

    return result;
  }

  /**
   * 构建字段映射
   */
  _buildFieldMapping(headers, entityConfig) {
    const mapping = {};

    Object.entries(entityConfig.fields).forEach(([field, config]) => {
      const title = config.title || field;
      const index = headers.indexOf(title);
      if (index !== -1) {
        mapping[field] = index;
      }
    });

    // 检查必填字段
    const requiredFields = entityConfig.requiredFields || [];
    const missingFields = requiredFields.filter(
      (field) => !mapping.hasOwnProperty(field),
    );

    if (missingFields.length > 0) {
      const missingTitles = missingFields.map((f) => {
        const config = entityConfig.fields[f];
        return config?.title || f;
      });
      throw new Error(`缺少必填字段：${missingTitles.join("、")}`);
    }

    return mapping;
  }

  /**
   * 覆盖模式导入
   */
  _overwriteData(entityName, items) {
    // 验证数据（简单验证）
    const errors = [];
    items.forEach((item, index) => {
      if (!item._rowNumber) item._rowNumber = index + 2;
    });

    if (errors.length > 0) {
      throw new Error(`数据验证失败：\n${errors.join("\n")}`);
    }

    // 直接保存（覆盖）
    this._repository.save(entityName, items);

    const entityConfig = this._config.get(entityName);
    return {
      success: true,
      entityName: entityConfig.worksheet,
      mode: "overwrite",
      total: items.length,
      message: `成功导入 ${items.length} 条数据到【${entityConfig.worksheet}】`,
    };
  }

  /**
   * 追加模式导入（销售数据）
   */
  _appendData(entityName, newItems) {
    // 验证新数据
    const errors = [];
    newItems.forEach((item, index) => {
      if (!item._rowNumber) item._rowNumber = index + 2;

      // 检查必填
      if (!item.salesDate || !item.itemNumber) {
        errors.push(`第${item._rowNumber}行：缺少销售日期或货号`);
      }
    });

    if (errors.length > 0) {
      throw new Error(`数据验证失败：\n${errors.join("\n")}`);
    }

    // 读取历史数据
    let existingItems = [];
    try {
      existingItems = this._repository.findAll(entityName);
    } catch (e) {
      existingItems = [];
    }

    // 建立索引：货号+日期
    const existingMap = new Map();
    existingItems.forEach((item) => {
      if (item.itemNumber && item.salesDate) {
        const key = `${item.itemNumber}|${item.salesDate}`;
        existingMap.set(key, item);
      }
    });

    // 合并数据
    const mergedItems = [...existingItems];
    let updatedCount = 0;
    let newCount = 0;

    newItems.forEach((newItem) => {
      const key = `${newItem.itemNumber}|${newItem.salesDate}`;

      if (existingMap.has(key)) {
        // 更新
        const index = mergedItems.findIndex(
          (item) =>
            item.itemNumber === newItem.itemNumber &&
            item.salesDate === newItem.salesDate,
        );
        if (index !== -1) {
          mergedItems[index] = newItem;
          updatedCount++;
        }
      } else {
        // 新增
        mergedItems.push(newItem);
        newCount++;
      }
    });

    // 保存
    this._repository.save(entityName, mergedItems);

    const entityConfig = this._config.get(entityName);
    return {
      success: true,
      entityName: entityConfig.worksheet,
      mode: "append",
      total: newItems.length,
      new: newCount,
      updated: updatedCount,
      message: `销售数据导入完成：新增${newCount}条，更新${updatedCount}条`,
    };
  }

  /**
   * 清空导入数据表
   */
  _clearImportSheet() {
    try {
      const wb = this._excelDAO.getWorkbook();
      const sheet = wb.Sheets("导入数据");

      // 只清除数据，保留标题行
      const usedRange = sheet.UsedRange;
      if (usedRange && usedRange.Rows.Count > 1) {
        const dataRange = sheet.Range(
          sheet.Cells(2, 1),
          sheet.Cells(usedRange.Rows.Count, usedRange.Columns.Count),
        );
        dataRange.ClearContents();
      }

      wb.Save();
    } catch (e) {
      // 忽略清空错误
    }
  }

  /**
   * 转换为数字
   */
  _toNumber(value) {
    if (value === undefined || value === null || value === "") {
      return 0;
    }

    if (typeof value === "boolean") {
      return 0;
    }

    const num = Number(value);
    return isNaN(num) ? 0 : num;
  }

  /**
   * 转换为字符串
   */
  _toString(value) {
    if (value === undefined || value === null) {
      return "";
    }

    if (typeof value === "string") {
      return value;
    }

    return String(value);
  }
}
