class DataImportService {
  constructor(repository, excelDAO) {
    this._repository = repository;
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();
    this._identifier = EntityIdentifier.getInstance();
    this._validationEngine = ValidationEngine.getInstance();
    this._importableEntities = this._identifier.getImportableEntities();
  }

  // 覆盖模式导入
  _overwriteData(entityName, items) {
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

  // 追加模式导入
  _appendData(entityName, newItems) {
    // 检查业务实体是否配置主键
    const entityConfig = this._config.get(entityName);
    const uniqueKeyConfig = this._config.parseUniqueKey(entityConfig.uniqueKey);
    const fields = uniqueKeyConfig.fields;

    if (fields.length === 0) {
      throw new Error("追加模式导入的业务实体必须配置主键");
    }

    // 提前验证新数据
    const validationResult = this._validationEngine.validateAll(
      newItems,
      entityConfig,
    );

    if (!validationResult.valid) {
      const errorMsg = this._validationEngine.formatErrors(
        validationResult,
        entityConfig.worksheet,
      );
      throw new Error(errorMsg);
    }

    // 读取历史数据
    let existingItems = [];
    existingItems = this._repository.findAll(entityName);

    // 建立索引：货号+日期
    const existingMap = new Map();
    existingItems.forEach((item) => {
      const key = fields.map((field) => item[field]).join("|");
      existingMap.set(key, item);
    });

    // 合并数据
    const mergedItems = [...existingItems];
    let updatedCount = 0;
    let newCount = 0;

    newItems.forEach((newItem) => {
      const key = fields.map((field) => newItem[field]).join("|");

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

    return {
      success: true,
      entityName: entityConfig.worksheet,
      mode: "append",
      total: newItems.length,
      new: newCount,
      updated: updatedCount,
      message: `【${entityConfig.worksheet}】导入完成：：新增${newCount}条，更新${updatedCount}条`,
    };
  }

  // 更新系统记录
  _updateSystemRecord(entityName) {
    if (!this._importableEntities.includes(entityName)) return;

    const systemRecord = this._repository.getSystemRecord();
    const now = new Date();

    switch (entityName) {
      case "RegularProduct":
        systemRecord.importDateOfRegularProduct = now;
        break;
      case "Inventory":
        systemRecord.importDateOfInventory = now;
        break;
      case "ComboProduct":
        systemRecord.importDateOfComboProduct = now;
        break;
      case "ProductSales":
        systemRecord.importDateOfProductSales = now;
    }

    this._repository.save("SystemRecord", [systemRecord]);
  }

  // 执行数据导入
  import() {
    // 1. 获取导入数据
    const wb = this._excelDAO.getWorkbook();
    const sheet = wb.Sheets("导入数据");

    if (!sheet) {
      throw new Error("未找到【导入数据】工作表");
    }

    const usedRange = sheet.UsedRange;
    if (!usedRange || usedRange.Value2 == null) {
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
      const requiredTitles = this._importableEntities
        .map((entityName) => {
          const entityConfig = this._config.get(entityName);
          return `【${entityConfig.worksheet}】: ${entityConfig.requiredTitles.toString()}`;
        })
        .join("\n");
      throw new Error(
        "无法识别导入数据的类型，请确保表头包含以下必填字段之一：\n" +
          requiredTitles,
      );
    }

    // 4. 验证是否支持导入
    if (!this._identifier.canImport(entityName)) {
      throw new Error(`实体【${entityName}】不支持导入操作`);
    }

    // 5. 获取导入模式
    const mode = this._identifier.getImportMode(entityName);

    // 6. 读取数据
    const items = this._excelDAO.read(entityName, "导入数据");

    // 7. 根据模式处理数据
    let result;
    if (mode === "append") {
      result = this._appendData(entityName, items);
    } else {
      result = this._overwriteData(entityName, items);
    }

    // 8.更新系统记录
    this._updateSystemRecord(entityName);

    // 9. 清空导入数据表
    this._excelDAO.clear("ImportData");

    return result;
  }
}
