class ReportEngine {
  static _instance = null;

  constructor(repository, excelDAO) {
    if (ReportEngine._instance) {
      return ReportEngine._instance;
    }

    this._repository = repository || Repository.getInstance();
    this._excelDAO = excelDAO || ExcelDAO.getInstance();
    this._config = DataConfig.getInstance();
    this._validationEngine = ValidationEngine.getInstance();

    this._templateManager = new ReportTemplateManager(
      this._repository,
      this._excelDAO,
    );

    // 初始化销售统计服务
    this._statisticsService = new StatisticsService(this._repository);
    this._statisticsFields = new StatisticsFields(this._statisticsService);

    // 缓存所有销售统计字段配置
    this._statisticsFieldsMap = this._buildStatisticsFieldsMap();

    this._context = {
      query: {},
      groupBy: null,
    };

    ReportEngine._instance = this;
  }

  // 单例模式
  static getInstance(repository, excelDAO) {
    if (!ReportEngine._instance) {
      ReportEngine._instance = new ReportEngine(repository, excelDAO);
    }
    return ReportEngine._instance;
  }

  // 构建统计字段映射（包含展开逻辑）
  _buildStatisticsFieldsMap() {
    const map = new Map();
    const baseFields = this._statisticsFields.getAllFields();

    // 展开所有可展开字段
    baseFields.forEach((field) => {
      if (field.type === "expandable") {
        const expanded = this._statisticsFields.expandField(field);
        expanded.forEach((expandedField) => {
          map.set(expandedField.field, expandedField);
        });
      } else {
        map.set(field.field, field);
      }
    });

    return map;
  }

  // 获取所有可用的字段（用于模板配置）
  getAvailableFields() {
    const baseFields = this._statisticsFields.getAllFields();

    return baseFields.map((field) => ({
      field: field.field,
      title: field.title,
      type: field.type,
      group: field.group,
      description: field.description,
    }));
  }

  // 设置上下文环境
  setContext(context) {
    Object.assign(this._context, context);
  }

  // 初始化模板
  initializeTemplates() {
    return this._templateManager.initializeDefaultTemplates();
  }

  // 获取模板列表
  getTemplateList() {
    return this._templateManager.getTemplateList();
  }

  // 设置当前模板
  setCurrentTemplate(templateName) {
    return this._templateManager.setCurrentTemplate(templateName);
  }

  // 获取当前模板
  getCurrentTemplate() {
    return this._templateManager.getCurrentTemplate();
  }

  // 刷新模板
  refreshTemplates() {
    this._templateManager.refresh();
    // 重新构建字段映射（因为时间变化）
    this._statisticsFieldsMap = this._buildStatisticsFieldsMap();
    return this._templateManager.loadTemplates();
  }

  //获取字段值（支持统计字段）
  _getFieldValue(product, column) {
    const field = column.field;

    if (this._statisticsFieldsMap.has(field)) {
      const statField = this._statisticsFieldsMap.get(field);
      try {
        return statField.compute(product);
      } catch (e) {
        console.log(`计算统计字段 ${field} 失败：`, e.message);
        return 0;
      }
    }

    return product[field];
  }

  // 生成报表
  generateReport() {
    // ----- 1. 获取当前模板的列配置 -----
    let columns = this._templateManager.getCurrentColumns();

    // ----- 2. 展开所有可展开字段 -----
    const expandedColumns = [];
    columns.forEach((col) => {
      // 获取原始字段定义（未展开的）
      const baseField = this._statisticsFields.getField(col.field);

      if (baseField?.type === "expandable") {
        // 展开抽象字段
        const expanded = this._statisticsFields.expandField(baseField);
        expanded.forEach((exp) => {
          expandedColumns.push({
            ...col,
            field: exp.field,
            title: exp.title,
            width: exp.width || col.width,
            format: exp.format || col.format,
            color: col.color,
          });
        });
      } else {
        // 普通字段或具体统计字段
        expandedColumns.push(col);
      }
    });

    // // ----- 3. 获取筛选条件 -----
    // const query = this._buildQueryFromUI();

    // // ----- 4. 获取分组字段 -----
    // let groupBy = null;
    // if (UserForm1.ComboBox4?.Value) {
    //   const groupMap = {
    //     上市年份: "listingYear",
    //     四级品类: "fourthLevelCategory",
    //     运营分类: "operationClassification",
    //     下线原因: "offlineReason",
    //     三级品类: "thirdLevelCategory",
    //   };
    //   groupBy = groupMap[UserForm1.ComboBox4.Value];
    // }

    // // ----- 5. 获取排序字段 -----
    // let sortField = null;
    // let sortAscending = true;
    // if (UserForm1.ComboBox6?.Value) {
    //   const sortMap = {
    //     首次上架时间: "firstListingTime",
    //     成本价: "costPrice",
    //     白金价: "silverPrice",
    //     利润: "profit",
    //     利润率: "profitRate",
    //     近7天销售量: "sales_last7Days",
    //     可售库存: "sellableInventory",
    //     可售天数: "sellableDays",
    //     合计库存: "totalInventory",
    //     销量总计: "totalSales",
    //   };
    //   sortField = sortMap[UserForm1.ComboBox6.Value];
    //   sortAscending = UserForm1.OptionButton26?.Value;
    // }

    // ----- 6. 获取数据 -----
    const products = this._repository.query("Product", this._context.query);

    // ----- 9. 创建新工作簿 -----
    const sourceWb = this._excelDAO.getWorkbook();
    sourceWb.Sheets(this._config.get("Product").worksheet).Copy();
    const newWb = ActiveWorkbook;

    // ----- 10. 获取可见列 -----
    const visibleColumns = expandedColumns.filter(
      (col) => col.visible !== false,
    );

    // ----- 11. 按分组输出 -----
    if (groupBy) {
      const groups = {};

      products.forEach((product) => {
        let groupValue = product[groupBy];
        if (
          groupValue === undefined ||
          groupValue === null ||
          groupValue === ""
        ) {
          groupValue = "未分组";
        }

        if (!groups[groupValue]) {
          groups[groupValue] = [];
        }
        groups[groupValue].push(product);
      });

      Object.entries(groups).forEach(([groupValue, groupProducts]) => {
        this._excelDAO.copySheet("Product", newWb);
        const sheet = ActiveSheet;

        let sheetName = String(groupValue).replace(/[\\/:*?\"<>|]/g, "");
        if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
        sheet.Name = sheetName;

        this._writeReportSheet(sheet, groupProducts, visibleColumns);
      });
    } else {
      const sheet = newWb.Sheets(1);
      this._writeReportSheet(sheet, products, visibleColumns);
    }

    // ----- 12. 删除原始工作表 -----
    try {
      newWb.Sheets(this._config.get("Product").worksheet).Delete();
    } catch (e) {
      // 忽略删除错误
    }

    return newWb;
  }

  // 写入报表工作表
  _writeReportSheet(sheet, products, columns) {
    const outputData = [];

    // 标题行
    outputData.push(columns.map((col) => col.title));

    // 数据行
    products.forEach((product) => {
      const row = columns.map((col) => {
        let value = this._getFieldValue(product, col);

        if (typeof value === "number") {
          if (col.field.includes("Rate") || col.field.includes("率")) {
            return value;
          }
          return Number(value.toFixed(2));
        }

        return value !== undefined && value !== null ? String(value) : "";
      });

      outputData.push(row);
    });

    sheet.Cells.ClearContents();

    if (outputData.length > 0) {
      const range = sheet
        .Range("A1")
        .Resize(outputData.length, outputData[0].length);
      range.Value2 = outputData;

      columns.forEach((col, index) => {
        const columnIndex = index + 1;
        const column = sheet.Columns(columnIndex);

        column.ColumnWidth = col.width || 10;

        if (col.format) {
          column.NumberFormat = col.format;
        }

        if (col.color) {
          const headerCell = sheet.Cells(1, columnIndex);
          excelColorReader.applyColorToRange(headerCell, col.color);
        }
      });

      const headerRange = sheet.Rows("1:1");
      headerRange.Font.Bold = true;
      headerRange.RowHeight = 20;
    }
  }

  // 预览报表
  previewReport() {
    try {
      const wb = this.generateReport();
      wb.Activate();
      MsgBox("报表生成成功！", 64, "成功");
    } catch (e) {
      MsgBox("报表生成失败：" + e.message, 16, "错误");
    }
  }
}
