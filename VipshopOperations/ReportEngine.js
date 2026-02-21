// ============================================================================
// ReportEngine.js - 报表引擎（支持字段展开）
// 功能：根据模板配置生成报表，支持字段自动展开
// ============================================================================

class ReportEngine {
  /**
   * 构造函数
   * @param {Object} repository - 数据仓库实例
   * @param {Object} excelDAO - Excel数据访问对象
   */
  constructor(repository, excelDAO) {
    this._repository = repository;
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();
    this._templateManager = new ReportTemplateManager(repository, excelDAO);

    // 初始化销售统计服务
    this._salesStatisticsService = new SalesStatisticsService(repository);
    this._salesStatisticsFields = new SalesStatisticsFields(
      this._salesStatisticsService,
    );

    // 缓存所有统计字段配置
    this._statisticsFieldsMap = this._buildStatisticsFieldsMap();
  }

  /**
   * 构建统计字段映射（包含展开逻辑）
   * @returns {Map} 字段名到字段配置的映射
   * @private
   */
  _buildStatisticsFieldsMap() {
    const map = new Map();
    const baseFields = this._salesStatisticsFields.getAllFields();

    // 展开所有可展开字段
    baseFields.forEach((field) => {
      if (field.type === "expandable") {
        const expanded = this._salesStatisticsFields.expandField(field);
        expanded.forEach((expandedField) => {
          map.set(expandedField.field, expandedField);
        });
      } else {
        map.set(field.field, field);
      }
    });

    return map;
  }

  /**
   * 获取所有可用的字段（用于模板配置）
   * @returns {Array} 字段配置数组
   */
  getAvailableFields() {
    const baseFields = this._salesStatisticsFields.getAllFields();

    return baseFields.map((field) => ({
      field: field.field,
      title: field.title,
      type: field.type,
      group: field.group,
      description: field.description,
    }));
  }

  /**
   * 初始化模板
   * @returns {Map} 初始化后的模板列表
   */
  initializeTemplates() {
    return this._templateManager.initializeDefaultTemplates();
  }

  /**
   * 获取模板列表
   * @returns {Array} 模板名称数组
   */
  getTemplateList() {
    return this._templateManager.getTemplateList();
  }

  /**
   * 设置当前模板
   * @param {string} templateName - 模板名称
   * @returns {boolean} 是否设置成功
   */
  setCurrentTemplate(templateName) {
    return this._templateManager.setCurrentTemplate(templateName);
  }

  /**
   * 获取当前模板
   * @returns {string|null} 当前模板名称
   */
  getCurrentTemplate() {
    return this._templateManager.getCurrentTemplate();
  }

  /**
   * 刷新模板
   * @returns {Map} 刷新后的模板列表
   */
  refreshTemplates() {
    this._templateManager.refresh();
    // 重新构建字段映射（因为时间变化）
    this._statisticsFieldsMap = this._buildStatisticsFieldsMap();
    return this._templateManager.loadTemplates();
  }

  /**
   * 从UI获取筛选条件
   * @returns {Object} 查询条件对象
   * @private
   */
  _buildQueryFromUI() {
    const query = {};

    // 主销季节
    const seasons = [];
    if (UserForm1.CheckBox2?.Value) seasons.push("春秋");
    if (UserForm1.CheckBox3?.Value) seasons.push("夏");
    if (UserForm1.CheckBox4?.Value) seasons.push("冬");
    if (UserForm1.CheckBox5?.Value) seasons.push("四季");
    if (seasons.length > 0) query.mainSalesSeason = seasons;

    // 适用性别
    const genders = [];
    if (UserForm1.CheckBox14?.Value) genders.push("男童");
    if (UserForm1.CheckBox15?.Value) genders.push("女童");
    if (UserForm1.CheckBox16?.Value) genders.push("中性");
    if (genders.length > 0) query.applicableGender = genders;

    // 商品状态
    const statuses = [];
    if (UserForm1.CheckBox17?.Value) statuses.push("商品上线");
    if (UserForm1.CheckBox18?.Value) statuses.push("部分上线");
    if (UserForm1.CheckBox19?.Value) statuses.push("商品下线");
    if (statuses.length > 0) query.itemStatus = statuses;

    // 营销定位
    const positions = [];
    if (UserForm1.CheckBox20?.Value) positions.push("引流款");
    if (UserForm1.CheckBox21?.Value) positions.push("利润款");
    if (UserForm1.CheckBox22?.Value) positions.push("清仓款");
    if (positions.length > 0) query.marketingPositioning = positions;

    // 备货模式
    const stockModes = [];
    if (UserForm1.CheckBox23?.Value) stockModes.push("现货");
    if (UserForm1.CheckBox24?.Value) stockModes.push("通版通货");
    if (UserForm1.CheckBox25?.Value) stockModes.push("专版通货");
    if (stockModes.length > 0) query.stockingMode = stockModes;

    // 活动状态
    if (UserForm1.OptionButton23?.Value) query.activityStatus = "活动中";
    if (UserForm1.OptionButton24?.Value) query.activityStatus = "未提报";

    // 数值范围
    if (UserForm1.TextEdit1?.Value || UserForm1.TextEdit11?.Value) {
      query.salesAge = [
        UserForm1.TextEdit1?.Value
          ? Number(UserForm1.TextEdit1.Value)
          : undefined,
        UserForm1.TextEdit11?.Value
          ? Number(UserForm1.TextEdit11.Value)
          : undefined,
      ];
    }

    if (UserForm1.TextEdit3?.Value || UserForm1.TextEdit4?.Value) {
      query.profit = [
        UserForm1.TextEdit3?.Value
          ? Number(UserForm1.TextEdit3.Value)
          : undefined,
        UserForm1.TextEdit4?.Value
          ? Number(UserForm1.TextEdit4.Value)
          : undefined,
      ];
    }

    if (UserForm1.TextEdit7?.Value || UserForm1.TextEdit8?.Value) {
      query.sellableInventory = [
        UserForm1.TextEdit7?.Value
          ? Number(UserForm1.TextEdit7.Value)
          : undefined,
        UserForm1.TextEdit8?.Value
          ? Number(UserForm1.TextEdit8.Value)
          : undefined,
      ];
    }

    if (UserForm1.TextEdit14?.Value || UserForm1.TextEdit15?.Value) {
      query.salesQuantityOfLast7Days = [
        UserForm1.TextEdit14?.Value
          ? Number(UserForm1.TextEdit14.Value)
          : undefined,
        UserForm1.TextEdit15?.Value
          ? Number(UserForm1.TextEdit15.Value)
          : undefined,
      ];
    }

    return query;
  }

  /**
   * 应用筛选条件
   * @param {Array} products - 商品数据
   * @param {Object} query - 查询条件
   * @returns {Array} 筛选后的数据
   * @private
   */
  _applyQuery(products, query) {
    if (Object.keys(query).length === 0) {
      return products;
    }

    return products.filter((product) => {
      return Object.entries(query).every(([key, condition]) => {
        if (Array.isArray(condition)) {
          if (key === "offlineReason") {
            return (
              condition.includes(product[key]) &&
              product.itemStatus === "商品下线"
            );
          }
          return condition.includes(product[key]);
        }

        if (Array.isArray(condition) && condition.length === 2) {
          const [min, max] = condition;
          const value = product[key];

          if (value === undefined || value === null) return false;

          if (min !== undefined && max !== undefined) {
            return Number(value) >= min && Number(value) <= max;
          } else if (min !== undefined) {
            return Number(value) >= min;
          } else if (max !== undefined) {
            return Number(value) <= max;
          }
          return true;
        }

        return product[key] === condition;
      });
    });
  }

  /**
   * 应用排序
   * @param {Array} products - 商品数据
   * @param {string} sortField - 排序字段
   * @param {boolean} sortAscending - 是否升序
   * @returns {Array} 排序后的数据
   * @private
   */
  _applySort(products, sortField, sortAscending) {
    if (!sortField) return products;

    return [...products].sort((a, b) => {
      let valA = a[sortField];
      let valB = b[sortField];

      if (typeof valA === "number" && typeof valB === "number") {
        return sortAscending ? valA - valB : valB - valA;
      }

      if (sortField === "firstListingTime") {
        const dateA = Date.parse(valA) || 0;
        const dateB = Date.parse(valB) || 0;
        return sortAscending ? dateA - dateB : dateB - dateA;
      }

      valA = String(valA || "");
      valB = String(valB || "");
      return sortAscending
        ? valA.localeCompare(valB)
        : valB.localeCompare(valA);
    });
  }

  /**
   * 获取字段值（支持统计字段）
   * @param {Object} product - 商品对象
   * @param {Object} column - 列配置
   * @returns {*} 字段值
   * @private
   */
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

  /**
   * 生成报表
   * @returns {Object} 生成的工作簿对象
   */
  generateReport() {
    // ----- 1. 获取当前模板的列配置 -----
    let columns = this._templateManager.getCurrentColumns();

    // ----- 2. 展开所有可展开字段 -----
    const expandedColumns = [];
    columns.forEach((col) => {
      // 获取原始字段定义（未展开的）
      const baseField = this._salesStatisticsFields.getField(col.field);

      if (baseField?.type === "expandable") {
        // 展开抽象字段
        const expanded = this._salesStatisticsFields.expandField(baseField);
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

    // ----- 3. 获取筛选条件 -----
    const query = this._buildQueryFromUI();

    // ----- 4. 获取分组字段 -----
    let groupBy = null;
    if (UserForm1.ComboBox4?.Value) {
      const groupMap = {
        上市年份: "listingYear",
        四级品类: "fourthLevelCategory",
        运营分类: "operationClassification",
        下线原因: "offlineReason",
        三级品类: "thirdLevelCategory",
      };
      groupBy = groupMap[UserForm1.ComboBox4.Value];
    }

    // ----- 5. 获取排序字段 -----
    let sortField = null;
    let sortAscending = true;
    if (UserForm1.ComboBox6?.Value) {
      const sortMap = {
        首次上架时间: "firstListingTime",
        成本价: "costPrice",
        白金价: "silverPrice",
        利润: "profit",
        利润率: "profitRate",
        近7天销售量: "sales_last7Days",
        可售库存: "sellableInventory",
        可售天数: "sellableDays",
        合计库存: "totalInventory",
        销量总计: "totalSales",
      };
      sortField = sortMap[UserForm1.ComboBox6.Value];
      sortAscending = UserForm1.OptionButton26?.Value;
    }

    // ----- 6. 获取所有商品数据 -----
    let products = this._repository.findProducts();

    // ----- 7. 应用筛选 -----
    products = this._applyQuery(products, query);

    // ----- 8. 应用排序 -----
    products = this._applySort(products, sortField, sortAscending);

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

  /**
   * 写入报表工作表
   * @param {Object} sheet - Excel工作表
   * @param {Array} products - 商品数据
   * @param {Array} columns - 列配置数组
   * @private
   */
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

  /**
   * 预览报表
   */
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
