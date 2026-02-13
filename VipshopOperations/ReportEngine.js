// ============================================================================
// 报表引擎
// 功能：根据模板配置生成报表，支持列背景色
// ============================================================================

class ReportEngine {
  constructor(repository, excelDAO) {
    this._repository = repository;
    this._excelDAO = excelDAO;
    this._config = dataConfig;
    this._templateManager = new ReportTemplateManager(repository, excelDAO);
  }

  /**
   * 初始化模板
   */
  initializeTemplates() {
    return this._templateManager.initializeDefaultTemplates();
  }

  /**
   * 获取模板列表
   */
  getTemplateList() {
    return this._templateManager.getTemplateList();
  }

  /**
   * 设置当前模板
   */
  setCurrentTemplate(templateName) {
    return this._templateManager.setCurrentTemplate(templateName);
  }

  /**
   * 获取当前模板
   */
  getCurrentTemplate() {
    return this._templateManager.getCurrentTemplate();
  }

  /**
   * 刷新模板
   */
  refreshTemplates() {
    return this._templateManager.refresh();
  }

  /**
   * 从UI获取筛选条件
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
   */
  _applyQuery(products, query) {
    if (Object.keys(query).length === 0) {
      return products;
    }

    return products.filter((product) => {
      return Object.entries(query).every(([key, condition]) => {
        // 数组条件（多选）
        if (Array.isArray(condition)) {
          if (key === "offlineReason") {
            return (
              condition.includes(product[key]) &&
              product.itemStatus === "商品下线"
            );
          }
          return condition.includes(product[key]);
        }

        // 范围条件
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

        // 精确匹配
        return product[key] === condition;
      });
    });
  }

  /**
   * 应用排序
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
   * 生成报表
   */
  generateReport() {
    // 1. 获取当前模板的列配置
    const columns = this._templateManager.getCurrentColumns();

    // 2. 获取筛选条件
    const query = this._buildQueryFromUI();

    // 3. 获取分组字段
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

    // 4. 获取排序字段
    let sortField = null;
    let sortAscending = true;
    if (UserForm1.ComboBox6?.Value) {
      const sortMap = {
        首次上架时间: "firstListingTime",
        成本价: "costPrice",
        白金价: "silverPrice",
        利润: "profit",
        利润率: "profitRate",
        近7天销售量: "salesQuantityOfLast7Days",
        可售库存: "sellableInventory",
        可售天数: "sellableDays",
        合计库存: "totalInventory",
        销量总计: "totalSales",
      };
      sortField = sortMap[UserForm1.ComboBox6.Value];
      sortAscending = UserForm1.OptionButton26?.Value; // true:升序
    }

    // 5. 获取所有商品数据
    let products = this._repository.findProducts();

    // 6. 应用筛选
    products = this._applyQuery(products, query);

    // 7. 应用排序
    products = this._applySort(products, sortField, sortAscending);

    // 8. 创建新工作簿
    const sourceWb = this._excelDAO.getWorkbook();
    sourceWb.Sheets(this._config.get("Product").worksheet).Copy();
    const newWb = ActiveWorkbook;

    // 9. 获取可见列
    const visibleColumns = columns.filter((col) => col.visible !== false);

    // 10. 按分组输出
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

    // 11. 删除原始工作表
    try {
      newWb.Sheets(this._config.get("Product").worksheet).Delete();
    } catch (e) {
      // 忽略
    }

    return newWb;
  }

  /**
   * 写入报表工作表（支持颜色）
   */
  _writeReportSheet(sheet, products, columns) {
    // 构建输出数据
    const outputData = [];

    // 标题行
    outputData.push(columns.map((col) => col.title));

    // 数据行
    products.forEach((product) => {
      const row = columns.map((col) => {
        let value = product[col.field];

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

    // 清空并写入
    sheet.Cells.ClearContents();

    if (outputData.length > 0) {
      const range = sheet
        .Range("A1")
        .Resize(outputData.length, outputData[0].length);
      range.Value2 = outputData;

      // 设置列宽和应用颜色
      columns.forEach((col, index) => {
        const column = sheet.Columns(index + 1);
        column.ColumnWidth = col.width || 10;

        // 应用标题颜色
        if (col.color) {
          const headerCell = sheet.Cells(1, index + 1);
          excelColorReader.applyColorToRange(headerCell, col.color);
        }

        // 应用数字格式
        if (col.format) {
          column.NumberFormat = col.format;
        }
      });
    }
  }
}
