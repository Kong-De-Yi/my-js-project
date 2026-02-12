// ============================================================================
// 报表输出引擎
// 功能：从工作表读取报表配置，动态生成报表
// 特点：列顺序、显示列完全可配置，无需修改代码
// ============================================================================

class ReportEngine {
  constructor(repository, excelDAO, profitCalculator) {
    this._repository = repository;
    this._excelDAO = excelDAO;
    this._profitCalculator = profitCalculator;
    this._config = DataConfig.getInstance();
  }

  // 从工作表读取报表列配置
  _loadReportColumnConfig() {
    const wb = this._excelDAO.getWorkbook();

    // 检查是否存在报表配置表
    let hasConfigSheet = false;
    try {
      hasConfigSheet = wb.Sheets("报表配置") !== undefined;
    } catch (e) {
      hasConfigSheet = false;
    }

    if (!hasConfigSheet) {
      return this._getDefaultColumnConfig();
    }

    try {
      const sheet = wb.Sheets("报表配置");
      const usedRange = sheet.UsedRange;

      if (!usedRange || usedRange.Value2 === null) {
        return this._getDefaultColumnConfig();
      }

      const data = usedRange.Value2;
      const titleRow = data[0];

      // 查找列索引
      const fieldIdx = titleRow.findIndex((t) => String(t).includes("字段"));
      const titleIdx = titleRow.findIndex((t) => String(t).includes("标题"));
      const widthIdx = titleRow.findIndex((t) => String(t).includes("宽度"));
      const visibleIdx = titleRow.findIndex((t) => String(t).includes("显示"));
      const orderIdx = titleRow.findIndex((t) => String(t).includes("顺序"));

      if (fieldIdx === -1) {
        return this._getDefaultColumnConfig();
      }

      const columns = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || !row[fieldIdx]) continue;

        columns.push({
          field: String(row[fieldIdx]),
          title: row[titleIdx] || row[fieldIdx],
          width: row[widthIdx] ? Number(row[widthIdx]) : 10,
          visible:
            row[visibleIdx] === undefined
              ? true
              : String(row[visibleIdx]).toLowerCase() === "是",
          order: row[orderIdx] ? Number(row[orderIdx]) : i,
        });
      }

      if (columns.length === 0) {
        return this._getDefaultColumnConfig();
      }

      // 按顺序排序
      return columns.sort((a, b) => (a.order || 999) - (b.order || 999));
    } catch (e) {
      return this._getDefaultColumnConfig();
    }
  }

  // 默认列配置（当没有配置表时）
  _getDefaultColumnConfig() {
    const productConfig = this._config.get("Product");
    const columns = [];

    // 核心字段
    const defaultFields = [
      "itemNumber",
      "styleNumber",
      "color",
      "thirdLevelCategory",
      "firstListingTime",
      "salesAge",
      "itemStatus",
      "marketingPositioning",
      "costPrice",
      "silverPrice",
      "finalPrice",
      "profit",
      "profitRate",
      "sellableInventory",
      "sellableDays",
      "totalInventory",
      "salesQuantityOfLast7Days",
      "unitPriceOfLast7Days",
      "totalSales",
    ];

    defaultFields.forEach((field, index) => {
      const fieldConfig = productConfig.fields[field];
      if (fieldConfig) {
        columns.push({
          field,
          title: fieldConfig.title || field,
          width: field === "itemNumber" ? 15 : 12,
          visible: true,
          order: index + 1,
        });
      }
    });

    return columns;
  }

  // 从UI获取筛选条件
  _buildQueryFromUI() {
    const query = {};

    // 主销季节
    const seasons = [];
    if (UserForm1.CheckBox2.Value) seasons.push("春秋");
    if (UserForm1.CheckBox3.Value) seasons.push("夏");
    if (UserForm1.CheckBox4.Value) seasons.push("冬");
    if (UserForm1.CheckBox5.Value) seasons.push("四季");
    if (seasons.length > 0) query.mainSalesSeason = seasons;

    // 适用性别
    const genders = [];
    if (UserForm1.CheckBox14.Value) genders.push("男童");
    if (UserForm1.CheckBox15.Value) genders.push("女童");
    if (UserForm1.CheckBox16.Value) genders.push("中性");
    if (genders.length > 0) query.applicableGender = genders;

    // 商品状态
    const statuses = [];
    if (UserForm1.CheckBox17.Value) statuses.push("商品上线");
    if (UserForm1.CheckBox18.Value) statuses.push("部分上线");
    if (UserForm1.CheckBox19.Value) statuses.push("商品下线");
    if (statuses.length > 0) query.itemStatus = statuses;

    // 营销定位
    const positions = [];
    if (UserForm1.CheckBox20.Value) positions.push("引流款");
    if (UserForm1.CheckBox21.Value) positions.push("利润款");
    if (UserForm1.CheckBox22.Value) positions.push("清仓款");
    if (positions.length > 0) query.marketingPositioning = positions;

    // 备货模式
    const stockModes = [];
    if (UserForm1.CheckBox23.Value) stockModes.push("现货");
    if (UserForm1.CheckBox24.Value) stockModes.push("通版通货");
    if (UserForm1.CheckBox25.Value) stockModes.push("专版通货");
    if (stockModes.length > 0) query.stockingMode = stockModes;

    // 活动状态
    if (UserForm1.OptionButton23.Value) query.activityStatus = "活动中";
    if (UserForm1.OptionButton24.Value) query.activityStatus = "未提报";

    // 是否破价
    if (UserForm1.CheckBox10.Value) query.isPriceBroken = true;

    // 中台券
    if (UserForm1.CheckBox8.Value) query.hasUserOperations1 = true;
    if (UserForm1.CheckBox9.Value) query.hasUserOperations2 = true;

    // 是否断码
    if (UserForm1.CheckBox12.Value) query.isOutOfStock = true;

    // 数值范围
    if (UserForm1.TextEdit1.Value || UserForm1.TextEdit11.Value) {
      query.salesAge = [
        UserForm1.TextEdit1.Value
          ? Number(UserForm1.TextEdit1.Value)
          : undefined,
        UserForm1.TextEdit11.Value
          ? Number(UserForm1.TextEdit11.Value)
          : undefined,
      ];
    }

    if (UserForm1.TextEdit3.Value || UserForm1.TextEdit4.Value) {
      query.profit = [
        UserForm1.TextEdit3.Value
          ? Number(UserForm1.TextEdit3.Value)
          : undefined,
        UserForm1.TextEdit4.Value
          ? Number(UserForm1.TextEdit4.Value)
          : undefined,
      ];
    }

    if (UserForm1.TextEdit5.Value || UserForm1.TextEdit6.Value) {
      query.profitRate = [
        UserForm1.TextEdit5.Value
          ? Number(UserForm1.TextEdit5.Value) / 100
          : undefined,
        UserForm1.TextEdit6.Value
          ? Number(UserForm1.TextEdit6.Value) / 100
          : undefined,
      ];
    }

    if (UserForm1.TextEdit7.Value || UserForm1.TextEdit8.Value) {
      query.sellableInventory = [
        UserForm1.TextEdit7.Value
          ? Number(UserForm1.TextEdit7.Value)
          : undefined,
        UserForm1.TextEdit8.Value
          ? Number(UserForm1.TextEdit8.Value)
          : undefined,
      ];
    }

    if (UserForm1.TextEdit9.Value || UserForm1.TextEdit10.Value) {
      query.sellableDays = [
        UserForm1.TextEdit9.Value
          ? Number(UserForm1.TextEdit9.Value)
          : undefined,
        UserForm1.TextEdit10.Value
          ? Number(UserForm1.TextEdit10.Value)
          : undefined,
      ];
    }

    if (UserForm1.TextEdit14.Value || UserForm1.TextEdit15.Value) {
      query.salesQuantityOfLast7Days = [
        UserForm1.TextEdit14.Value
          ? Number(UserForm1.TextEdit14.Value)
          : undefined,
        UserForm1.TextEdit15.Value
          ? Number(UserForm1.TextEdit15.Value)
          : undefined,
      ];
    }

    // 销量排名
    if (UserForm1.TextEdit16.Value) {
      query.topProductsBySales = Number(UserForm1.TextEdit16.Value);
    }

    return query;
  }

  // 应用筛选条件
  _applyQuery(products, query) {
    if (Object.keys(query).length === 0) {
      return products;
    }

    return products.filter((product) => {
      return Object.entries(query).every(([key, condition]) => {
        // 数组条件（多选）
        if (Array.isArray(condition) && !Array.isArray(product[key])) {
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

        // 布尔条件
        if (condition === true) {
          return (
            product[key] !== undefined &&
            product[key] !== null &&
            product[key] !== ""
          );
        }

        // 精确匹配
        return product[key] === condition;
      });
    });
  }

  // 应用排序
  _applySort(products, sortField, sortAscending) {
    if (!sortField) return products;

    const productConfig = this._config.get("Product");
    const fieldConfig = productConfig.fields[sortField];

    return [...products].sort((a, b) => {
      let valA = a[sortField];
      let valB = b[sortField];

      if (fieldConfig?.type === "number") {
        valA = Number(valA || 0);
        valB = Number(valB || 0);
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

  // 生成报表
  generateReport() {
    // 1. 读取报表列配置
    const columns = this._loadReportColumnConfig();

    // 2. 获取筛选条件
    const query = this._buildQueryFromUI();

    // 3. 获取分组字段
    let groupBy = null;
    if (UserForm1.ComboBox4.Value) {
      const groupField = {
        上市年份: "listingYear",
        四级品类: "fourthLevelCategory",
        运营分类: "operationClassification",
        下线原因: "offlineReason",
        三级品类: "thirdLevelCategory",
      }[UserForm1.ComboBox4.Value];

      if (groupField) {
        groupBy = groupField;
      }
    }

    // 4. 获取排序字段
    let sortField = null;
    let sortAscending = true;
    if (UserForm1.ComboBox6.Value) {
      sortField = {
        首次上架时间: "firstListingTime",
        成本价: "costPrice",
        白金价: "silverPrice",
        利润: "profit",
        利润率: "profitRate",
        近7天件单价: "unitPriceOfLast7Days",
        近7天曝光UV: "exposureUVOfLast7Days",
        近7天商详UV: "productDetailsUVOfLast7Days",
        近7天加购UV: "addToCartUVOfLast7Days",
        近7天客户数: "customerCountOfLast7Days",
        近7天拒退数: "rejectAndReturnCountOfLast7Days",
        近7天销售量: "salesQuantityOfLast7Days",
        近7天销售额: "salesAmountOfLast7Days",
        近7天点击率: "clickThroughRateOfLast7Days",
        近7天加购率: "addToCartRateOfLast7Days",
        近7天转化率: "purchaseRateOfLast7Days",
        近7天拒退率: "rejectAndReturnRateOfLast7Days",
        近7天款销量: "styleSalesOfLast7Days",
        可售库存: "sellableInventory",
        可售天数: "sellableDays",
        合计库存: "totalInventory",
        成品合计: "finishedGoodsTotalInventory",
        通货合计: "generalGoodsTotalInventory",
        销量总计: "totalSales",
      }[UserForm1.ComboBox6.Value];

      sortAscending = UserForm1.OptionButton26.Value; // true:升序, false:降序
    }

    // 5. 获取所有商品数据
    let products = this._repository.findProducts();

    // 6. 应用筛选
    products = this._applyQuery(products, query);

    // 7. 应用排序
    products = this._applySort(products, sortField, sortAscending);

    // 8. 销量排名截取
    if (query.topProductsBySales) {
      products = products.slice(0, query.topProductsBySales);
    }

    // 9. 创建新工作簿
    const sourceWb = this._excelDAO.getWorkbook();
    sourceWb.Sheets(this._config.get("Product").worksheet).Copy();
    const newWb = ActiveWorkbook;

    // 10. 获取可见列
    const visibleColumns = columns.filter((col) => col.visible !== false);

    // 11. 按分组输出
    if (groupBy) {
      const groups = {};

      products.forEach((product) => {
        const groupValue = product[groupBy] || "未分组";
        if (!groups[groupValue]) {
          groups[groupValue] = [];
        }
        groups[groupValue].push(product);
      });

      Object.entries(groups).forEach(([groupValue, groupProducts]) => {
        // 复制工作表
        this._excelDAO.copySheet("Product", newWb);
        const sheet = ActiveSheet;

        // 清理非法字符
        const sheetName = String(groupValue).replace(/[\\/:*?\"<>|]/g, "");
        sheet.Name = sheetName.substring(0, 31);

        // 写入数据
        this._writeReportSheet(sheet, groupProducts, visibleColumns);
      });
    } else {
      // 单个工作表输出
      const sheet = newWb.Sheets(1);
      this._writeReportSheet(sheet, products, visibleColumns);
    }

    // 12. 删除原始工作表
    try {
      newWb.Sheets(this._config.get("Product").worksheet).Delete();
    } catch (e) {
      // 忽略删除失败
    }

    // 13. 应用冻结窗格
    if (UserForm1.CheckBox33.Value) {
      const sheet = ActiveSheet;
      sheet.Activate();
      const win = ActiveWindow;
      win.FreezePanes = false;
      win.SplitRow = 1;
      win.FreezePanes = true;
    }

    return newWb;
  }

  // 写入报表工作表
  _writeReportSheet(sheet, products, columns) {
    // 构建输出数据
    const outputData = [];

    // 标题行
    outputData.push(columns.map((col) => col.title));

    // 数据行
    products.forEach((product) => {
      const row = columns.map((col) => {
        let value = product[col.field];

        // 格式化
        if (typeof value === "number") {
          if (col.field.includes("Rate") || col.field.includes("率")) {
            return value; // 保留小数
          }
          if (Number.isInteger(value)) {
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

      // 设置列宽
      columns.forEach((col, index) => {
        sheet.Columns(index + 1).ColumnWidth = col.width || 10;
      });
    }
  }
}
