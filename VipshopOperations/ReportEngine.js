/**
 * 报表引擎 - 负责报表的生成、预览和模板管理
 *
 * @class ReportEngine
 * @description 作为报表生成的核心业务逻辑层，提供以下功能：
 * - 基于模板的报表列配置管理
 * - 统计字段的动态计算和展开（如"近N天销量"展开为多个具体字段）
 * - 数据分组和筛选
 * - 自动创建报表工作簿
 * - 模板的保存、加载和管理
 *
 * 该类采用单例模式，确保全局只有一个报表引擎实例。
 *
 * @example
 * // 获取报表引擎实例
 * const reportEngine = ReportEngine.getInstance(repository, excelDAO);
 *
 * // 设置查询条件和分组
 * reportEngine.setContext({
 *   query: { itemStatus: "上线" },
 *   groupBy: "listingYear"
 * });
 *
 * // 生成报表
 * const workbook = reportEngine.generateReport();
 * workbook.Activate();
 */
class ReportEngine {
  /** @type {ReportEngine} 单例实例 */
  static _instance = null;

  /**
   * 创建报表引擎实例
   * @param {Repository} [repository] - 数据仓库实例，若不提供则自动获取
   * @param {ExcelDAO} [excelDAO] - Excel数据访问对象实例，若不提供则自动获取
   */
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

  /**
   * 获取报表引擎的单例实例
   * @static
   * @param {Repository} [repository] - 数据仓库实例
   * @param {ExcelDAO} [excelDAO] - Excel数据访问对象实例
   * @returns {ReportEngine} 报表引擎实例
   */
  static getInstance(repository, excelDAO) {
    if (!ReportEngine._instance) {
      ReportEngine._instance = new ReportEngine(repository, excelDAO);
    }
    return ReportEngine._instance;
  }

  /**
   * 构建统计字段映射表（包含字段展开逻辑）
   * @private
   * @returns {Map<string, Object>} 字段名到字段配置的映射
   * @description
   * 构建逻辑：
   * 1. 获取所有基础统计字段（来自 StatisticsFields）
   * 2. 遍历每个字段：
   *    - 如果是可展开字段（type === "expandable"），调用 expandField 展开
   *    - 将展开后的每个具体字段加入映射
   *    - 普通字段直接加入映射
   * 3. 建立字段名到字段配置的快速查找映射
   *
   * 这样设计的目的是：
   * - 模板中可以使用抽象字段（如"近7天销量"）
   * - 运行时自动展开为具体字段（如"sales_last7Days"）
   */
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

  /**
   * 获取所有可用的统计字段（用于模板配置界面）
   * @returns {Array<Object>} 可用字段列表
   * @returns {string} return[].field - 字段名（用于程序内部识别）
   * @returns {string} return[].title - 显示标题（用于UI展示）
   * @returns {string} return[].type - 字段类型（normal/expandable/computed）
   * @returns {string} return[].group - 所属分组（如"销售分析"、"库存分析"）
   * @returns {string} return[].description - 字段描述说明
   *
   * @example
   * // 获取可用字段列表用于UI下拉框
   * const fields = reportEngine.getAvailableFields();
   * fields.forEach(f => {
   *   UserForm1.ComboBox1.AddItem(f.title);
   * });
   */
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

  /**
   * 设置报表上下文环境
   * @param {Object} context - 上下文配置
   * @param {Object} [context.query] - 数据查询条件（传递给 Repository.query）
   * @param {string} [context.groupBy] - 分组字段名（如"listingYear"、"fourthLevelCategory"）
   * @description
   * 用于设置报表生成的查询条件和分组规则。
   * 这些设置会影响 generateReport 方法的行为。
   *
   * @example
   * // 按年份分组，只查询上线的商品
   * setContext({
   *   query: {
   *     itemStatus: "上线",
   *     brandSN: { $in: ["A001", "A002"] }
   *   },
   *   groupBy: "listingYear"
   * });
   */
  setContext(context) {
    Object.assign(this._context, context);
  }

  /**
   * 初始化默认报表模板
   * @returns {boolean} 是否初始化成功
   * @description
   * 创建系统预设的默认报表模板，例如：
   * - 商品基础信息报表
   * - 销售分析报表
   * - 库存预警报表
   * - 利润分析报表
   *
   * 这些模板会被保存到 Excel 的配置工作表中。
   */
  initializeTemplates() {
    return this._templateManager.initializeDefaultTemplates();
  }

  /**
   * 获取所有可用的报表模板列表
   * @returns {Array<Object>} 模板列表
   * @returns {string} return[].name - 模板名称
   * @returns {string} return[].description - 模板描述
   * @returns {Array} return[].columns - 列配置数组
   * @returns {string} return[].columns[].field - 字段名
   * @returns {string} return[].columns[].title - 列标题
   * @returns {number} return[].columns[].width - 列宽
   * @returns {string} return[].columns[].format - 数字格式（如"#,##0.00"）
   * @returns {number} return[].columns[].color - 标题行颜色
   */
  getTemplateList() {
    return this._templateManager.getTemplateList();
  }

  /**
   * 设置当前使用的报表模板
   * @param {string} templateName - 模板名称
   * @returns {Object} 当前模板对象
   * @throws {Error} 模板不存在时抛出
   */
  setCurrentTemplate(templateName) {
    return this._templateManager.setCurrentTemplate(templateName);
  }

  /**
   * 获取当前使用的报表模板
   * @returns {Object|null} 当前模板对象，未设置时返回null
   */
  getCurrentTemplate() {
    return this._templateManager.getCurrentTemplate();
  }

  /**
   * 刷新报表模板和统计字段
   * @returns {Array} 刷新后的模板列表
   * @description
   * 刷新流程：
   * 1. 重新加载模板配置（从 Excel 读取）
   * 2. 重新构建统计字段映射（因为时间相关字段可能会变化）
   * 3. 返回更新后的模板列表
   *
   * 通常在以下情况调用：
   * - 用户手动点击刷新按钮
   * - 切换日期后（影响"近N天"类字段）
   */
  refreshTemplates() {
    this._templateManager.refresh();
    // 重新构建字段映射（因为时间变化）
    this._statisticsFieldsMap = this._buildStatisticsFieldsMap();
    return this._templateManager.loadTemplates();
  }

  /**
   * 获取字段值（支持普通字段和统计字段）
   * @private
   * @param {Object} product - 产品对象
   * @param {Object} column - 列配置
   * @param {string} column.field - 字段名
   * @returns {*} 字段值
   * @description
   * 获取逻辑：
   * 1. 检查字段名是否在统计字段映射中
   * 2. 如果是统计字段，调用对应的 compute 方法计算值
   * 3. 如果是普通字段，直接从产品对象中获取
   * 4. 计算失败时返回 0 并记录错误日志
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
   * @returns {Excel.Workbook} 生成的报表工作簿
   * @throws {Error} 报表生成失败时抛出（如模板配置错误、数据读取失败等）
   *
   * @description
   * 报表生成流程（共12步）：
   * 1. 获取当前模板的列配置
   * 2. 展开所有可展开字段（如"近7天销量"展开为具体字段）
   * 3. 获取筛选条件（从 UI 或上下文）
   * 4. 获取分组字段（从上下文）
   * 5. 获取排序字段（从 UI）
   * 6. 从仓库查询产品数据
   * 7. （步骤7-8被注释，可能是预留）
   * 8.
   * 9. 复制产品模板创建新工作簿
   * 10. 过滤出可见列
   * 11. 按分组输出数据（如果设置了分组）：
   *     - 按分组值将产品分组
   *     - 为每个组复制一个工作表
   *     - 工作表命名为分组值（过滤非法字符，限制长度≤31）
   * 12. 删除原始模板工作表
   *
   * @example
   * // 生成按年份分组的报表
   * reportEngine.setContext({
   *   query: { itemStatus: "上线" },
   *   groupBy: "listingYear"
   * });
   * const wb = reportEngine.generateReport();
   * wb.Activate();
   */
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
    newWb.Sheets(this._config.get("Product").worksheet).Cells.Clear();

    // ----- 10. 获取可见列 -----
    const visibleColumns = expandedColumns.filter(
      (col) => col.visible !== false,
    );

    // ----- 11. 按分组输出 -----
    const groupBy = this._context.groupBy;
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
      Application.DisplayAlerts = false;
      newWb.Sheets(this._config.get("Product").worksheet).Delete();
    } catch (e) {
      // 忽略删除错误
    } finally {
      Application.DisplayAlerts = true;
    }

    return newWb;
  }

  /**
   * 写入报表工作表
   * @private
   * @param {Excel.Worksheet} sheet - 目标工作表
   * @param {Object[]} products - 产品数据数组
   * @param {Array<Object>} columns - 列配置数组
   * @description
   * 写入流程：
   * 1. 构建输出数据二维数组
   *    - 第一行：标题行（使用列的 title）
   *    - 后续行：数据行，每列调用 _getFieldValue 获取值
   * 2. 数值处理：
   *    - 保留两位小数
   *    - 利润率字段保持原样（不四舍五入）
   * 3. 清空工作表原有内容
   * 4. 写入数据（从 A1 单元格开始）
   * 5. 设置列格式：
   *    - 列宽（默认10）
   *    - 数字格式（如果配置了 format）
   *    - 标题行颜色（如果配置了 color）
   * 6. 设置标题行格式：加粗、行高20
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
   * @returns {void}
   * @description
   * 预览流程：
   * 1. 调用 generateReport 生成报表
   * 2. 激活生成的报表工作簿（显示给用户）
   * 3. 显示成功消息框
   * 4. 如果生成失败，显示错误消息框
   *
   * @example
   * // 在按钮点击事件中调用
   * function onPreviewButtonClick() {
   *   const reportEngine = ReportEngine.getInstance();
   *   reportEngine.previewReport();
   * }
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
