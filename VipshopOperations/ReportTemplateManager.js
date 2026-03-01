/**
 * 报表模板管理器 - 负责报表模板的加载、管理和持久化
 *
 * @class ReportTemplateManager
 * @description 作为报表模板的核心管理类，提供以下功能：
 * - 从 Excel 的"报表配置"工作表加载模板配置
 * - 创建和管理多个报表模板（如库存预警、利润分析、销售分析等）
 * - 模板字段的排序、可见性控制
 * - 标题行颜色的读取和应用
 * - 默认模板的初始化
 *
 * 模板配置数据结构：
 * - 每个模板包含多个字段（列）配置
 * - 字段配置包括：字段名、标题、列宽、可见性、显示顺序、数字格式、标题颜色
 * - 标题颜色单独存储在 Excel 单元格背景色中
 *
 * 该类采用单例模式，确保全局只有一个模板管理器实例。
 *
 * @example
 * // 获取模板管理器实例
 * const templateManager = ReportTemplateManager.getInstance(repository, excelDAO);
 *
 * // 初始化默认模板
 * templateManager.initializeDefaultTemplates();
 *
 * // 获取所有模板名称
 * const templateList = templateManager.getTemplateList();
 *
 * // 设置当前模板
 * templateManager.setCurrentTemplate("库存预警报表");
 *
 * // 获取当前模板的列配置
 * const columns = templateManager.getCurrentColumns();
 */
class ReportTemplateManager {
  /** @type {ReportTemplateManager} 单例实例 */
  static _instance = null;

  /**
   * 创建报表模板管理器实例
   * @param {Repository} [repository] - 数据仓库实例，若不提供则自动获取
   * @param {ExcelDAO} [excelDAO] - Excel数据访问对象实例，若不提供则自动获取
   */
  constructor(repository, excelDAO) {
    if (ReportTemplateManager._instance) {
      return ReportTemplateManager._instance;
    }

    this._repository = repository || Repository.getInstance();
    this._excelDAO = excelDAO || ExcelDAO.getInstance();

    this._config = DataConfig.getInstance();
    this._templates = null;
    this._currentTemplate = null;

    ReportTemplateManager._instance = this;
  }

  /**
   * 获取报表模板管理器的单例实例
   * @static
   * @param {Repository} [repository] - 数据仓库实例
   * @param {ExcelDAO} [excelDAO] - Excel数据访问对象实例
   * @returns {ReportTemplateManager} 模板管理器实例
   */
  static getInstance(repository, excelDAO) {
    if (!ReportTemplateManager._instance) {
      ReportTemplateManager._instance = new ReportTemplateManager(
        repository,
        excelDAO,
      );
    }
    return ReportTemplateManager._instance;
  }

  /**
   * 初始化默认报表模板
   * @returns {Map<string, Array>} 初始化后的模板映射（模板名 -> 字段配置数组）
   * @description
   * 初始化流程：
   * 1. 尝试从工作表加载已有模板配置
   * 2. 如果已有模板配置，直接返回
   * 3. 如果没有配置，创建以下默认模板：
   *    - 库存预警报表：货号、款号、颜色、三级品类、可售库存、可售天数、成品库存、通货库存、合计库存、是否断码
   *    - 利润分析报表：货号、款号、营销定位、成本价、白金价、到手价、利润、利润率、中台操作字段
   *    - 销售分析报表：包含年/月/周/日维度的销量对比、近N天销量、曝光UV、商详UV、加购UV、拒退件数等
   * 4. 将默认模板保存到"报表配置"工作表
   */
  initializeDefaultTemplates() {
    // 载入工作表中配置的模板
    const templates = this.loadTemplates();

    if (templates.size > 0) {
      return templates;
    }

    // 没有配置模板则新增默认模板
    this._createDefaultTemplate("库存预警报表", [
      { field: "itemNumber", title: "货号", width: 15, order: 1 },
      { field: "styleNumber", title: "款号", width: 12, order: 2 },
      { field: "color", title: "颜色", width: 10, order: 3 },
      { field: "thirdLevelCategory", title: "三级品类", width: 12, order: 4 },
      { field: "sellableInventory", title: "可售库存", width: 10, order: 5 },
      { field: "sellableDays", title: "可售天数", width: 10, order: 6 },
      {
        field: "finishedGoodsTotalInventory",
        title: "成品库存",
        width: 10,
        order: 7,
      },
      {
        field: "generalGoodsTotalInventory",
        title: "通货库存",
        width: 10,
        order: 8,
      },
      { field: "totalInventory", title: "合计库存", width: 10, order: 9 },
      { field: "isOutOfStock", title: "是否断码", width: 10, order: 10 },
    ]);

    this._createDefaultTemplate("利润分析报表", [
      { field: "itemNumber", title: "货号", width: 15, order: 1 },
      { field: "styleNumber", title: "款号", width: 12, order: 2 },
      { field: "marketingPositioning", title: "营销定位", width: 10, order: 3 },
      { field: "costPrice", title: "成本价", width: 10, order: 4 },
      { field: "silverPrice", title: "白金价", width: 10, order: 5 },
      { field: "finalPrice", title: "到手价", width: 10, order: 6 },
      { field: "profit", title: "利润", width: 10, order: 7 },
      { field: "profitRate", title: "利润率", width: 10, order: 8 },
      { field: "userOperations1", title: "中台1", width: 8, order: 9 },
      { field: "userOperations2", title: "中台2", width: 8, order: 10 },
    ]);

    this._createDefaultTemplate("销售分析报表", [
      { field: "itemNumber", title: "货号", width: 15, order: 1 },
      { field: "styleNumber", title: "款号", width: 12, order: 2 },
      {
        field: "yearSales_beforeLast",
        title: "前年年销量",
        width: 12,
        order: 3,
      },
      { field: "yearSales_last", title: "去年年销量", width: 12, order: 4 },
      { field: "yearSales_current", title: "今年年销量", width: 12, order: 5 },
      {
        field: "monthSales_beforeLast",
        title: "前年月销量",
        width: 10,
        order: 6,
      },
      { field: "monthSales_last", title: "去年月销量", width: 10, order: 7 },
      { field: "monthSales_current", title: "今年月销量", width: 10, order: 8 },
      {
        field: "weekSales_beforeLast",
        title: "前年周销量",
        width: 10,
        order: 9,
      },
      { field: "weekSales_last", title: "去年周销量", width: 10, order: 10 },
      { field: "weekSales_current", title: "今年周销量", width: 10, order: 11 },
      {
        field: "daySales_beforeLast",
        title: "前年日销量",
        width: 12,
        order: 12,
      },
      { field: "daySales_last", title: "去年日销量", width: 12, order: 13 },
      { field: "daySales_current", title: "今年日销量", width: 12, order: 14 },
      { field: "sales_last7Days", title: "近7天销量", width: 12, order: 15 },
      {
        field: "uv_exposure_last7Days",
        title: "近7天曝光UV",
        width: 14,
        order: 16,
      },
      {
        field: "uv_productDetails_last7Days",
        title: "近7天商详UV",
        width: 14,
        order: 17,
      },
      {
        field: "uv_addToCart_last7Days",
        title: "近7天加购UV",
        width: 14,
        order: 18,
      },
      {
        field: "rejectCount_last7Days",
        title: "近7天拒退件数",
        width: 14,
        order: 19,
      },
      { field: "sales_last15Days", title: "近15天销量", width: 10, order: 20 },
      { field: "sales_last30Days", title: "近30天销量", width: 10, order: 21 },
      { field: "sales_last45Days", title: "近45天销量", width: 10, order: 22 },
    ]);

    return this.loadTemplates();
  }

  /**
   * 创建默认报表模板并保存至报表配置工作表
   * @private
   * @param {string} templateName - 模板名称
   * @param {Array<Object>} columns - 列配置数组
   * @param {string} columns[].field - 字段名
   * @param {string} columns[].title - 列标题
   * @param {number} columns[].width - 列宽
   * @param {number} columns[].order - 显示顺序
   * @description
   * 创建流程：
   * 1. 将列配置转换为 ReportTemplate 实体格式
   * 2. 查询已存在的模板配置
   * 3. 剔除同名的旧模板配置
   * 4. 合并新旧配置并保存
   *
   * ReportTemplate 实体字段：
   * - templateName: 模板名称
   * - fieldName: 字段名
   * - columnTitle: 列标题
   * - columnWidth: 列宽
   * - isVisible: 是否可见（默认"是"）
   * - displayOrder: 显示顺序
   * - numberFormat: 数字格式（默认空）
   * - description: 描述（默认空）
   */
  _createDefaultTemplate(templateName, columns) {
    // 构造报表模板项目
    const templateData = [];

    columns.forEach((col) => {
      templateData.push({
        templateName: templateName,
        fieldName: col.field,
        columnTitle: col.title,
        columnWidth: col.width || 10,
        isVisible: "是",
        displayOrder: col.order,
        numberFormat: "",
        description: "",
      });
    });

    // 查询已存在的报表模板项目
    let existingTemplates = [];
    try {
      existingTemplates = this._repository.findAll("ReportTemplate");
    } catch (e) {
      existingTemplates = [];
    }

    // 剔除已配置的同名报表模板项目
    const filtered = existingTemplates.filter(
      (t) => t.templateName !== templateName,
    );
    const allTemplates = [...filtered, ...templateData];

    this._repository.save("ReportTemplate", allTemplates);
  }

  /**
   * 读取报表配置工作表中每行的标题颜色
   * @private
   * @param {Excel.Worksheet} sheet - 报表配置工作表
   * @returns {Map<string, number>} 颜色映射，键为"模板名称|字段名"，值为VBA颜色值
   * @description
   * 读取逻辑：
   * 1. 定位"模板名称"、"字段"、"标题颜色"三列的索引
   * 2. 遍历数据行（从第2行开始）
   * 3. 对每一行，读取"标题颜色"列的单元格背景色
   * 4. 构建映射键：`${模板名称}|${字段名}`
   * 5. 将颜色值存入映射
   *
   * 这种设计允许每个字段独立配置标题颜色，实现更灵活的报表样式。
   */
  _readRowColors(sheet) {
    const colorMap = new Map();

    try {
      if (!sheet) return colorMap;

      const usedRange = sheet.UsedRange;
      if (!usedRange) return colorMap;

      const data = usedRange.Value2;
      if (!data || data.length < 2) return colorMap;

      const headers = data[0];

      let nameColIdx = -1;
      let fieldColIdx = -1;
      let colorColIdx = -1;

      headers.forEach((header, idx) => {
        const headerStr = String(header || "").trim();
        if (headerStr.includes("模板名称")) {
          nameColIdx = idx;
        } else if (headerStr.includes("字段")) {
          fieldColIdx = idx;
        } else if (headerStr.includes("标题颜色")) {
          colorColIdx = idx;
        }
      });

      if (nameColIdx === -1 || fieldColIdx === -1 || colorColIdx === -1) {
        return colorMap;
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;

        const templateName = String(row[nameColIdx] || "").trim();
        const fieldName = String(row[fieldColIdx] || "").trim();

        if (!templateName || !fieldName) continue;

        const cell = sheet.Cells(i + 1, colorColIdx + 1);
        const color = excelColorReader.getCellColor(cell);

        if (color) {
          const key = `${templateName}|${fieldName}`;
          colorMap.set(key, color);
        }
      }
    } catch (e) {
      MsgBox("读取标题颜色失败：" + e.message);
    }

    return colorMap;
  }

  /**
   * 获取默认的列配置（当没有模板或模板为空时使用）
   * @private
   * @returns {Array<Object>} 默认列配置数组
   * @description
   * 默认配置包含10个基础字段：
   * - 货号、款号、颜色、三级品类（颜色：15773696）
   * - 首次上架、状态（颜色：13434879）
   * - 成本价、白金价、到手价（颜色：10092543）
   * - 可售库存（颜色：10079487）
   *
   * 每个字段包含：field、title、width、visible、order、color
   */
  _getDefaultColumns() {
    return [
      {
        field: "itemNumber",
        title: "货号",
        width: 15,
        visible: true,
        order: 1,
        color: 15773696,
      },
      {
        field: "styleNumber",
        title: "款号",
        width: 12,
        visible: true,
        order: 2,
        color: 15773696,
      },
      {
        field: "color",
        title: "颜色",
        width: 10,
        visible: true,
        order: 3,
        color: 15773696,
      },
      {
        field: "thirdLevelCategory",
        title: "三级品类",
        width: 12,
        visible: true,
        order: 4,
        color: 15773696,
      },
      {
        field: "firstListingTime",
        title: "首次上架",
        width: 12,
        visible: true,
        order: 5,
        color: 13434879,
      },
      {
        field: "itemStatus",
        title: "状态",
        width: 10,
        visible: true,
        order: 6,
        color: 13434879,
      },
      {
        field: "costPrice",
        title: "成本价",
        width: 10,
        visible: true,
        order: 7,
        color: 10092543,
      },
      {
        field: "silverPrice",
        title: "白金价",
        width: 10,
        visible: true,
        order: 8,
        color: 10092543,
      },
      {
        field: "finalPrice",
        title: "到手价",
        width: 10,
        visible: true,
        order: 9,
        color: 10092543,
      },
      {
        field: "sellableInventory",
        title: "可售库存",
        width: 10,
        visible: true,
        order: 10,
        color: 10079487,
      },
    ];
  }

  /**
   * 从报表配置工作表加载所有模板配置
   * @returns {Map<string, Array>} 模板映射，键为模板名称，值为字段配置数组
   * @description
   * 加载流程：
   * 1. 如果已有缓存(this._templates)，直接返回缓存
   * 2. 从仓库查询所有 ReportTemplate 实体
   * 3. 获取报表配置工作表对象
   * 4. 读取每行的标题颜色映射
   * 5. 遍历所有模板项，按模板名称分组：
   *    - 过滤掉 templateName 或 fieldName 为空的项
   *    - 为每个字段添加颜色属性
   *    - 构建字段配置对象
   * 6. 按 displayOrder 对每个模板的字段进行排序
   * 7. 缓存结果并返回
   *
   * 字段配置对象属性：
   * - field: 字段名
   * - title: 列标题
   * - width: 列宽
   * - visible: 是否可见
   * - order: 显示顺序
   * - format: 数字格式
   * - description: 描述
   * - color: 标题颜色（从Excel读取）
   */
  loadTemplates() {
    if (this._templates) {
      return this._templates;
    }

    try {
      // 1.获取工作表中的模板字段对象
      const templateItems = this._repository.findAll("ReportTemplate");
      const wb = this._excelDAO.getWorkbook();
      const sheet = wb.Sheets("报表配置");

      // 2.工作表没有数据则返回空Map
      if (!templateItems || templateItems.length === 0) {
        this._templates = new Map();
        return this._templates;
      }

      // 3.获取颜色Map(模板名称|字段—>颜色)
      const rowColors = this._readRowColors(sheet);
      const templates = new Map();

      // 4.遍历所有的模板字段对象，添加颜色属性后push到数组
      templateItems.forEach((item) => {
        if (!item.templateName || !item.fieldName) return;

        if (!templates.has(item.templateName)) {
          templates.set(item.templateName, []);
        }

        // 获取字段的标题颜色
        const key = `${item.templateName}|${item.fieldName}`;
        const color = rowColors.get(key) || null;

        templates.get(item.templateName).push({
          field: item.fieldName,
          title: item.columnTitle || item.fieldName,
          width: item.columnWidth || 10,
          visible: item.isVisible === "是",
          order: item.displayOrder || 999,
          format: item.numberFormat || "",
          description: item.description || "",
          color: item.titleColor || color,
        });
      });

      // 按照显式顺序排序
      templates.forEach((columns, name) => {
        templates.set(
          name,
          columns.sort((a, b) => (a.order || 999) - (b.order || 999)),
        );
      });

      this._templates = templates;
      return templates;
    } catch (e) {
      MsgBox("报表初始化失败" + e.message);
      this._templates = new Map();
      return this._templates;
    }
  }

  /**
   * 获取所有可用的报表模板名称列表
   * @returns {string[]} 按字母顺序排序的模板名称数组
   */
  getTemplateList() {
    const templates = this.loadTemplates();
    return Array.from(templates.keys()).sort();
  }

  /**
   * 获取指定模板的字段配置
   * @param {string} templateName - 模板名称
   * @returns {Array<Object>} 字段配置数组，如果模板不存在则返回空数组
   */
  getTemplate(templateName) {
    const templates = this.loadTemplates();
    return templates.get(templateName) || [];
  }

  /**
   * 设置当前使用的报表模板
   * @param {string} templateName - 模板名称
   * @returns {boolean} 设置是否成功（模板必须存在）
   */
  setCurrentTemplate(templateName) {
    if (templateName && this.loadTemplates().has(templateName)) {
      this._currentTemplate = templateName;
      return true;
    }
    return false;
  }

  /**
   * 获取当前使用的报表模板名称
   * @returns {string|null} 当前模板名称，未设置时返回null
   */
  getCurrentTemplate() {
    return this._currentTemplate;
  }

  /**
   * 获取当前模板的字段配置
   * @returns {Array<Object>} 字段配置数组
   * @description
   * 获取逻辑：
   * 1. 如果未设置当前模板，返回默认列配置
   * 2. 如果设置了当前模板但模板为空，返回默认列配置
   * 3. 否则返回当前模板的字段配置
   */
  getCurrentColumns() {
    if (!this._currentTemplate) {
      return this._getDefaultColumns();
    }

    const columns = this.getTemplate(this._currentTemplate);
    if (columns.length === 0) {
      return this._getDefaultColumns();
    }

    return columns;
  }

  /**
   * 刷新模板缓存
   * @returns {Map<string, Array>} 刷新后的模板映射
   * @description
   * 刷新流程：
   * 1. 清除模板缓存(this._templates = null)
   * 2. 重新调用 loadTemplates 从工作表加载最新配置
   *
   * 通常在以下情况调用：
   * - 用户在UI中修改了模板配置后
   * - 需要获取最新配置时
   */
  refresh() {
    this._templates = null;
    return this.loadTemplates();
  }
}
