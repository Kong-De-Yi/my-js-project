class ReportTemplateManager {
  constructor(repository, excelDAO) {
    this._repository = repository;
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();
    this._templates = null;
    this._currentTemplate = null;
  }

  initializeDefaultTemplates() {
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

  // 创建报表模板并保存至报表配置工作表
  _createDefaultTemplate(templateName, columns) {
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

    let existingTemplates = [];
    try {
      existingTemplates = this._repository.findAll("ReportTemplate");
    } catch (e) {
      existingTemplates = [];
    }

    // 剔除已配置的报表模板
    const filtered = existingTemplates.filter(
      (t) => t.templateName !== templateName,
    );
    const allTemplates = [...filtered, ...templateData];

    this._repository.save("ReportTemplate", allTemplates);
  }

  // 从报表配置工作表载入模板配置项目，返回Map(模板名称——>配置对象数组)
  loadTemplates() {
    if (this._templates) {
      return this._templates;
    }

    try {
      const templateItems = this._repository.findAll("ReportTemplate");
      const wb = this._excelDAO.getWorkbook();
      const sheet = wb.Sheets("报表配置");

      if (!templateItems || templateItems.length === 0) {
        this._templates = new Map();
        return this._templates;
      }

      const rowColors = this._readRowColors(sheet, templateItems);
      const templates = new Map();

      templateItems.forEach((item) => {
        if (!item.templateName || !item.fieldName) return;

        if (!templates.has(item.templateName)) {
          templates.set(item.templateName, []);
        }

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
          color: color,
        });
      });

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

  // 返回报表配置的模板项目标题颜色Map(模板名称|字段名称——>颜色)
  _readRowColors(sheet, templateItems) {
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

  //  返回所有的报表模板名称数组
  getTemplateList() {
    const templates = this.loadTemplates();
    return Array.from(templates.keys()).sort();
  }

  // 获取指定的报表模板的报表项目，返回数组
  getTemplate(templateName) {
    const templates = this.loadTemplates();
    return templates.get(templateName) || [];
  }

  // 设置当前的报表模板
  setCurrentTemplate(templateName) {
    if (templateName && this.loadTemplates().has(templateName)) {
      this._currentTemplate = templateName;
      return true;
    }
    return false;
  }

  // 获取当前的报表模板
  getCurrentTemplate() {
    return this._currentTemplate;
  }

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

  refresh() {
    this._templates = null;
    return this.loadTemplates();
  }
}
