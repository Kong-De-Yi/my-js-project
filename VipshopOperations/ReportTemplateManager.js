// ============================================================================
// 报表模板管理器
// 功能：读取报表配置工作表中的模板和颜色设置
// 特点：直接从Excel单元格读取背景色，无需额外配置
// ============================================================================

class ReportTemplateManager {
  constructor(repository, excelDAO) {
    this._repository = repository;
    this._excelDAO = excelDAO;
    this._config = dataConfig;
    this._templates = null;
    this._currentTemplate = null;
  }

  /**
   * 初始化默认模板
   */
  initializeDefaultTemplates() {
    const templates = this.loadTemplates();

    if (templates.size > 0) {
      return templates;
    }

    // 创建默认模板
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
      {
        field: "salesQuantityOfLast7Days",
        title: "近7天销量",
        width: 10,
        order: 11,
      },
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

    this._createDefaultTemplate("销售排行报表", [
      { field: "itemNumber", title: "货号", width: 15, order: 1 },
      { field: "styleNumber", title: "款号", width: 12, order: 2 },
      { field: "color", title: "颜色", width: 10, order: 3 },
      { field: "thirdLevelCategory", title: "三级品类", width: 12, order: 4 },
      {
        field: "salesQuantityOfLast7Days",
        title: "近7天销量",
        width: 12,
        order: 5,
      },
      {
        field: "salesAmountOfLast7Days",
        title: "近7天销售额",
        width: 12,
        order: 6,
      },
      { field: "unitPriceOfLast7Days", title: "件单价", width: 10, order: 7 },
      { field: "styleSalesOfLast7Days", title: "款销量", width: 10, order: 8 },
      { field: "totalSales", title: "销量总计", width: 10, order: 9 },
    ]);

    return this.loadTemplates();
  }

  /**
   * 创建默认模板
   */
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

    // 读取现有模板
    let existingTemplates = [];
    try {
      existingTemplates = this._repository.findAll("ReportTemplate");
    } catch (e) {
      existingTemplates = [];
    }

    // 过滤掉同名的旧模板
    const filtered = existingTemplates.filter(
      (t) => t.templateName !== templateName,
    );
    const allTemplates = [...filtered, ...templateData];

    this._repository.save("ReportTemplate", allTemplates);
  }

  /**
   * 加载所有模板（包含颜色信息）
   */
  loadTemplates() {
    if (this._templates) {
      return this._templates;
    }

    try {
      const templateItems = this._repository.findAll("ReportTemplate");
      const wb = this._excelDAO.getWorkbook();
      const sheet = wb.Sheets("报表配置");

      // 读取标题行的颜色
      const headerColors = excelColorReader.readHeaderColors(sheet, 1);

      // 按模板名称分组
      const templates = new Map();

      templateItems.forEach((item) => {
        if (!item.templateName || !item.fieldName) return;

        if (!templates.has(item.templateName)) {
          templates.set(item.templateName, []);
        }

        templates.get(item.templateName).push({
          field: item.fieldName,
          title: item.columnTitle || item.fieldName,
          width: item.columnWidth || 10,
          visible: item.isVisible === "是",
          order: item.displayOrder || 999,
          format: item.numberFormat || "",
          description: item.description || "",
        });
      });

      // 对每个模板的列按顺序排序
      templates.forEach((columns, name) => {
        templates.set(
          name,
          columns.sort((a, b) => (a.order || 999) - (b.order || 999)),
        );
      });

      // 存储颜色信息（每个模板可能有不同的颜色配置）
      this._templateColors = new Map();

      // 根据报表配置工作表的实际列位置，将颜色关联到模板
      if (sheet) {
        const usedRange = sheet.UsedRange;
        if (usedRange) {
          const data = usedRange.Value2;
          if (data && data.length > 0) {
            const headers = data[0];

            // 找到各列的索引
            const nameColIdx = headers.findIndex((h) =>
              String(h).includes("模板名称"),
            );
            const fieldColIdx = headers.findIndex((h) =>
              String(h).includes("字段"),
            );

            if (nameColIdx !== -1 && fieldColIdx !== -1) {
              // 遍历每一行，记录每个模板的字段颜色
              for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (!row || !row[nameColIdx]) continue;

                const templateName = String(row[nameColIdx]).trim();
                const fieldName = String(row[fieldColIdx] || "").trim();

                if (!templateName || !fieldName) continue;

                // 读取该行对应字段所在列的颜色
                const cell = sheet.Cells(i + 1, fieldColIdx + 1);
                const color = excelColorReader.getCellColor(cell);

                if (color) {
                  if (!this._templateColors.has(templateName)) {
                    this._templateColors.set(templateName, new Map());
                  }
                  this._templateColors.get(templateName).set(fieldName, color);
                }
              }
            }
          }
        }
      }

      this._templates = templates;
      return templates;
    } catch (e) {
      return new Map();
    }
  }

  /**
   * 获取模板的字段颜色
   */
  getFieldColor(templateName, fieldName) {
    if (!this._templateColors) return null;

    const templateColors = this._templateColors.get(templateName);
    return templateColors?.get(fieldName) || null;
  }

  /**
   * 获取模板列表
   */
  getTemplateList() {
    const templates = this.loadTemplates();
    return Array.from(templates.keys()).sort();
  }

  /**
   * 获取模板配置
   */
  getTemplate(templateName) {
    const templates = this.loadTemplates();
    const columns = templates.get(templateName) || [];

    // 为每列附加颜色信息
    return columns.map((col) => ({
      ...col,
      color: this.getFieldColor(templateName, col.field),
    }));
  }

  /**
   * 设置当前模板
   */
  setCurrentTemplate(templateName) {
    if (templateName && this.loadTemplates().has(templateName)) {
      this._currentTemplate = templateName;
      return true;
    }
    return false;
  }

  /**
   * 获取当前模板
   */
  getCurrentTemplate() {
    return this._currentTemplate;
  }

  /**
   * 获取当前模板的列配置
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
   * 获取默认列配置
   */
  _getDefaultColumns() {
    return [
      {
        field: "itemNumber",
        title: "货号",
        width: 15,
        visible: true,
        order: 1,
      },
      {
        field: "styleNumber",
        title: "款号",
        width: 12,
        visible: true,
        order: 2,
      },
      { field: "color", title: "颜色", width: 10, visible: true, order: 3 },
      {
        field: "thirdLevelCategory",
        title: "三级品类",
        width: 12,
        visible: true,
        order: 4,
      },
      {
        field: "firstListingTime",
        title: "首次上架",
        width: 12,
        visible: true,
        order: 5,
      },
      {
        field: "itemStatus",
        title: "状态",
        width: 10,
        visible: true,
        order: 6,
      },
      {
        field: "costPrice",
        title: "成本价",
        width: 10,
        visible: true,
        order: 7,
      },
      {
        field: "silverPrice",
        title: "白金价",
        width: 10,
        visible: true,
        order: 8,
      },
      {
        field: "finalPrice",
        title: "到手价",
        width: 10,
        visible: true,
        order: 9,
      },
      {
        field: "sellableInventory",
        title: "可售库存",
        width: 10,
        visible: true,
        order: 10,
      },
    ];
  }

  /**
   * 刷新模板
   */
  refresh() {
    this._templates = null;
    this._templateColors = null;
    return this.loadTemplates();
  }
}
