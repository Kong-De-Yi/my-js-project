// ============================================================================
// Excel数据访问对象
// 功能：封装所有Excel读写操作，自动识别工作簿名称
// 特点：从文件名提取品牌，无硬编码配置
// ============================================================================

class ExcelDAO {
  constructor() {
    this._config = DataConfig.getInstance();
    this._workbookName = this._detectWorkbookName();
  }

  // 从当前活动工作簿或文件名识别工作簿名称
  _detectWorkbookName() {
    const appName = this._config.getAppName();
    try {
      // 1. 尝试获取当前活动工作簿
      const activeWb = ActiveWorkbook;
      if (activeWb?.Name) {
        // 验证文件名格式：必须包含"商品运营表"
        if (!activeWb.Name.includes(appName)) {
          throw new Error(
            `工作簿名称必须包含"${appName}"，当前名称：${activeWb.Name}`,
          );
        }
        return activeWb.Name;
      }
    } catch (e) {
      // 忽略，继续尝试其他方法
    }

    // 2. 遍历所有打开的工作簿
    try {
      for (let i = 1; i <= Workbooks.Count; i++) {
        const wb = Workbooks(i);
        if (wb.Name.includes(appName)) {
          return wb.Name;
        }
      }
    } catch (e) {
      // 忽略
    }

    throw new Error(`未找到包含"${appName}"的工作簿，请先打开文件`);
  }

  // 从工作簿名称提取品牌
  extractBrand() {
    const match = this._workbookName.match(/【(.*?)】/);
    if (!match) {
      throw new Error(`无法从工作簿名称中提取品牌：${this._workbookName}`);
    }
    return match[1];
  }

  // 获取工作薄名称
  getWorkbookName() {
    return this._workbookName;
  }

  // 获取工作薄
  getWorkbook() {
    try {
      return Workbooks(this._workbookName);
    } catch (e) {
      throw new Error(`找不到工作簿：${this._workbookName}`);
    }
  }

  // 读取工作表数据
  read(entityName, wsName = null) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    wsName = wsName || entityConfig.worksheet;
    const fields = entityConfig.fields;

    // 获取需要持久化的字段映射
    const keyToTitle = {};
    Object.entries(fields).forEach(([key, config]) => {
      keyToTitle[key] = config.title || key;
    });

    // 读取工作表
    let data = [];
    try {
      const sheet = this.getWorkbook().Sheets(wsName);
      const usedRange = sheet.UsedRange;

      if (!usedRange || usedRange.Value2 == null) {
        return [];
      }

      data = usedRange.Value2;
    } catch (e) {
      throw new Error(`读取工作表【${wsName}】失败：${e.message}`);
    }

    // 过滤空行
    data = data.filter(
      (row) =>
        row && row.some((cell) => cell != null && String(cell).trim() !== ""),
    );

    if (data.length === 0) {
      throw new Error(`【${wsName}】中没有任何数据`);
    }

    // 标题行
    const titleRow = data.shift();

    // 构建列索引
    const columnIndex = {};
    Object.entries(keyToTitle).forEach(([key, title]) => {
      const index = titleRow.findIndex((t) => String(t).trim() === title);
      if (index === -1) {
        throw new Error(`【${wsName}】中找不到列：${title}`);
      }
      columnIndex[key] = index;
    });

    // 转换数据
    const results = [];
    data.forEach((row, idx) => {
      const obj = {};

      Object.keys(columnIndex).forEach((key) => {
        const colIdx = columnIndex[key];
        const rawValue = row[colIdx];
        const fieldConfig = fields[key];

        // 类型转换
        switch (fieldConfig?.type) {
          case "number":
            obj[key] = this._toNumber(rawValue);
            break;
          case "date":
            obj[key] = this._toDate(rawValue);
            break;
          case "string":
            obj[key] = this._toString(rawValue);
            break;
          default:
            obj[key] = this._toString(rawValue);
        }
      });

      // 记录行号
      obj._rowNumber = idx + 2;

      results.push(obj);
    });

    return results;
  }

  // 写入工作表数据
  write(entityName, data, targetWorkbook = null) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    const wsName = entityConfig.worksheet;
    const fields = entityConfig.fields;

    // 获取需要持久化的字段
    const persistFields = [];
    const keyToTitle = {};

    Object.entries(fields).forEach(([key, config]) => {
      if (config.persist !== false) {
        persistFields.push(key);
        keyToTitle[key] = config.title || key;
      }
    });

    // 构建输出数据
    const outputData = [];

    // 标题行
    outputData.push(persistFields.map((key) => keyToTitle[key]));

    // 数据行
    data.forEach((item) => {
      const row = persistFields.map((key) => {
        const value = item[key];
        const fieldConfig = fields[key];

        if (
          value == undefined ||
          String(value).trim() === "" ||
          typeof value === "boolean"
        ) {
          return "";
        }

        // 格式化
        switch (fieldConfig?.type) {
          case "number":
            return Number(value);
          case "date":
            return this._formatDate(value);
          case "string":
            return String(value);
          default:
            return String(value);
        }
      });
      outputData.push(row);
    });

    // 写入Excel
    const wb = targetWorkbook || this.getWorkbook();
    const sheet = wb.Sheets(wsName);

    sheet.Cells.ClearContents();

    if (outputData.length > 0) {
      const range = sheet
        .Range("A1")
        .Resize(outputData.length, outputData[0].length);
      range.Value2 = outputData;
    }

    // 自动保存
    if (!targetWorkbook) {
      wb.Save();
    }
  }

  // 清空工作表
  clear(entityName, targetWorkbook = null) {
    const entityConfig = this._config.get(entityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${entityName}`);
    }

    const wb = targetWorkbook || this.getWorkbook();
    wb.Sheets(entityConfig.worksheet).Cells.Clear();

    if (!targetWorkbook) {
      wb.Save();
    }
  }

  // 复制工作表
  copySheet(sourceEntityName, targetWorkbook) {
    const entityConfig = this._config.get(sourceEntityName);
    if (!entityConfig) {
      throw new Error(`未知实体：${sourceEntityName}`);
    }

    const sourceWb = this.getWorkbook();
    sourceWb
      .Sheets(entityConfig.worksheet)
      .Copy(null, targetWorkbook.Sheets(targetWorkbook.Sheets.Count));
    return ActiveSheet;
  }

  // 对外提供格式化日期服务
  formatDate(value) {
    return this._formatDate(value);
  }

  // 对外提供日期转化服务
  parseDate(dateStr) {
    return this._parseDate(dateStr);
  }

  _parseDate(dateStr) {
    if (!dateStr) return null;

    const timestamp = Date.parse(String(dateStr));
    if (isNaN(timestamp)) return null;

    return new Date(timestamp);
  }

  // 转换为数字
  _toNumber(value) {
    if (
      value == null ||
      String(value).trim() === "" ||
      typeof value === "boolean"
    ) {
      return undefined;
    }

    const num = Number(value);
    return isFinite(num) ? num : undefined;
  }

  // 转换为字符串
  _toString(value) {
    if (
      value == null ||
      String(value).trim() === "" ||
      typeof value === "boolean"
    ) {
      return undefined;
    }

    return String(value);
  }

  // 转换为日期
  _toDate(value) {
    if (
      value == null ||
      String(value).trim() === "" ||
      typeof value === "boolean"
    ) {
      return undefined;
    }

    const cleanValue = String(value).replace(/^'/, "");

    const date = this._parseDate(cleanValue);
    if (!date) return undefined;

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  // 格式化日期
  _formatDate(value) {
    const date = this._parseDate(value);
    if (!date) return "";

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");

    return `'${year}-${month}-${day}`;
  }
}
