/**
 * Excel 数据访问对象 (Data Access Object)
 * @class ExcelDAO
 * @description 负责与 Excel 工作簿进行底层交互，封装了所有对工作簿和工作表的读取、写入、清空和复制操作。
 * 它是系统中唯一直接操作 Excel 对象的层，为上层（如 Repository）提供数据持久化能力。
 * 该类采用单例模式，确保全局只有一个实例。
 *
 * @example
 * // 获取 ExcelDAO 实例
 * const excelDAO = ExcelDAO.getInstance();
 *
 * // 读取 "Product" 实体的数据
 * const products = excelDAO.read("Product");
 *
 * // 将处理后的数据写回 "Product" 工作表
 * excelDAO.write("Product", updatedProducts);
 */
class ExcelDAO {
  /** @type {ExcelDAO} 单例实例 */
  static _instance = null;

  /**
   * 创建 ExcelDAO 实例。
   * 私有构造函数，防止外部使用 `new` 创建实例。
   * @private
   */
  constructor() {
    if (ExcelDAO._instance) {
      return ExcelDAO._instance;
    }

    this._config = DataConfig.getInstance();
    this._workbookName = this._detectWorkbookName();
    this._converter = Converter.getInstance();

    ExcelDAO._instance = this;
  }

  /**
   * 获取 ExcelDAO 的单例实例。
   * @static
   * @returns {ExcelDAO} ExcelDAO 的单例实例。
   */
  static getInstance() {
    if (!ExcelDAO._instance) {
      ExcelDAO._instance = new ExcelDAO();
    }
    return ExcelDAO._instance;
  }

  /**
   * 自动检测并返回符合要求的工作簿名称。
   * @private
   * @returns {string} 找到的工作簿名称。
   * @throws {Error} 如果未找到包含配置的应用名称（如“商品运营表”）的工作簿，则抛出错误。
   * @description
   * 检测策略如下：
   * 1. 优先检查当前活动工作簿 (`ActiveWorkbook`)，验证其名称是否包含配置的应用名称。
   * 2. 如果活动工作簿不符合要求，则遍历所有已打开的工作簿，查找名称中包含应用名称的工作簿。
   * 3. 如果都未找到，则抛出错误。
   */
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

  /**
   * 从工作簿名称中提取品牌信息。
   * @returns {string} 提取出的品牌名称（位于方括号【】中的内容）。
   * @throws {Error} 如果工作簿名称中不包含格式正确的品牌信息（如“【某品牌】”），则抛出错误。
   *
   * @example
   * // 假设工作簿名为 "2024商品运营表【ABC品牌】.xlsx"
   * excelDAO.extractBrand(); // 返回 "ABC品牌"
   */
  extractBrand() {
    const match = this._workbookName.match(/【(.*?)】/);
    if (!match) {
      throw new Error(`无法从工作簿名称中提取品牌：${this._workbookName}`);
    }
    return match[1];
  }

  /**
   * 获取当前工作簿的名称。
   * @returns {string} 工作簿名称。
   */
  getWorkbookName() {
    return this._workbookName;
  }

  /**
   * 获取当前工作簿的 Excel 对象。
   * @returns {Excel.Workbook} Excel 工作簿对象。
   * @throws {Error} 如果根据存储的名称找不到对应的工作簿，则抛出错误。
   */
  getWorkbook() {
    try {
      return Workbooks(this._workbookName);
    } catch (e) {
      throw new Error(`找不到工作簿：${this._workbookName}`);
    }
  }

  /**
   * 从指定的工作表读取实体数据。
   * @param {string} entityName - 实体名称，用于从配置中获取字段映射和工作表名称。
   * @param {string} [wsName] - 可选，要读取的工作表名称。如果不提供，则使用实体配置中的 `worksheet` 属性。
   * @returns {Object[]} 返回一个对象数组，每个对象对应数据表中的一行，键为字段名，值为经过类型转换后的数据。
   * 每个对象还会包含一个特殊的 `_rowNumber` 属性，表示该行数据在 Excel 中的实际行号（从2开始，因为第1行是标题）。
   * @throws {Error} 如果实体不存在、工作表读取失败、工作表中无数据或找不到配置的列，则抛出错误。
   *
   * @description
   * 读取流程如下：
   * 1. 根据 `entityName` 获取实体配置。
   * 2. 获取目标工作表及其已使用区域 (`UsedRange`)。
   * 3. 过滤掉完全为空的行。
   * 4. 将第一行作为标题行，并根据实体配置中 `type !== "computed"` 的字段，建立字段名到列索引的映射。
   * 5. 遍历数据行，根据字段配置的 `type`（如 'number', 'date'）使用 `Converter` 进行类型转换。
   * 6. 如果转换后值为 `undefined` 且字段配置了 `default`，则应用默认值。
   * 7. 为计算字段 (`type: "computed"`) 预留 `undefined` 占位符。
   * 8. 为每行数据添加 `_rowNumber` 属性。
   */
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
      if (config.type !== "computed") {
        keyToTitle[key] = config.title || key;
      }
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
            obj[key] = this._converter.toNumber(rawValue);
            break;
          case "date":
            obj[key] = this._converter.toDateStr(rawValue);
            break;
          default:
            obj[key] = this._converter.toString(rawValue);
        }

        // 处理默认值
        if (obj[key] === undefined && fieldConfig?.default !== undefined) {
          obj[key] = fieldConfig.default;
        }
      });

      //  添加计算字段
      Object.entries(fields).forEach(([key, config]) => {
        if (config.type === "computed") {
          obj[key] = undefined;
        }
      });

      // 记录行号
      obj._rowNumber = idx + 2;

      results.push(obj);
    });

    return results;
  }

  /**
   * 将实体数据写入指定的工作表。
   * @param {string} entityName - 实体名称。
   * @param {Object[]} data - 要写入的数据对象数组。
   * @param {Excel.Workbook} [targetWorkbook] - 可选，目标工作簿。如果不提供，则写入当前工作簿。
   * @throws {Error} 如果实体不存在，则抛出错误。
   *
   * @description
   * 写入流程如下：
   * 1. 根据 `entityName` 获取实体配置。
   * 2. 确定需要持久化的字段 (`config.persist !== false`)，并获取它们的标题 (`title`)。
   * 3. 构建输出数据数组，第一行为标题行。
   * 4. 遍历输入的数据数组，为每一行构建数据行。对于每个字段：
   *    - 值为 `null`、`undefined`、空字符串或布尔值时，输出 `undefined`（Excel 中将显示为空单元格）。
   *    - 如果值等于字段配置的 `default` 值，也输出 `undefined`（避免写入默认值，保持整洁）。
   *    - 根据字段配置的 `type` 对值进行格式化（如数字转 Number，日期用 `Converter.formatDate` 格式化）。
   * 5. 清空目标工作表的原内容 (`ClearContents`)。
   * 6. 将构建好的二维数组一次性写入工作表，从 A1 单元格开始。
   * 7. 如果未指定 `targetWorkbook`，则自动保存当前工作簿。
   */
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
          value == null ||
          String(value).trim() === "" ||
          typeof value === "boolean"
        ) {
          return undefined;
        }

        // 处理默认值
        if (
          fieldConfig?.default !== undefined &&
          value === fieldConfig.default
        ) {
          return undefined;
        }

        // 格式化
        switch (fieldConfig?.type) {
          case "number":
            return Number(value);
          case "date":
            return this._converter.formatDate(value);
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

  /**
   * 清空指定实体对应的工作表。
   * @param {string} entityName - 实体名称。
   * @param {Excel.Workbook} [targetWorkbook] - 可选，目标工作簿。如果不提供，则操作当前工作簿。
   * @throws {Error} 如果实体不存在，则抛出错误。
   * @description
   * 此方法会清除工作表中的所有内容和格式 (`Cells.Clear()`)。
   * 如果未指定 `targetWorkbook`，则自动保存当前工作簿。
   */
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

  /**
   * 将源实体对应的工作表复制到目标工作簿的末尾。
   * @param {string} sourceEntityName - 源实体名称。
   * @param {Excel.Workbook} targetWorkbook - 目标工作簿。
   * @returns {Excel.Worksheet} 新复制的工作表对象（复制后将成为活动工作表）。
   * @throws {Error} 如果源实体不存在，则抛出错误。
   *
   * @example
   * // 将当前工作簿中的 "Product" 模板表复制到一个新工作簿
   * const newWb = Workbooks.Add();
   * const newSheet = excelDAO.copySheet("Product", newWb);
   * newSheet.Name = "生成报表"; // 重命名新工作表
   */
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
}
