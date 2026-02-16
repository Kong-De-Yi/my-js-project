// ============================================================================
// ExcelColorReader.js - Excel颜色读取器（使用可选链操作符）
// 功能：读取Excel单元格的背景色，支持各种颜色格式
// ============================================================================

class ExcelColorReader {
  /**
   * 读取单元格的背景色
   * @param {Object} cell - Excel单元格对象
   * @returns {number|null} VBA颜色值，没有颜色返回null
   */
  static getCellColor(cell) {
    try {
      // 使用可选链操作符简化判断
      if (!cell?.Interior) return null;

      const color = cell.Interior.Color;

      // Excel默认无填充的颜色值

      // 16777215 = 白色 (RGB 255,255,255)
      // 0 = 黑色 (可能表示无颜色)
      // -4142 = xlNone (无填充)
      if (color === 16777215 || color === 0 || color === -4142) {
        return null;
      }

      return color;
    } catch (e) {
      return null;
    }
  }

  /**
   * 读取单元格的RGB值（如果需要）
   * @param {Object} cell - Excel单元格对象
   * @returns {Object|null} RGB对象 {r, g, b}
   */
  static getCellRGB(cell) {
    try {
      const color = this.getCellColor(cell);
      if (!color) return null;

      // 将VBA颜色值转换为RGB
      // 注意：VBA中颜色值计算公式：Color = Blue * 65536 + Green * 256 + Red
      const r = color % 256;
      const g = Math.floor(color / 256) % 256;
      const b = Math.floor(color / 65536) % 256;

      return { r, g, b };
    } catch (e) {
      return null;
    }
  }

  /**
   * 读取工作表标题行的颜色配置
   * @param {Object} sheet - Excel工作表
   * @param {number} headerRow - 标题行号（默认为1）
   * @returns {Map} 列索引到颜色的映射
   */
  static readHeaderColors(sheet, headerRow = 1) {
    const colorMap = new Map();

    try {
      // 使用可选链操作符
      if (!sheet?.UsedRange) return colorMap;

      const usedRange = sheet.UsedRange;
      const lastColumn = usedRange.Columns.Count;

      for (let col = 1; col <= lastColumn; col++) {
        const cell = sheet.Cells(headerRow, col);
        const color = this.getCellColor(cell);

        if (color) {
          colorMap.set(col, color);
        }
      }
    } catch (e) {
      // 忽略读取错误
    }

    return colorMap;
  }

  /**
   * 将颜色应用到单元格范围
   * @param {Object} range - Excel单元格范围
   * @param {number} color - VBA颜色值
   */
  static applyColorToRange(range, color) {
    // 使用可选链操作符
    if (!range || !color) return;

    try {
      range.Interior.Color = color;
    } catch (e) {
      // 忽略应用错误
    }
  }

  /**
   * 将颜色应用到整列
   * @param {Object} sheet - Excel工作表
   * @param {number} column - 列号
   * @param {number} color - VBA颜色值
   * @param {number} startRow - 起始行号（默认为1）
   */
  static applyColorToColumn(sheet, column, color, startRow = 1) {
    // 使用可选链操作符
    if (!sheet || !color) return;

    try {
      // 使用可选链操作符
      const lastRow = sheet.UsedRange?.Rows.Count || 1;
      const range = sheet.Range(
        sheet.Cells(startRow, column),
        sheet.Cells(lastRow, column),
      );
      range.Interior.Color = color;
    } catch (e) {
      // 忽略应用错误
    }
  }

  /**
   * 将颜色应用到整行
   * @param {Object} sheet - Excel工作表
   * @param {number} row - 行号
   * @param {number} color - VBA颜色值
   * @param {number} startColumn - 起始列号（默认为1）
   * @param {number} endColumn - 结束列号（可选）
   */
  static applyColorToRow(sheet, row, color, startColumn = 1, endColumn = null) {
    // 使用可选链操作符
    if (!sheet || !color) return;

    try {
      // 使用可选链操作符
      const lastColumn = endColumn || sheet.UsedRange?.Columns.Count || 1;
      const range = sheet.Range(
        sheet.Cells(row, startColumn),
        sheet.Cells(row, lastColumn),
      );
      range.Interior.Color = color;
    } catch (e) {
      // 忽略应用错误
    }
  }

  /**
   * 获取颜色的友好名称（用于调试）
   * @param {number} color - VBA颜色值
   * @returns {string} 颜色名称
   */
  static getColorName(color) {
    const colorMap = {
      255: "红色",
      65535: "黄色",
      65280: "绿色",
      16711680: "蓝色",
      16711935: "紫色",
      65535: "青色",
      8421504: "灰色",
      0: "黑色",
      16777215: "白色",
      12632256: "灰色",
      10079487: "浅蓝",
      13434879: "浅黄",
      10092543: "浅绿",
    };

    return colorMap[color] || `颜色(${color})`;
  }

  /**
   * 测试颜色是否可见（用于调试）
   * @param {Object} cell - Excel单元格
   * @returns {boolean} 是否有颜色
   */
  static hasColor(cell) {
    return this.getCellColor(cell) !== null;
  }

  /**
   * 清除单元格颜色
   * @param {Object} range - Excel单元格范围
   */
  static clearColor(range) {
    if (!range) return;

    try {
      range.Interior.Color = -4142; // xlNone
    } catch (e) {
      // 忽略错误
    }
  }
}

// 导出单例
const excelColorReader = ExcelColorReader;
