// ============================================================================
// Excel颜色读取器
// 功能：读取Excel单元格的背景色，无需配置RGB值
// 特点：用户直接在Excel中填充颜色，程序自动读取
// ============================================================================

class ExcelColorReader {
  /**
   * 读取单元格的背景色
   * @param {Object} cell - Excel单元格对象
   * @returns {number|null} VBA颜色值，没有颜色返回null
   */
  static getCellColor(cell) {
    try {
      if (!cell || !cell.Interior) return null;

      const color = cell.Interior.Color;
      // Excel默认无填充的颜色值
      if (color === 16777215 || color === 0 || color === -4142) {
        return null;
      }
      return color;
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
      const usedRange = sheet.UsedRange;
      if (!usedRange) return colorMap;

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
   */
  static applyColorToRange(range, color) {
    if (!range || !color) return;

    try {
      range.Interior.Color = color;
    } catch (e) {
      // 忽略应用错误
    }
  }

  /**
   * 将颜色应用到整列
   */
  static applyColorToColumn(sheet, column, color, startRow = 1) {
    if (!sheet || !color) return;

    try {
      const lastRow = sheet.UsedRange.Rows.Count;
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
   * 获取颜色的友好名称
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
}

// 导出
const excelColorReader = ExcelColorReader;
