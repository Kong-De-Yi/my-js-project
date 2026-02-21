// ============================================================================
// SalesStatisticsService.js - 销售统计服务（增强版）
// 功能：提供各种销售统计指标的计算，支持按年/月/周/日维度
// ============================================================================

class SalesStatisticsService {
  /**
   * 构造函数
   * @param {Object} repository - 数据仓库实例
   */
  constructor(repository) {
    this._repository = repository;
    this._salesCache = null;
    this._currentDate = new Date();
  }

  /**
   * 刷新销售数据缓存
   * @private
   */
  _refreshSalesCache() {
    this._salesCache = this._repository.findAll("ProductSales");
  }

  /**
   * 按货号分组销售数据
   * @returns {Map} 货号到销售数据数组的映射
   * @private
   */
  _groupSalesByItemNumber() {
    if (!this._salesCache) {
      this._refreshSalesCache();
    }

    const grouped = new Map();

    this._salesCache.forEach((sale) => {
      if (!sale.itemNumber) return;

      if (!grouped.has(sale.itemNumber)) {
        grouped.set(sale.itemNumber, []);
      }
      grouped.get(sale.itemNumber).push(sale);
    });

    return grouped;
  }

  /**
   * 按货号和日期范围过滤销售数据
   * @param {Array} sales - 销售数据数组
   * @param {Date} startDate - 开始日期
   * @param {Date} endDate - 结束日期
   * @returns {Array} 过滤后的销售数据
   * @private
   */
  _filterSalesByDateRange(sales, startDate, endDate) {
    const start = startDate.getTime();
    const end = endDate.getTime();

    return sales.filter((sale) => {
      const saleDate = _validationEngine.parseDate(sale.salesDate);
      if (!saleDate) return false;
      const saleTime = saleDate.getTime();
      return saleTime >= start && saleTime <= end;
    });
  }

  /**
   * 计算销售数据的总和
   * @param {Array} sales - 销售数据数组
   * @param {string} field - 字段名
   * @returns {number} 总和
   * @private
   */
  _sumField(sales, field) {
    return sales.reduce((sum, sale) => {
      return sum + (Number(sale[field]) || 0);
    }, 0);
  }

  /**
   * 获取指定货号在指定年份的全年销量（优化版-使用索引）
   * @param {string} itemNumber - 货号
   * @param {number} year - 年份
   * @returns {number} 年销量
   */
  getYearTotalSales(itemNumber, year) {
    if (!itemNumber) return 0;

    // 使用索引查询
    const sales = this._repository.find("ProductSales", {
      itemNumber: itemNumber,
      salesYear: year,
    });

    return this._sumField(sales, "salesQuantity");
  }

  /**
   * 获取指定货号在指定年份和月份的销量（优化版-使用索引）
   * @param {string} itemNumber - 货号
   * @param {number} year - 年份
   * @param {number} month - 月份（1-12）
   * @returns {number} 月销量
   */
  getMonthSales(itemNumber, year, month) {
    if (!itemNumber) return 0;

    // 使用索引查询
    const sales = this._repository.find("ProductSales", {
      itemNumber: itemNumber,
      salesYear: year,
      salesMonth: month,
    });

    return this._sumField(sales, "salesQuantity");
  }

  /**
   * 获取指定货号在指定年份和周数的销量（优化版-使用索引）
   * @param {string} itemNumber - 货号
   * @param {number} year - 年份
   * @param {number} week - 周数（1-53）
   * @returns {number} 周销量
   */
  getWeekSales(itemNumber, year, week) {
    if (!itemNumber) return 0;

    // 使用索引查询
    const sales = this._repository.find("ProductSales", {
      itemNumber: itemNumber,
      salesYear: year,
      salesWeekOfYear: week,
    });

    return this._sumField(sales, "salesQuantity");
  }

  /**
   * 获取指定货号在指定日期的销量（优化版-使用主键）
   * @param {string} itemNumber - 货号
   * @param {Date} date - 日期
   * @returns {number} 日销量
   */
  getDaySales(itemNumber, date) {
    if (!itemNumber) return 0;

    const dateStr = _validationEngine.formatDate(date);

    const sale = this._repository.findOne("ProductSales", {
      itemNumber: itemNumber,
      salesDate: dateStr,
    });

    return sale ? Number(sale.salesQuantity) || 0 : 0;
  }

  /**
   * 获取近N天的指定字段总和（优化版-使用daysSinceSale索引）
   * @param {string} itemNumber - 货号
   * @param {number} days - 天数
   * @param {string} field - 字段名
   * @returns {number} 总和
   */
  getLastNDaysSum(itemNumber, days, field) {
    if (!itemNumber) return 0;

    const sales = this._repository.query("ProductSales", {
      filter: { itemNumber: itemNumber, daysSinceSale: { $lte: days } },
    });

    return this._sumField(sales, field);
  }

  /**
   * 获取近N天的每日销量
   * @param {string} itemNumber - 货号
   * @param {number} days - 天数
   * @returns {Array} 每日销量数组
   */
  getLastNDaysDailySales(itemNumber, days) {
    const result = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = days - 1; i >= 0; i--) {
      const date = new Date(today);
      date.setDate(today.getDate() - i);
      const sales = this.getDaySales(itemNumber, date);

      result.push({
        date: date,
        dateStr: _validationEngine.formatDate(date),
        sales: sales,
      });
    }

    return result;
  }

  /**
   * 获取当前日期
   * @returns {Date} 当前日期
   */
  getCurrentDate() {
    return this._currentDate;
  }

  /**
   * 获取当前年份
   * @returns {number} 当前年份
   */
  getCurrentYear() {
    return this._currentDate.getFullYear();
  }

  /**
   * 获取当前月份
   * @returns {number} 当前月份（1-12）
   */
  getCurrentMonth() {
    return this._currentDate.getMonth() + 1;
  }

  /**
   * 获取当前日期
   * @returns {number} 当前日期
   */
  getCurrentDay() {
    return this._currentDate.getDate();
  }

  /**
   * 获取当前周数
   * @returns {number} 当前周数
   */
  getCurrentWeek() {
    return this._getISOWeekNumber(this._currentDate);
  }

  /**
   * 获取前年、去年、今年的年份
   * @returns {Object} 年份对象
   */
  getYearRange() {
    const currentYear = this.getCurrentYear();
    return {
      beforeLast: currentYear - 2,
      last: currentYear - 1,
      current: currentYear,
    };
  }

  /**
   * 获取指定年份的月份范围
   * @param {number} year - 年份
   * @returns {Array} 月份数组
   */
  getMonthsOfYear(year) {
    const currentYear = this.getCurrentYear();
    const currentMonth = this.getCurrentMonth();

    if (year === currentYear) {
      return Array.from({ length: currentMonth }, (_, i) => i + 1);
    } else {
      return Array.from({ length: 12 }, (_, i) => i + 1);
    }
  }

  /**
   * 获取指定年份的周数范围
   * @param {number} year - 年份
   * @returns {Array} 周数数组
   */
  getWeeksOfYear(year) {
    const currentYear = this.getCurrentYear();
    const currentWeek = this.getCurrentWeek();

    if (year === currentYear) {
      return Array.from({ length: currentWeek }, (_, i) => i + 1);
    } else {
      const lastDay = new Date(year, 11, 31);
      const lastWeek = this._getISOWeekNumber(lastDay);
      return Array.from({ length: lastWeek }, (_, i) => i + 1);
    }
  }

  /**
   * 获取指定年份的日期范围（按月分组）
   * @param {number} year - 年份
   * @returns {Array} 日期信息数组
   */
  getDaysOfYear(year) {
    const currentYear = this.getCurrentYear();
    const currentMonth = this.getCurrentMonth();
    const currentDay = this.getCurrentDay();

    const months =
      year === currentYear
        ? Array.from({ length: currentMonth }, (_, i) => i + 1)
        : Array.from({ length: 12 }, (_, i) => i + 1);

    const result = [];

    months.forEach((month) => {
      const daysInMonth = new Date(year, month, 0).getDate();
      const maxDay =
        year === currentYear && month === currentMonth
          ? currentDay
          : daysInMonth;

      result.push({
        month: month,
        days: Array.from({ length: maxDay }, (_, i) => i + 1),
      });
    });

    return result;
  }

  /**
   * 获取近N天的日期数组
   * @param {number} days - 天数
   * @returns {Array} 日期数组
   */
  getRecentDays(days) {
    const result = [];
    const today = new Date(this._currentDate);
    today.setHours(0, 0, 0, 0);

    for (let i = days - 1; i >= 0; i--) {
      const date = new Date(today);
      date.setDate(today.getDate() - i);
      result.push(date);
    }

    return result;
  }

  /**
   * 获取ISO周数
   * @param {Date} date - 日期
   * @returns {number} 周数（1-53）
   * @private
   */
  _getISOWeekNumber(date) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
    const week1 = new Date(d.getFullYear(), 0, 4);
    return (
      1 +
      Math.round(((d - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7)
    );
  }

  /**
   * 根据ISO周数获取该周的第一天（周一）
   * @param {number} year - 年份
   * @param {number} week - 周数（1-53）
   * @returns {Date} 该周的第一天
   * @private
   */
  _getDateOfISOWeek(year, week) {
    const simple = new Date(year, 0, 1 + (week - 1) * 7);
    const dow = simple.getDay();
    const ISOweekStart = simple;
    if (dow <= 4) {
      ISOweekStart.setDate(simple.getDate() - simple.getDay() + 1);
    } else {
      ISOweekStart.setDate(simple.getDate() + 8 - simple.getDay());
    }
    return ISOweekStart;
  }

  /**
   * 格式化日期为显示格式
   * @param {Date} date - 日期
   * @returns {string} 格式化后的日期
   */
  formatDisplayDate(date) {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    return `${month}月${day}日`;
  }
}
