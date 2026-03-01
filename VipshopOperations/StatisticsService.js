/**
 * 统计服务 - 负责销售数据的统计和计算
 *
 * @class StatisticsService
 * @description 作为销售数据的统计计算核心，提供以下功能：
 * - 年/月/周/日维度的销量查询和汇总
 * - 近N天销售数据的统计（销量、曝光UV、商详UV、加购UV、拒退件数）
 * - 日期范围计算（年份、月份、周数、天数）
 * - ISO周数计算
 * - 销售数据缓存管理
 *
 * 该类采用单例模式，确保全局只有一个统计服务实例。
 * 所有统计方法都基于 ProductSales 实体数据计算。
 *
 * @example
 * // 获取统计服务实例
 * const statisticsService = StatisticsService.getInstance(repository);
 *
 * // 获取指定货号的年销量
 * const yearSales = statisticsService.getYearTotalSales("A001", 2024);
 *
 * // 获取近7天销量
 * const last7DaysSales = statisticsService.getLastNDaysSum("A001", 7, "salesQuantity");
 *
 * // 获取年份范围
 * const years = statisticsService.getYearRange(); // { beforeLast: 2022, last: 2023, current: 2024 }
 */
class StatisticsService {
  /** @type {StatisticsService} 单例实例 */
  static _instance = null;

  /**
   * 创建统计服务实例
   * @param {Repository} [repository] - 数据仓库实例，若不提供则自动获取
   */
  constructor(repository) {
    if (StatisticsService._instance) {
      return StatisticsService._instance;
    }

    this._repository = repository || Repository.getInstance();
    this._converter = Converter.getInstance();

    this._salesCache = null;
    this._currentDate = new Date();

    StatisticsService._instance = this;
  }

  /**
   * 获取统计服务的单例实例
   * @static
   * @param {Repository} [repository] - 数据仓库实例
   * @returns {StatisticsService} 统计服务实例
   */
  static getInstance(repository) {
    if (!StatisticsService._instance) {
      StatisticsService._instance = new StatisticsService(repository);
    }
    return StatisticsService._instance;
  }

  /**
   * 刷新销售数据缓存
   * @private
   * @returns {void}
   * @description
   * 从仓库重新加载所有 ProductSales 数据到缓存。
   * 目前该方法定义了但不被使用，预留用于性能优化。
   */
  _refreshSalesCache() {
    this._salesCache = this._repository.findAll("ProductSales");
  }

  /**
   * 计算数据集中指定字段的总和
   * @private
   * @param {Object[]} items - 数据项数组
   * @param {string} field - 字段名
   * @returns {number} 字段总和
   * @description
   * 遍历数组，累加每个元素的指定字段值。
   * 如果字段值不存在，按0处理。
   */
  _sumField(items, field) {
    return items.reduce((sum, item) => {
      return sum + item[field] || 0;
    }, 0);
  }

  /**
   * 获取指定货号在指定年份的全年销量
   * @param {string} itemNumber - 货号
   * @param {number} year - 年份
   * @returns {number} 全年销量总和
   * @description
   * 调用仓库的 findSalesByItemAndYear 方法查询该年份的所有销售记录，
   * 然后累加 salesQuantity 字段。
   */
  getYearTotalSales(itemNumber, year) {
    if (!itemNumber) return 0;

    // 查询销售数据
    const sales = this._repository.findSalesByItemAndYear(itemNumber, year);

    return this._sumField(sales, "salesQuantity");
  }

  /**
   * 获取指定货号在指定年份和月份的销量
   * @param {string} itemNumber - 货号
   * @param {number} year - 年份
   * @param {number} month - 月份（1-12）
   * @returns {number} 月销量总和
   * @description
   * 调用仓库的 findSalesByItemAndYearMonth 方法查询该月的所有销售记录，
   * 然后累加 salesQuantity 字段。
   */
  getMonthSales(itemNumber, year, month) {
    if (!itemNumber) return 0;

    // 查询销售数据
    const sales = this._repository.findSalesByItemAndYearMonth(
      itemNumber,
      year,
      month,
    );

    return this._sumField(sales, "salesQuantity");
  }

  /**
   * 获取指定货号在指定年份和周数的销量
   * @param {string} itemNumber - 货号
   * @param {number} year - 年份
   * @param {number} week - 周数（1-53）
   * @returns {number} 周销量总和
   * @description
   * 调用仓库的 findSalesByItemAndYearWeek 方法查询该周的所有销售记录，
   * 然后累加 salesQuantity 字段。
   */
  getWeekSales(itemNumber, year, week) {
    if (!itemNumber) return 0;

    // 查询销售数据
    const sales = this._repository.findSalesByItemAndYearWeek(
      itemNumber,
      year,
      week,
    );

    return this._sumField(sales, "salesQuantity");
  }

  /**
   * 获取指定货号在指定日期的销量
   * @param {string} itemNumber - 货号
   * @param {Date|string} date - 日期
   * @returns {number} 日销量
   * @description
   * 调用仓库的 findSalesByItemAndDate 方法查询该日期的销售记录，
   * 返回 salesQuantity 字段，如果没有记录则返回0。
   */
  getDaySales(itemNumber, date) {
    if (!itemNumber) return 0;

    // 查询销售数据
    const sale = this._repository.findSalesByItemAndDate(itemNumber, date);

    return sale ? sale.salesQuantity : 0;
  }

  /**
   * 获取近N天内销售数据的指定字段总和
   * @param {string} itemNumber - 货号
   * @param {number} days - 天数
   * @param {string} field - 字段名（如 salesQuantity, exposureUV, productDetailsUV 等）
   * @returns {number} 指定字段的总和
   * @description
   * 调用仓库的 findSalesLastNDays 方法查询近N天的所有销售记录，
   * 然后累加指定字段的值。
   *
   * @example
   * // 获取近7天曝光UV总和
   * const uv = statisticsService.getLastNDaysSum("A001", 7, "exposureUV");
   */
  getLastNDaysSum(itemNumber, days, field) {
    if (!itemNumber) return 0;

    // 查询销售数据
    const sales = this._repository.findSalesLastNDays(itemNumber, days);

    return this._sumField(sales, field);
  }

  /**
   * 获取近N天的每日销量明细
   * @param {string} itemNumber - 货号
   * @param {number} days - 天数
   * @returns {Array<Object>} 每日销量明细数组
   * @returns {Date} return[].date - 日期对象
   * @returns {string} return[].dateStr - 格式化后的日期字符串（YYYY-MM-DD）
   * @returns {number} return[].sale - 当日销量
   * @description
   * 从最早到最晚返回近N天的每日销量。
   * 例如 days=3 时，返回 [前天, 昨天, 今天] 的销量。
   */
  getLastNDaysDailySales(itemNumber, days) {
    const result = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = days - 1; i >= 0; i--) {
      const date = new Date(today);
      date.setDate(today.getDate() - i);
      const sale = this.getDaySales(itemNumber, date);

      result.push({
        date,
        dateStr: this._converter.toDateStr(date),
        sale,
      });
    }

    return result;
  }

  /**
   * 获取当前日期
   * @returns {Date} 当前日期对象
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
   * 获取当前日期（当月第几天）
   * @returns {number} 当前日期（1-31）
   */
  getCurrentDay() {
    return this._currentDate.getDate();
  }

  /**
   * 获取当前周数（ISO周数）
   * @returns {number} 当前周数（1-53）
   */
  getCurrentWeek() {
    return this._getISOWeekNumber(this._currentDate);
  }

  /**
   * 获取前年、去年、今年的年份
   * @returns {Object} 年份范围
   * @returns {number} return.beforeLast - 前年（当前年份-2）
   * @returns {number} return.last - 去年（当前年份-1）
   * @returns {number} return.current - 今年
   *
   * @example
   * // 假设当前是2024年
   * getYearRange() // 返回 { beforeLast: 2022, last: 2023, current: 2024 }
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
   * @returns {number[]} 月份数组（1-12 或 1-当前月份）
   * @description
   * 如果是当前年份，返回 1 到当前月份；
   * 如果是其他年份，返回 1 到 12。
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
   * @returns {number[]} 周数数组
   * @description
   * 计算逻辑：
   * - 如果是当前年份，返回 1 到当前周数
   * - 如果是其他年份，计算该年最后一天的周数，返回 1 到该周数
   *
   * 注意：一年最多有53周（ISO周数系统）
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
   * @returns {Array<Object>} 按月分组的日期数组
   * @returns {number} return[].month - 月份
   * @returns {number[]} return[].days - 该月的日期数组
   * @description
   * 计算逻辑：
   * - 如果是当前年份，返回 1月到当前月，每月到当前日
   * - 如果是其他年份，返回 1月到12月，每月到该月最后一天
   *
   * @example
   * // 假设当前是2024年3月15日
   * getDaysOfYear(2024) // 返回 [
   * //   { month: 1, days: [1,2,...,31] },
   * //   { month: 2, days: [1,2,...,29] },
   * //   { month: 3, days: [1,2,...,15] }
   * // ]
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
   * @returns {Date[]} 日期对象数组（从最早到最晚）
   * @description
   * 返回从 days-1 天前到今天的所有日期。
   * 例如 days=3 时，返回 [前天, 昨天, 今天] 的日期对象。
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
   * @private
   * @param {Date} date - 日期对象
   * @returns {number} ISO周数（1-53）
   * @description
   * ISO周数定义：
   * - 一周从周一开始
   * - 一年的第一周是包含该年第一个周四的那一周
   * - 周数范围 1-53
   *
   * 算法说明：
   * 1. 将日期调整到目标周（加3天，减去(星期几-1)的调整）
   * 2. 找到该年第一个周四（1月4日）
   * 3. 计算相差天数，除以7得到周数
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
   * 格式化日期为显示格式
   * @param {Date} date - 日期对象
   * @returns {string} 格式化后的日期字符串（如 "3月15日"）
   */
  formatDisplayDate(date) {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    return `${month}月${day}日`;
  }
}
