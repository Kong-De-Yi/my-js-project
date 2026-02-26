class SalesStatisticsService {
  constructor(repository) {
    this._repository = repository;
    this._salesCache = null;
    this._currentDate = new Date();
  }

  // 刷新销售数据缓存
  _refreshSalesCache() {
    this._salesCache = this._repository.findAll("ProductSales");
  }

  // 计算销售数据字段filed的总和
  _sumField(sales, field) {
    return sales.reduce((sum, sale) => {
      return sum + sale[field];
    }, 0);
  }

  // 获取指定货号在指定年份的全年销量
  getYearTotalSales(itemNumber, year) {
    if (!itemNumber) return 0;

    // 查询销售数据
    const sales = this._repository.findSalesByItemAndYear(itemNumber, year);

    return this._sumField(sales, "salesQuantity");
  }

  // 获取指定货号在指定年份和月份的销量
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

  // 获取指定货号在指定年份和周数的销量
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

  // 获取指定货号在指定日期的销量
  getDaySales(itemNumber, date) {
    if (!itemNumber) return 0;

    // 查询销售数据
    const sale = this._repository.findSalesByItemAndDate(itemNumber, date);

    return sale ? sale.salesQuantity : 0;
  }

  // 获取近N天内销售数据的指定字段总和
  getLastNDaysSum(itemNumber, days, field) {
    if (!itemNumber) return 0;

    // 查询销售数据
    const sales = this._repository.findSalesLastNDays(itemNumber, days);

    return this._sumField(sales, field);
  }

  // 获取近N天的每日销量
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
        dateStr: _converter.toDateStr(date),
        sale,
      });
    }

    return result;
  }

  // 获取当前日期
  getCurrentDate() {
    return this._currentDate;
  }

  // 获取当前年份
  getCurrentYear() {
    return this._currentDate.getFullYear();
  }

  // 获取当前月份
  getCurrentMonth() {
    return this._currentDate.getMonth() + 1;
  }

  // 获取当前日期
  getCurrentDay() {
    return this._currentDate.getDate();
  }

  // 获取当前周数
  getCurrentWeek() {
    return this._getISOWeekNumber(this._currentDate);
  }

  // 获取前年、去年、今年的年份
  getYearRange() {
    const currentYear = this.getCurrentYear();
    return {
      beforeLast: currentYear - 2,
      last: currentYear - 1,
      current: currentYear,
    };
  }

  // 获取指定年份的月份范围
  getMonthsOfYear(year) {
    const currentYear = this.getCurrentYear();
    const currentMonth = this.getCurrentMonth();

    if (year === currentYear) {
      return Array.from({ length: currentMonth }, (_, i) => i + 1);
    } else {
      return Array.from({ length: 12 }, (_, i) => i + 1);
    }
  }

  // 获取指定年份的周数范围
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

  // 获取指定年份的日期范围（按月分组）
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

  // 获取近N天的日期数组
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

  // 获取ISO周数
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

  // 格式化日期为显示格式
  formatDisplayDate(date) {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    return `${month}月${day}日`;
  }
}
