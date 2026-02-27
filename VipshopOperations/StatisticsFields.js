class StatisticsFields {
  static _instance = null;

  constructor(statisticsService) {
    if (StatisticsFields._instance) {
      return StatisticsFields._instance;
    }

    this._service = statisticsService || StatisticsService.getInstance();
    this._years = this._service.getYearRange();

    StatisticsFields._instance = this;
  }

  // 单例模式
  static getInstance(statisticsService) {
    if (!StatisticsFields._instance) {
      StatisticsFields._instance = new StatisticsFields(statisticsService);
    }
    return StatisticsFields._instance;
  }

  // 获取所有可用的统计字段
  getAllFields() {
    return [
      // ========== 年销量（3个汇总字段）==========
      {
        field: "yearSales_beforeLast",
        title: `前年(${this._years.beforeLast})年销量`,
        type: "summary",
        width: 12,
        format: "#,##0",
        group: "年销量汇总",
        description: `${this._years.beforeLast}年全年销量总计`,
        compute: (product) => {
          return this._service.getYearTotalSales(
            product.itemNumber,
            this._years.beforeLast,
          );
        },
      },
      {
        field: "yearSales_last",
        title: `去年(${this._years.last})年销量`,
        type: "summary",
        width: 12,
        format: "#,##0",
        group: "年销量汇总",
        description: `${this._years.last}年全年销量总计`,
        compute: (product) => {
          return this._service.getYearTotalSales(
            product.itemNumber,
            this._years.last,
          );
        },
      },
      {
        field: "yearSales_current",
        title: `今年(${this._years.current})年销量`,
        type: "summary",
        width: 12,
        format: "#,##0",
        group: "年销量汇总",
        description: `${this._years.current}年截止当前销量总计`,
        compute: (product) => {
          return this._service.getYearTotalSales(
            product.itemNumber,
            this._years.current,
          );
        },
      },

      // ========== 月销量（3个展开字段）==========
      {
        field: "monthSales_beforeLast",
        title: "前年月销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "月销量明细",
        description: "前年每月销量",
        expandConfig: {
          type: "month",
          year: this._years.beforeLast,
          label: "前年",
        },
      },
      {
        field: "monthSales_last",
        title: "去年月销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "月销量明细",
        description: "去年每月销量",
        expandConfig: {
          type: "month",
          year: this._years.last,
          label: "去年",
        },
      },
      {
        field: "monthSales_current",
        title: "今年月销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "月销量明细",
        description: "今年每月销量（截止本月）",
        expandConfig: {
          type: "month",
          year: this._years.current,
          label: "今年",
          isCurrent: true,
        },
      },

      // ========== 周销量（3个展开字段）==========
      {
        field: "weekSales_beforeLast",
        title: "前年周销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "周销量明细",
        description: "前年每周销量",
        expandConfig: {
          type: "week",
          year: this._years.beforeLast,
          label: "前年",
        },
      },
      {
        field: "weekSales_last",
        title: "去年周销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "周销量明细",
        description: "去年每周销量",
        expandConfig: {
          type: "week",
          year: this._years.last,
          label: "去年",
        },
      },
      {
        field: "weekSales_current",
        title: "今年周销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "周销量明细",
        description: "今年每周销量（截止本周）",
        expandConfig: {
          type: "week",
          year: this._years.current,
          label: "今年",
          isCurrent: true,
        },
      },

      // ========== 日销量（3个展开字段）==========
      {
        field: "daySales_beforeLast",
        title: "前年日销量",
        type: "expandable",
        width: 12,
        format: "#,##0",
        group: "日销量明细",
        description: "前年每日销量",
        expandConfig: {
          type: "day",
          year: this._years.beforeLast,
          label: "前年",
        },
      },
      {
        field: "daySales_last",
        title: "去年日销量",
        type: "expandable",
        width: 12,
        format: "#,##0",
        group: "日销量明细",
        description: "去年每日销量",
        expandConfig: {
          type: "day",
          year: this._years.last,
          label: "去年",
        },
      },
      {
        field: "daySales_current",
        title: "今年日销量",
        type: "expandable",
        width: 12,
        format: "#,##0",
        group: "日销量明细",
        description: "今年每日销量（截止今日）",
        expandConfig: {
          type: "day",
          year: this._years.current,
          label: "今年",
          isCurrent: true,
        },
      },

      // ========== 近7天UV指标（5个汇总字段，不展开）==========
      {
        field: "sales_last7Days",
        title: "近7天销量",
        type: "summary",
        width: 12,
        format: "#,##0",
        group: "近7天汇总",
        description: "最近7天销量总和",
        compute: (product) => {
          return this._service.getLastNDaysSum(
            product.itemNumber,
            7,
            "salesQuantity",
          );
        },
      },
      {
        field: "uv_exposure_last7Days",
        title: "近7天曝光UV",
        type: "summary",
        width: 14,
        format: "#,##0",
        group: "近7天汇总",
        description: "最近7天曝光UV总和",
        compute: (product) => {
          return this._service.getLastNDaysSum(
            product.itemNumber,
            7,
            "exposureUV",
          );
        },
      },
      {
        field: "uv_productDetails_last7Days",
        title: "近7天商详UV",
        type: "summary",
        width: 14,
        format: "#,##0",
        group: "近7天汇总",
        description: "最近7天商详UV总和",
        compute: (product) => {
          return this._service.getLastNDaysSum(
            product.itemNumber,
            7,
            "productDetailsUV",
          );
        },
      },
      {
        field: "uv_addToCart_last7Days",
        title: "近7天加购UV",
        type: "summary",
        width: 14,
        format: "#,##0",
        group: "近7天汇总",
        description: "最近7天加购UV总和",
        compute: (product) => {
          return this._service.getLastNDaysSum(
            product.itemNumber,
            7,
            "addToCartUV",
          );
        },
      },
      {
        field: "rejectCount_last7Days",
        title: "近7天拒退件数",
        type: "summary",
        width: 14,
        format: "#,##0",
        group: "近7天汇总",
        description: "最近7天拒退件数总和",
        compute: (product) => {
          return this._service.getLastNDaysSum(
            product.itemNumber,
            7,
            "rejectAndReturnCount",
          );
        },
      },

      // ========== 近15天销量（展开字段）==========
      {
        field: "sales_last15Days",
        title: "近15天销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "近15天销量明细",
        description: "最近15天每日销量",
        expandConfig: {
          type: "recent",
          days: 15,
          label: "近15天",
        },
      },

      // ========== 近30天销量（展开字段）==========
      {
        field: "sales_last30Days",
        title: "近30天销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "近30天销量明细",
        description: "最近30天每日销量",
        expandConfig: {
          type: "recent",
          days: 30,
          label: "近30天",
        },
      },

      // ========== 近45天销量（展开字段）==========
      {
        field: "sales_last45Days",
        title: "近45天销量",
        type: "expandable",
        width: 10,
        format: "#,##0",
        group: "近45天销量明细",
        description: "最近45天每日销量",
        expandConfig: {
          type: "recent",
          days: 45,
          label: "近45天",
        },
      },
    ];
  }

  // 展开字段配置
  expandField(field) {
    if (field.type !== "expandable" || !field.expandConfig) {
      return [field];
    }

    const config = field.expandConfig;
    const expandedFields = [];

    switch (config.type) {
      case "month":
        expandedFields.push(...this._expandMonthField(config, field));
        break;
      case "week":
        expandedFields.push(...this._expandWeekField(config, field));
        break;
      case "day":
        expandedFields.push(...this._expandDayField(config, field));
        break;
      case "recent":
        expandedFields.push(...this._expandRecentField(config, field));
        break;
      default:
        return [field];
    }

    return expandedFields;
  }

  // 展开月份字段
  _expandMonthField(config, baseField) {
    const fields = [];
    const months = this._service.getMonthsOfYear(config.year);

    months.forEach((month) => {
      const monthStr = String(month).padStart(2, "0");
      fields.push({
        field: `month_${config.year}_${monthStr}`,
        title: `${config.label}${month}月`,
        type: "computed",
        width: baseField.width || 10,
        format: baseField.format,
        group: baseField.group,
        description: `${config.year}年${month}月销量`,
        parentField: baseField.field,
        compute: (product) => {
          return this._service.getMonthSales(
            product.itemNumber,
            config.year,
            month,
          );
        },
      });
    });

    return fields;
  }

  // 展开周数字段
  _expandWeekField(config, baseField) {
    const fields = [];
    const weeks = this._service.getWeeksOfYear(config.year);

    weeks.forEach((week) => {
      const weekStr = String(week).padStart(2, "0");
      fields.push({
        field: `week_${config.year}_${weekStr}`,
        title: `${config.label}第${week}周`,
        type: "computed",
        width: baseField.width || 10,
        format: baseField.format,
        group: baseField.group,
        description: `${config.year}年第${week}周销量`,
        parentField: baseField.field,
        compute: (product) => {
          return this._service.getWeekSales(
            product.itemNumber,
            config.year,
            week,
          );
        },
      });
    });

    return fields;
  }

  // 展开日数字段
  _expandDayField(config, baseField) {
    const fields = [];
    const months = this._service.getDaysOfYear(config.year);

    months.forEach(({ month, days }) => {
      days.forEach((day) => {
        const monthStr = String(month).padStart(2, "0");
        const dayStr = String(day).padStart(2, "0");

        fields.push({
          field: `day_${config.year}_${monthStr}_${dayStr}`,
          title: `${config.label}${month}月${day}日`,
          type: "computed",
          width: baseField.width || 12,
          format: baseField.format,
          group: baseField.group,
          description: `${config.year}年${month}月${day}日销量`,
          parentField: baseField.field,
          compute: (product) => {
            const date = new Date(config.year, month - 1, day);
            return this._service.getDaySales(product.itemNumber, date);
          },
        });
      });
    });

    return fields;
  }

  // 展开近N天字段
  _expandRecentField(config, baseField) {
    const fields = [];
    const dates = this._service.getRecentDays(config.days);

    dates.forEach((date) => {
      const month = date.getMonth() + 1;
      const day = date.getDate();
      const monthStr = String(month).padStart(2, "0");
      const dayStr = String(day).padStart(2, "0");

      fields.push({
        field: `recent_${config.days}_${monthStr}_${dayStr}`,
        title: `${month}月${day}日`,
        type: "computed",
        width: baseField.width || 10,
        format: baseField.format,
        group: baseField.group,
        description: `${config.label}的${month}月${day}日销量`,
        parentField: baseField.field,
        compute: (product) => {
          return this._service.getDaySales(product.itemNumber, date);
        },
      });
    });

    return fields;
  }
}
