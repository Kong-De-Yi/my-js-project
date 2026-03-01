/**
 * 统计字段定义 - 提供所有可用的统计字段配置和展开逻辑
 *
 * @class StatisticsFields
 * @description 作为统计字段的元数据管理类，提供以下功能：
 * - 定义所有可用的统计字段（年/月/周/日销量、UV指标、拒退率等）
 * - 字段类型分为三类：
 *   - summary: 汇总字段，直接计算返回单个值
 *   - expandable: 可展开字段，运行时展开为多个具体字段
 *   - computed: 计算字段，由expandable展开后生成
 * - 字段展开逻辑：将抽象字段（如"近7天销量"）展开为多个具体日期的字段
 * - 动态年份计算：根据当前时间自动计算前年、去年、今年
 *
 * 该类采用单例模式，确保全局只有一个统计字段定义实例。
 *
 * @example
 * // 获取统计字段实例
 * const statisticsFields = StatisticsFields.getInstance(statisticsService);
 *
 * // 获取所有可用字段
 * const allFields = statisticsFields.getAllFields();
 *
 * // 获取特定字段
 * const field = statisticsFields.getField("sales_last7Days");
 *
 * // 展开可展开字段
 * const expandedFields = statisticsFields.expandField(field);
 */
class StatisticsFields {
  /** @type {StatisticsFields} 单例实例 */
  static _instance = null;

  /**
   * 创建统计字段定义实例
   * @param {StatisticsService} [statisticsService] - 统计服务实例，若不提供则自动获取
   */
  constructor(statisticsService) {
    if (StatisticsFields._instance) {
      return StatisticsFields._instance;
    }

    this._service = statisticsService || StatisticsService.getInstance();
    this._years = this._service.getYearRange();

    StatisticsFields._instance = this;
  }

  /**
   * 获取统计字段定义的单例实例
   * @static
   * @param {StatisticsService} [statisticsService] - 统计服务实例
   * @returns {StatisticsFields} 统计字段定义实例
   */
  static getInstance(statisticsService) {
    if (!StatisticsFields._instance) {
      StatisticsFields._instance = new StatisticsFields(statisticsService);
    }
    return StatisticsFields._instance;
  }

  /**
   * 获取所有可用的统计字段配置
   * @returns {Array<Object>} 统计字段配置数组
   *
   * @description
   * 返回的字段按以下分类组织：
   *
   * 1. 年销量汇总（3个summary字段）
   *    - yearSales_beforeLast: 前年销量
   *    - yearSales_last: 去年销量
   *    - yearSales_current: 今年销量
   *
   * 2. 月销量明细（3个expandable字段）
   *    - monthSales_beforeLast: 前年各月销量
   *    - monthSales_last: 去年各月销量
   *    - monthSales_current: 今年各月销量
   *
   * 3. 周销量明细（3个expandable字段）
   *    - weekSales_beforeLast: 前年各周销量
   *    - weekSales_last: 去年各周销量
   *    - weekSales_current: 今年各周销量
   *
   * 4. 日销量明细（3个expandable字段）
   *    - daySales_beforeLast: 前年各日销量
   *    - daySales_last: 去年各日销量
   *    - daySales_current: 今年各日销量
   *
   * 5. 近7天UV指标（5个summary字段）
   *    - sales_last7Days: 近7天销量
   *    - uv_exposure_last7Days: 近7天曝光UV
   *    - uv_productDetails_last7Days: 近7天商详UV
   *    - uv_addToCart_last7Days: 近7天加购UV
   *    - rejectCount_last7Days: 近7天拒退件数
   *
   * 6. 近N天销量明细（3个expandable字段）
   *    - sales_last15Days: 近15天每日销量
   *    - sales_last30Days: 近30天每日销量
   *    - sales_last45Days: 近45天每日销量
   *
   * 每个字段配置包含：
   * - field: 字段标识
   * - title: 显示标题
   * - type: 字段类型（summary/expandable/computed）
   * - width: 默认列宽
   * - format: 数字格式
   * - group: 所属分组
   * - description: 字段描述
   * - compute: 计算函数（summary类型）
   * - expandConfig: 展开配置（expandable类型）
   */
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

  /**
   * 根据字段标识获取字段配置
   * @param {string} field - 字段标识（如 "sales_last7Days"）
   * @returns {Object|undefined} 字段配置对象，未找到时返回undefined
   */
  getField(field) {
    return this.getAllFields().find((f) => (f.field = field));
  }

  /**
   * 展开可展开字段
   * @param {Object} field - 字段配置对象
   * @returns {Array<Object>} 展开后的字段数组
   * @description
   * 展开逻辑：
   * 1. 如果字段类型不是 expandable 或没有 expandConfig，直接返回原字段
   * 2. 根据 expandConfig.type 调用对应的展开方法：
   *    - month: 按月份展开（12个月）
   *    - week: 按周数展开（最多53周）
   *    - day: 按日期展开（全年每天）
   *    - recent: 按最近N天展开
   * 3. 返回展开后的具体字段数组（每个字段类型为 computed）
   *
   * 展开后的字段包含：
   * - field: 具体字段标识（如 "month_2024_01"）
   * - title: 具体标题（如 "1月"）
   * - type: "computed"
   * - width/format/group: 继承自父字段
   * - parentField: 父字段标识
   * - compute: 具体计算函数
   */
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

  /**
   * 展开月份字段
   * @private
   * @param {Object} config - 展开配置
   * @param {number} config.year - 年份
   * @param {string} config.label - 年份标签（如"前年"、"去年"）
   * @param {Object} baseField - 基础字段配置
   * @returns {Array<Object>} 展开后的月份字段数组
   * @description
   * 为指定年份的每个月创建一个字段：
   * - 字段名格式：month_{年份}_{月份（两位）}
   * - 标题格式：{标签}{月份}月
   * - 计算函数：调用 service.getMonthSales()
   */
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

  /**
   * 展开周数字段
   * @private
   * @param {Object} config - 展开配置
   * @param {number} config.year - 年份
   * @param {string} config.label - 年份标签
   * @param {Object} baseField - 基础字段配置
   * @returns {Array<Object>} 展开后的周数字段数组
   * @description
   * 为指定年份的每周创建一个字段：
   * - 字段名格式：week_{年份}_{周数（两位）}
   * - 标题格式：{标签}第{周数}周
   * - 计算函数：调用 service.getWeekSales()
   */
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

  /**
   * 展开日数字段
   * @private
   * @param {Object} config - 展开配置
   * @param {number} config.year - 年份
   * @param {string} config.label - 年份标签
   * @param {Object} baseField - 基础字段配置
   * @returns {Array<Object>} 展开后的日数字段数组
   * @description
   * 为指定年份的每一天创建一个字段：
   * - 字段名格式：day_{年份}_{月份（两位）}_{日期（两位）}
   * - 标题格式：{标签}{月份}月{日期}日
   * - 计算函数：调用 service.getDaySales()
   *
   * 注意：全年最多366天，展开后会生成大量字段
   */
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

  /**
   * 展开近N天字段
   * @private
   * @param {Object} config - 展开配置
   * @param {number} config.days - 天数（如15、30、45）
   * @param {string} config.label - 标签（如"近15天"）
   * @param {Object} baseField - 基础字段配置
   * @returns {Array<Object>} 展开后的日数字段数组
   * @description
   * 为最近N天的每一天创建一个字段：
   * - 字段名格式：recent_{天数}_{月份（两位）}_{日期（两位）}
   * - 标题格式：{月份}月{日期}日
   * - 计算函数：调用 service.getDaySales()
   *
   * 日期顺序：从最早到最晚（与报表显示顺序一致）
   */
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
