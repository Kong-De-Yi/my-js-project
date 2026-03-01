/**
 * 索引配置 - 定义所有业务实体的索引配置
 *
 * @class IndexConfig
 * @description 作为系统的索引配置中心，提供以下功能：
 * - 定义每个业务实体的索引结构（字段组合、唯一性）
 * - 为 Repository 提供索引创建的依据
 * - 支持索引匹配度计算，用于查询优化建议
 * - 集中管理所有实体的索引配置，便于维护和调优
 *
 * 索引类型说明：
 * - 单字段索引：如 [{ fields: ["itemNumber"], unique: true }]
 * - 复合索引：如 [{ fields: ["productCode", "subProductCode"], unique: true }]
 * - 唯一索引：unique: true 表示该索引字段组合的值必须唯一
 * - 非唯一索引：unique: false 允许重复值
 *
 * 该类采用单例模式，确保全局只有一个索引配置实例。
 *
 * @example
 * // 获取索引配置实例
 * const indexConfig = IndexConfig.getInstance(); *
 * // 获取 Product 实体的索引配置
 * const productIndexes = indexConfig.getIndexes("Product");
 *
 * // 获取推荐索引（基于查询字段）
 * const recommendations = indexConfig.getRecommendedIndex("Product", ["itemNumber", "styleNumber"]);
 */
class IndexConfig {
  /** @type {IndexConfig} 单例实例 */
  static _instance = null;

  /**
   * 创建索引配置实例
   * @private
   */
  constructor() {
    if (IndexConfig._instance) {
      return IndexConfig._instance;
    }

    // ========== 商品主实体索引配置 ==========
    this.PRODUCT_INDEXES = [
      { fields: ["itemNumber"], unique: true },
      { fields: ["styleNumber"], unique: false },
      { fields: ["thirdLevelCategory"], unique: false },
      { fields: ["fourthLevelCategory"], unique: false },
    ];

    // ========== 商品价格实体索引配置 ==========
    this.PRODUCT_PRICE_INDEXES = [{ fields: ["itemNumber"], unique: true }];

    // ========== 常态商品实体索引配置 ==========
    this.REGULAR_PRODUCT_INDEXES = [{ fields: ["itemNumber"], unique: false }];

    // ========== 库存实体索引配置 ==========
    this.INVENTORY_INDEXES = [{ fields: ["productCode"], unique: true }];

    // ========== 组合商品实体索引配置 ==========
    this.COMBO_PRODUCT_INDEXES = [
      { fields: ["productCode"], unique: false },
      { fields: ["productCode", "subProductCode"], unique: true },
    ];

    // ========== 商品销售实体索引配置）==========
    this.PRODUCT_SALES_INDEXES = [
      // 联合主键索引：货号 + 日期
      { fields: ["itemNumber", "salesDate"], unique: true },

      // 按货号+年份索引 - 快速查询某货号某年的数据
      { fields: ["itemNumber", "salesYear"], unique: false },

      // 按货号+年月索引 - 快速查询某货号某月的数据
      { fields: ["itemNumber", "yearMonth"], unique: false },

      // 按货号+年周索引 - 快速查询某货号某周的数据
      { fields: ["itemNumber", "yearWeek"], unique: false },
    ];

    // ========== 品牌配置实体索引配置 ==========
    this.BRAND_CONFIG_INDEXES = [{ fields: ["brandSN"], unique: true }];

    // ========== 报表模板实体索引配置 ==========
    this.REPORT_TEMPLATE_INDEXES = [
      { fields: ["templateName"], unique: false },
      { fields: ["templateName", "fieldName"], unique: true },
    ];

    IndexConfig._instance = this;
  }

  /**
   * 获取索引配置的单例实例
   * @static
   * @returns {IndexConfig} 索引配置实例
   */
  static getInstance() {
    if (!IndexConfig._instance) {
      IndexConfig._instance = new IndexConfig();
    }
    return IndexConfig._instance;
  }

  /**
   * 获取指定实体的索引配置
   * @param {string} entityName - 实体名称
   * @returns {Array<Object>} 索引配置数组
   *
   * @description
   * 支持的实体及其索引配置：
   *
   * 1. Product（商品主实体）
   *    - [{ fields: ["itemNumber"], unique: true }] - 货号唯一索引
   *    - [{ fields: ["styleNumber"], unique: false }] - 款号索引
   *    - [{ fields: ["thirdLevelCategory"], unique: false }] - 三级品类索引
   *    - [{ fields: ["fourthLevelCategory"], unique: false }] - 四级品类索引
   *
   * 2. ProductPrice（商品价格）
   *    - [{ fields: ["itemNumber"], unique: true }] - 货号唯一索引
   *
   * 3. RegularProduct（常态商品）
   *    - [{ fields: ["itemNumber"], unique: false }] - 货号索引
   *
   * 4. Inventory（库存）
   *    - [{ fields: ["productCode"], unique: true }] - 条码唯一索引
   *
   * 5. ComboProduct（组合商品）
   *    - [{ fields: ["productCode"], unique: false }] - 主商品条码索引
   *    - [{ fields: ["productCode", "subProductCode"], unique: true }] - 组合关系唯一索引
   *
   * 6. ProductSales（商品销售）
   *    - [{ fields: ["itemNumber", "salesDate"], unique: true }] - 联合主键索引
   *    - [{ fields: ["itemNumber", "salesYear"], unique: false }] - 货号+年份索引
   *    - [{ fields: ["itemNumber", "yearMonth"], unique: false }] - 货号+年月索引
   *    - [{ fields: ["itemNumber", "yearWeek"], unique: false }] - 货号+年周索引
   *
   * 7. BrandConfig（品牌配置）
   *    - [{ fields: ["brandSN"], unique: true }] - 品牌编号唯一索引
   *
   * 8. ReportTemplate（报表模板）
   *    - [{ fields: ["templateName"], unique: false }] - 模板名称索引
   *    - [{ fields: ["templateName", "fieldName"], unique: true }] - 模板字段唯一索引
   */

  getIndexes(entityName) {
    switch (entityName) {
      case "Product":
        return this.PRODUCT_INDEXES;
      case "ProductPrice":
        return this.PRODUCT_PRICE_INDEXES;
      case "RegularProduct":
        return this.REGULAR_PRODUCT_INDEXES;
      case "Inventory":
        return this.INVENTORY_INDEXES;
      case "ComboProduct":
        return this.COMBO_PRODUCT_INDEXES;
      case "ProductSales":
        return this.PRODUCT_SALES_INDEXES;
      case "BrandConfig":
        return this.BRAND_CONFIG_INDEXES;
      case "ReportTemplate":
        return this.REPORT_TEMPLATE_INDEXES;
      default:
        return [];
    }
  }

  /**
   * 获取所有实体的索引配置
   * @returns {Object.<string, Array<Object>>} 所有实体的索引配置映射
   *
   * @example
   * {
   *   Product: [...],
   *   ProductPrice: [...],
   *   RegularProduct: [...],
   *   ...
   * }
   */

  getAllIndexes() {
    return {
      Product: this.PRODUCT_INDEXES,
      ProductPrice: this.PRODUCT_PRICE_INDEXES,
      RegularProduct: this.REGULAR_PRODUCT_INDEXES,
      Inventory: this.INVENTORY_INDEXES,
      ComboProduct: this.COMBO_PRODUCT_INDEXES,
      ProductSales: this.PRODUCT_SALES_INDEXES,
      BrandConfig: this.BRAND_CONFIG_INDEXES,
      ReportTemplate: this.REPORT_TEMPLATE_INDEXES,
    };
  }

  /**
   * 获取推荐索引（基于查询字段）
   * @param {string} entityName - 实体名称
   * @param {string[]} queryPattern - 查询字段数组
   * @returns {Array<Object>} 按匹配度排序的推荐索引
   * @returns {string[]} return[].fields - 索引字段
   * @returns {boolean} return[].unique - 是否唯一索引
   * @returns {number} return[].matchScore - 匹配度分数（0-1）
   *
   * @description
   * 推荐算法：
   * 1. 获取实体的所有索引配置
   * 2. 对每个索引计算与查询字段的匹配度
   * 3. 过滤掉匹配度为0的索引
   * 4. 按匹配度从高到低排序
   *
   * 匹配度计算公式：匹配字段数 / Max(索引字段数, 查询字段数)
   *
   * @example
   * // 查询包含 itemNumber 和 salesDate 的推荐索引
   * const recommendations = indexConfig.getRecommendedIndex(
   *   "ProductSales",
   *   ["itemNumber", "salesDate"]
   * );
   * // 返回: [
   * //   { fields: ["itemNumber", "salesDate"], unique: true, matchScore: 1 },
   * //   { fields: ["itemNumber", "salesYear"], unique: false, matchScore: 0.5 }
   * // ]
   */
  getRecommendedIndex(entityName, queryPattern) {
    const indexes = this.getIndexes(entityName);

    return indexes
      .map((idx) => ({
        ...idx,
        matchScore: this._calculateMatchScore(idx.fields, queryPattern),
      }))
      .filter((idx) => idx.matchScore > 0)
      .sort((a, b) => b.matchScore - a.matchScore);
  }

  /**
   * 计算索引与查询字段的匹配度
   * @private
   * @param {string[]} indexFields - 索引字段数组
   * @param {string[]} queryFields - 查询字段数组
   * @returns {number} 匹配度分数（0-1）
   *
   * @description
   * 匹配度计算规则：
   * - 完全匹配：分数 = 1
   * - 部分匹配：分数 = 匹配字段数 / Max(索引字段数, 查询字段数)
   * - 无匹配：分数 = 0
   *
   * 例如：
   * - 索引 ["itemNumber", "salesDate"]，查询 ["itemNumber", "salesDate"] => 1
   * - 索引 ["itemNumber", "salesDate"]，查询 ["itemNumber"] => 0.5
   * - 索引 ["itemNumber", "salesDate"]，查询 ["styleNumber"] => 0
   */
  _calculateMatchScore(indexFields, queryFields) {
    const matchCount = indexFields.filter((f) =>
      queryFields.includes(f),
    ).length;
    return matchCount / Math.max(indexFields.length, queryFields.length);
  }
}
