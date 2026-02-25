// ============================================================================
// IndexConfig.js - 索引配置中心（增强版）
// 功能：添加销售统计相关的索引配置
// ============================================================================

class IndexConfig {
  static _instance = null;

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
      { fields: ["subProductCode"], unique: false },
      { fields: ["productCode", "subProductCode"], unique: true },
    ];

    // ========== 商品销售实体索引配置）==========
    this.PRODUCT_SALES_INDEXES = [
      // 联合主键索引：货号 + 日期
      { fields: ["itemNumber", "salesDate"], unique: true },

      // 按年份索引 - 快速筛选某年的数据
      { fields: ["salesYear"], unique: false },

      // 按年月索引 - 快速筛选某年某月的数据
      { fields: ["yearMonth"], unique: false },

      // 按年周索引 - 快速筛选某年某周的数据
      { fields: ["yearWeek"], unique: false },

      // 按距今天数索引 - 快速筛选近N天的数据
      { fields: ["daysSinceSale"], unique: false },

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

  static getInstance() {
    if (!IndexConfig._instance) {
      IndexConfig._instance = new IndexConfig();
    }
    return IndexConfig._instance;
  }

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

  _calculateMatchScore(indexFields, queryFields) {
    const matchCount = indexFields.filter((f) =>
      queryFields.includes(f),
    ).length;
    return matchCount / Math.max(indexFields.length, queryFields.length);
  }
}
