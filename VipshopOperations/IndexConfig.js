// ============================================================================
// 索引配置中心
// 功能：集中配置所有实体的索引策略
// 特点：支持单字段和多字段联合索引，可根据业务需求灵活配置
// ============================================================================

class IndexConfig {
  static _instance = null;

  constructor() {
    if (IndexConfig._instance) {
      return IndexConfig._instance;
    }

    // ========== 商品主实体索引配置 ==========
    this.PRODUCT_INDEXES = [
      // 唯一索引：货号
      { fields: ["itemNumber"], unique: true },

      // 联合索引：品牌 + 商品状态（用于按品牌筛选上架商品）
      { fields: ["brandSN", "itemStatus"], unique: false },

      // 联合索引：款号 + 颜色（用于同款不同色查询）
      { fields: ["styleNumber", "color"], unique: false },

      // 联合索引：SPU（用于活动提报按SPU去重）
      { fields: ["P_SPU"], unique: false },

      // 联合索引：营销定位 + 商品状态（用于运营分析）
      { fields: ["marketingPositioning", "itemStatus"], unique: false },

      // 联合索引：首次上架时间（用于新品筛选）
      { fields: ["firstListingTime"], unique: false },

      // 联合索引：品牌 + 营销定位（用于品牌策略分析）
      { fields: ["brandSN", "marketingPositioning"], unique: false },

      // 联合索引：四级品类 + 三级品类（用于品类分析）
      { fields: ["fourthLevelCategory", "thirdLevelCategory"], unique: false },

      // 联合索引：可售天数范围（用于库存预警）
      { fields: ["sellableDays"], unique: false },

      // 联合索引：利润范围（用于利润分析）
      { fields: ["profit"], unique: false },
    ];

    // ========== 商品价格实体索引配置 ==========
    this.PRODUCT_PRICE_INDEXES = [
      // 唯一索引：货号
      { fields: ["itemNumber"], unique: true },
    ];

    // ========== 常态商品实体索引配置 ==========
    this.REGULAR_PRODUCT_INDEXES = [
      // 联合索引：货号（用于查询某货号的所有尺码）
      { fields: ["itemNumber"], unique: false },

      // 联合索引：条码（唯一）
      { fields: ["productCode"], unique: true },

      // 联合索引：货号 + 尺码状态（用于断码分析）
      { fields: ["itemNumber", "sizeStatus"], unique: false },
    ];

    // ========== 库存实体索引配置 ==========
    this.INVENTORY_INDEXES = [
      // 唯一索引：商品编码
      { fields: ["productCode"], unique: true },
    ];

    // ========== 组合商品实体索引配置 ==========
    this.COMBO_PRODUCT_INDEXES = [
      // 联合索引：组合商品编码（用于查询某组合的所有子商品）
      { fields: ["productCode"], unique: false },

      // 联合索引：子商品编码（用于反向查询）
      { fields: ["subProductCode"], unique: false },

      // 联合索引：组合商品编码 + 子商品编码（用于唯一性检查）
      { fields: ["productCode", "subProductCode"], unique: true },
    ];

    // ========== 商品销售实体索引配置 ==========
    this.PRODUCT_SALES_INDEXES = [
      // 联合索引：货号 + 日期（用于查询某货号某天的销售）
      { fields: ["itemNumber", "salesDate"], unique: true },

      // 联合索引：日期范围（用于批量查询）
      { fields: ["salesDate"], unique: false },

      // 联合索引：货号（用于汇总）
      { fields: ["itemNumber"], unique: false },
    ];

    // ========== 品牌配置实体索引配置 ==========
    this.BRAND_CONFIG_INDEXES = [
      // 唯一索引：品牌SN
      { fields: ["brandSN"], unique: true },
    ];

    // ========== 报表模板实体索引配置 ==========
    this.REPORT_TEMPLATE_INDEXES = [
      // 联合索引：模板名称（用于查询某模板的所有列）
      { fields: ["templateName"], unique: false },

      // 联合索引：模板名称 + 字段名（用于唯一性）
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

  /**
   * 获取实体的索引配置
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
   * 获取推荐查询使用的索引字段
   */
  getRecommendedIndex(entityName, queryPattern) {
    const indexes = this.getIndexes(entityName);

    // 按字段匹配度排序
    return indexes
      .map((idx) => ({
        ...idx,
        matchScore: this._calculateMatchScore(idx.fields, queryPattern),
      }))
      .filter((idx) => idx.matchScore > 0)
      .sort((a, b) => b.matchScore - a.matchScore);
  }

  /**
   * 计算索引匹配度
   */
  _calculateMatchScore(indexFields, queryFields) {
    const matchCount = indexFields.filter((f) =>
      queryFields.includes(f),
    ).length;
    return matchCount / Math.max(indexFields.length, queryFields.length);
  }
}

// 单例导出
const indexConfig = IndexConfig.getInstance();
