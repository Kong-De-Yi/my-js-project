// ============================================================================
// 数据配置中心（精简版）
// 功能：集中管理所有业务实体的字段定义，所有字段都持久保存
// 特点：无需指定persist，所有字段默认保存
// ============================================================================

class DataConfig {
  static _instance = null;

  constructor() {
    if (DataConfig._instance) {
      return DataConfig._instance;
    }

    // ========== 1. 商品主实体 ==========
    this.PRODUCT = {
      worksheet: "货号总表",
      // 用于实体识别的必填字段
      requiredFields: ["itemNumber"],
      fields: {
        // 基础信息
        itemNumber: { title: "货号", type: "string" },
        brandSN: { title: "品牌SN", type: "string" },
        brandName: { title: "品牌名称", type: "string" },
        styleNumber: { title: "款号", type: "string" },
        designNumber: { title: "设计号", type: "string" },
        color: { title: "颜色", type: "string" },
        P_SPU: { title: "P_SPU", type: "string" },
        MID: { title: "MID", type: "string" },

        // 分类
        thirdLevelCategory: { title: "三级品类", type: "string" },
        fourthLevelCategory: { title: "四级品类", type: "string" },
        operationClassification: { title: "运营分类", type: "string" },

        // 时间
        firstListingTime: { title: "首次上架时间", type: "string" },
        listingYear: { title: "上市年份", type: "number" },
        mainSalesSeason: { title: "主销季节", type: "string" },
        applicableGender: { title: "适用性别", type: "string" },

        // 价格
        costPrice: { title: "成本价", type: "number" },
        lowestPrice: { title: "最低价", type: "number" },
        silverPrice: { title: "白金价", type: "number" },
        vipshopPrice: { title: "唯品价", type: "number" },
        finalPrice: { title: "到手价", type: "number" },

        // 运营
        userOperations1: { title: "中台1", type: "number" },
        userOperations2: { title: "中台2", type: "number" },

        // 状态
        itemStatus: { title: "商品状态", type: "string" },
        offlineReason: { title: "下线原因", type: "string" },
        marketingPositioning: { title: "营销定位", type: "string" },
        stockingMode: { title: "备货模式", type: "string" },

        // 库存
        sellableInventory: { title: "可售库存", type: "number" },
        sellableDays: { title: "可售天数", type: "number" },
        isOutOfStock: { title: "是否断码", type: "string" },

        // 成品库存
        finishedGoodsMainInventory: { title: "成品主仓", type: "number" },
        finishedGoodsIncomingInventory: { title: "成品进货", type: "number" },
        finishedGoodsFinishingInventory: { title: "成品后整", type: "number" },
        finishedGoodsOversoldInventory: { title: "成品超卖", type: "number" },
        finishedGoodsPrepareInventory: { title: "成品备货", type: "number" },
        finishedGoodsReturnInventory: { title: "成品销退", type: "number" },
        finishedGoodsPurchaseInventory: { title: "成品在途", type: "number" },
        finishedGoodsTotalInventory: { title: "成品合计", type: "number" },

        // 通货库存
        generalGoodsMainInventory: { title: "通货主仓", type: "number" },
        generalGoodsIncomingInventory: { title: "通货进货", type: "number" },
        generalGoodsFinishingInventory: { title: "通货后整", type: "number" },
        generalGoodsOversoldInventory: { title: "通货超卖", type: "number" },
        generalGoodsPrepareInventory: { title: "通货备货", type: "number" },
        generalGoodsReturnInventory: { title: "通货销退", type: "number" },
        generalGoodsPurchaseInventory: { title: "通货在途", type: "number" },
        generalGoodsTotalInventory: { title: "通货合计", type: "number" },

        totalInventory: { title: "合计库存", type: "number" },

        // 销售数据
        salesQuantityOfLast7Days: { title: "近7天销售量", type: "number" },
        salesAmountOfLast7Days: { title: "近7天销售额", type: "number" },
        unitPriceOfLast7Days: { title: "近7天件单价", type: "number" },
        exposureUVOfLast7Days: { title: "近7天曝光UV", type: "number" },
        productDetailsUVOfLast7Days: { title: "近7天商详UV", type: "number" },
        addToCartUVOfLast7Days: { title: "近7天加购UV", type: "number" },
        customerCountOfLast7Days: { title: "近7天客户数", type: "number" },
        rejectAndReturnCountOfLast7Days: {
          title: "近7天拒退数",
          type: "number",
        },
        clickThroughRateOfLast7Days: { title: "近7天点击率", type: "number" },
        addToCartRateOfLast7Days: { title: "近7天加购率", type: "number" },
        purchaseRateOfLast7Days: { title: "近7天转化率", type: "number" },
        rejectAndReturnRateOfLast7Days: {
          title: "近7天拒退率",
          type: "number",
        },
        styleSalesOfLast7Days: { title: "近7天款销量", type: "number" },

        totalSales: { title: "销量总计", type: "number" },

        // 链接（计算字段，但需要持久化？）
        link: { title: "链接", type: "string" },
      },
    };

    // ========== 2. 商品价格实体 ==========
    this.PRODUCT_PRICE = {
      worksheet: "商品价格",
      requiredFields: ["itemNumber"],
      fields: {
        itemNumber: { title: "货号", type: "string" },
        designNumber: { title: "设计号", type: "string" },
        picture: { title: "图片", type: "string" },
        costPrice: { title: "成本价", type: "number" },
        lowestPrice: { title: "最低价", type: "number" },
        silverPrice: { title: "白金价", type: "number" },
        userOperations1: { title: "中台1", type: "number" },
        userOperations2: { title: "中台2", type: "number" },
      },
    };

    // ========== 3. 常态商品实体 ==========
    this.REGULAR_PRODUCT = {
      worksheet: "常态商品",
      requiredFields: ["productCode", "itemNumber"],
      fields: {
        productCode: { title: "条码", type: "string" },
        itemNumber: { title: "货号", type: "string" },
        styleNumber: { title: "款号", type: "string" },
        color: { title: "颜色", type: "string" },
        size: { title: "尺码", type: "string" },
        thirdLevelCategory: { title: "三级品类", type: "string" },
        brandSN: { title: "品牌SN", type: "string" },
        brand: { title: "品牌名称", type: "string" },
        sizeStatus: { title: "尺码状态", type: "string" },
        itemStatus: { title: "商品状态", type: "string" },
        vipshopPrice: { title: "唯品价", type: "number" },
        finalPrice: { title: "到手价", type: "number" },
        sellableInventory: { title: "可售库存", type: "number" },
        sellableDays: { title: "可售天数", type: "number" },
        P_SPU: { title: "P_SPU", type: "string" },
        MID: { title: "MID", type: "string" },
      },
    };

    // ========== 4. 库存实体 ==========
    this.INVENTORY = {
      worksheet: "商品库存",
      requiredFields: ["productCode"],
      fields: {
        productCode: { title: "商品编码", type: "string" },
        mainInventory: { title: "数量", type: "number" },
        incomingInventory: { title: "进货仓库存", type: "number" },
        finishingInventory: { title: "后整车间", type: "number" },
        oversoldInventory: { title: "超卖车间", type: "number" },
        prepareInventory: { title: "备货车间", type: "number" },
        returnInventory: { title: "销退仓库存", type: "number" },
        purchaseInventory: { title: "采购在途数", type: "number" },
      },
    };

    // ========== 5. 组合商品实体 ==========
    this.COMBO_PRODUCT = {
      worksheet: "组合商品",
      requiredFields: ["productCode", "subProductCode"],
      fields: {
        productCode: { title: "组合商品实体编码", type: "string" },
        subProductCode: { title: "商品编码", type: "string" },
        subProductQuantity: { title: "数量", type: "number" },
      },
    };

    // ========== 6. 商品销售实体 ==========
    this.PRODUCT_SALES = {
      worksheet: "商品销售",
      requiredFields: ["salesDate", "itemNumber"],
      fields: {
        salesDate: { title: "日期", type: "string" },
        itemNumber: { title: "货号", type: "string" },
        exposureUV: { title: "曝光UV", type: "number" },
        productDetailsUV: { title: "商详UV", type: "number" },
        addToCartUV: { title: "加购UV", type: "number" },
        customerCount: { title: "客户数", type: "number" },
        rejectAndReturnCount: { title: "拒退件数", type: "number" },
        salesQuantity: { title: "销售量", type: "number" },
        salesAmount: { title: "销售额", type: "number" },
        firstListingTime: { title: "首次上架时间", type: "string" },
      },
    };

    // ========== 7. 系统记录实体 ==========
    this.SYSTEM_RECORD = {
      worksheet: "系统记录",
      requiredFields: [], // 系统记录只有一行，不需要识别
      fields: {
        updateDateOfLast7Days: { title: "近7天数据更新日期", type: "string" },
        updateDateOfProductPrice: { title: "商品价格更新日期", type: "string" },
        updateDateOfRegularProduct: {
          title: "常态商品更新日期",
          type: "string",
        },
        updateDateOfInventory: { title: "商品库存更新日期", type: "string" },
        updateDateOfProductSales: { title: "商品销售更新日期", type: "string" },
      },
    };

    // ========== 8. 品牌配置实体 ==========
    this.BRAND_CONFIG = {
      worksheet: "品牌配置",
      requiredFields: ["brandSN"],
      fields: {
        brandSN: { title: "品牌SN", type: "string" },
        brandName: { title: "品牌名称", type: "string" },
        packagingFee: { title: "打包费", type: "number" },
        shippingCost: { title: "运费", type: "number" },
        returnProcessingFee: { title: "退货整理费", type: "number" },
        vipDiscountRate: { title: "超V折扣率", type: "number" },
        vipDiscountBearingRatio: { title: "超V承担比", type: "number" },
        platformCommission: { title: "平台扣点", type: "number" },
        brandCommission: { title: "品牌扣点", type: "number" },
      },
    };

    // ========== 9. 导入数据实体（用于识别）==========
    this.IMPORT_DATA = {
      worksheet: "导入数据",
      fields: {}, // 不预定义字段，动态识别
    };

    // ========== 10. 报表模板配置实体 ==========
    this.REPORT_TEMPLATE = {
      worksheet: "报表配置",
      requiredFields: ["templateName", "fieldName"],
      fields: {
        templateName: { title: "模板名称", type: "string" },
        fieldName: { title: "字段", type: "string" },
        columnTitle: { title: "标题", type: "string" },
        columnWidth: { title: "宽度", type: "number" },
        isVisible: { title: "显示", type: "string" },
        displayOrder: { title: "顺序", type: "number" },
        numberFormat: { title: "格式", type: "string" },
        description: { title: "说明", type: "string" },
        // 标题颜色直接从Excel单元格背景色读取，不配置字段
      },
    };

    DataConfig._instance = this;
  }

  static getInstance() {
    if (!DataConfig._instance) {
      DataConfig._instance = new DataConfig();
    }
    return DataConfig._instance;
  }

  /**
   * 获取所有实体配置
   */
  getAll() {
    return {
      Product: this.PRODUCT,
      ProductPrice: this.PRODUCT_PRICE,
      RegularProduct: this.REGULAR_PRODUCT,
      Inventory: this.INVENTORY,
      ComboProduct: this.COMBO_PRODUCT,
      ProductSales: this.PRODUCT_SALES,
      SystemRecord: this.SYSTEM_RECORD,
      BrandConfig: this.BRAND_CONFIG,
      ImportData: this.IMPORT_DATA,
      ReportTemplate: this.REPORT_TEMPLATE,
    };
  }

  /**
   * 获取指定实体配置
   */
  get(entityName) {
    return this[entityName] || this.getAll()[entityName];
  }

  /**
   * 获取所有字段标题映射
   */
  getFieldTitles(entityName) {
    const entity = this.get(entityName);
    if (!entity) return {};

    const titles = {};
    Object.entries(entity.fields).forEach(([key, config]) => {
      titles[key] = config.title || key;
    });
    return titles;
  }

  /**
   * 根据标题查找字段名
   */
  findFieldByTitle(entityName, title) {
    const entity = this.get(entityName);
    if (!entity) return null;

    for (const [key, config] of Object.entries(entity.fields)) {
      if (config.title === title) {
        return key;
      }
    }
    return null;
  }
}

// 单例导出
const dataConfig = DataConfig.getInstance();
