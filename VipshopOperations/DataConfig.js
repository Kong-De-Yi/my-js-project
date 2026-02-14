// ============================================================================
// 数据配置中心
// 功能：集中管理所有业务实体的字段定义
// ============================================================================

class DataConfig {
  static _instance = null;

  constructor() {
    if (DataConfig._instance) {
      return DataConfig._instance;
    }

    // ========== 1. 货号总表实体配置 ==========
    this.PRODUCT = {
      worksheet: "货号总表",

      fields: {
        // 货号
        itemNumber: {
          title: "货号",
          type: "string",
          unique: true,
        },
        // 款号
        styleNumber: { title: "款号", type: "string" },
        // 颜色
        color: { title: "颜色", type: "string" },
        // 链接
        link: {
          title: "链接",
          type: "computed",
          compute: (obj) =>
            obj.MID
              ? `https://detail.vip.com/detail-1234-${obj.MID}.html`
              : undefined,
        },
        // 首次上架时间
        firstListingTime: {
          title: "首次上架时间",
          type: "string",
          validators: [{ type: "date" }],
        },
        // 售龄
        salesAge: {
          title: "售龄",
          type: "computed",
          compute: (obj) => {
            if (!obj.firstListingTime) return undefined;
            const ts = Date.parse(obj.firstListingTime);
            return Math.floor((Date.now() - ts) / 86400000);
          },
        },
        // 商品状态
        itemStatus: {
          title: "商品状态",
          type: "string",
          validators: [
            {
              type: "enum",
              params: { values: ["商品上线", "部分上线", "商品下线"] },
            },
          ],
        },
        // 活动状态
        activityStatus: {
          title: "活动状态",
          type: "computed",
          compute: (obj) => {
            switch (obj.finalPrice) {
              case obj.directTrainPrice:
                return "直通车";
              case obj.goldPrice:
                return "黄金促";
              case obj.goldLimit:
                return "黄金限量";
              case obj.silverPrice:
                return "白金促";
              case obj.silverLimit:
                return "白金限量";
            }
            return undefined;
          },
        },
        // 图片
        picture: { title: "图片", type: "string" },
        // 设计号
        designNumber: { title: "设计号", type: "string" },
        // 通货款号
        generalGoodsStyleNumber: {
          title: "通货款号",
          type: "string",
        },
        // 上市年份
        listingYear: {
          title: "上市年份",
          type: "number",
          validators: [
            {
              type: "enum",
              params: { values: [2023, 2024, 2025, 2026, 2027, 2028] },
            },
          ],
        },
        // 主销季节
        mainSalesSeason: {
          title: "主销季节",
          type: "string",
          validators: [
            { type: "enum", params: { values: ["春秋", "夏", "冬", "四季"] } },
          ],
        },
        // 适用性别
        applicableGender: {
          title: "适用性别",
          type: "string",
          validators: [
            { type: "enum", params: { values: ["男童", "女童", "中性"] } },
          ],
        },
        // 四级品类
        fourthLevelCategory: { title: "四级品类", type: "string" },
        // 运营分类
        operationClassification: { title: "运营分类", type: "string" },
        // 备货模式
        stockingMode: {
          title: "备货模式",
          type: "string",
          validators: [
            {
              type: "enum",
              params: { values: ["现货", "通版通货", "专版通货"] },
            },
          ],
        },
        // 下线原因
        offlineReason: {
          title: "下线原因",
          type: "string",
          validators: [
            {
              type: "enum",
              params: {
                values: [
                  "新品下架",
                  "过季下架",
                  "更换吊牌",
                  "转移品牌",
                  "清仓淘汰",
                  "内网撞款",
                  "资质问题",
                  "内在质检",
                ],
              },
            },
          ],
        },
        // 营销定位
        marketingPositioning: {
          title: "营销定位",
          type: "string",
          validators: [
            {
              type: "enum",
              params: { values: ["引流款", "利润款", "清仓款"] },
            },
          ],
        },
        // 营销备忘录
        marketingMemorandum: { title: "营销备忘录", type: "string" },
        // 成本价
        costPrice: {
          title: "成本价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 最低价
        lowestPrice: {
          title: "最低价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 白金价
        silverPrice: {
          title: "白金价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 中台1
        userOperations1: {
          title: "中台1",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        // 中台2
        userOperations2: {
          title: "中台2",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        // 到手价
        finalPrice: {
          title: "到手价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 首单价
        firstOrderPrice: {
          title: "首单价",
          type: "computed",
          compute: (obj) =>
            obj.finalPrice
              ? obj.finalPrice - (obj.userOperations1 || 0)
              : undefined,
        },
        // 超V价
        superVipPrice: {
          title: "超V价",
          type: "computed",
          compute: (obj, context) => {
            if (!obj.finalPrice || !context?.brandConfig) return undefined;
            const brand = context.brandConfig[obj.brandSN];
            if (!brand) return undefined;

            let discount =
              obj.finalPrice > 50
                ? Math.round(obj.finalPrice * brand.vipDiscountRate)
                : Number((obj.finalPrice * brand.vipDiscountRate).toFixed(1));

            return obj.finalPrice - discount - (obj.userOperations1 || 0);
          },
        },
        // 是否破价
        isPriceBroken: {
          title: "是否破价",
          type: "computed",
          compute: (obj) => {
            if (obj.finalPrice && obj.lowestPrice) {
              return obj.lowestPrice > obj.finalPrice ? "是" : undefined;
            }
            return "(未知)";
          },
        },
        // 利润
        profit: {
          title: "利润",
          type: "computed",
          compute: (obj, context) => {
            return context?.profitCalculator?.calculateProfit(
              obj.brandSN,
              obj.costPrice,
              obj.finalPrice,
              obj.userOperations1,
              obj.userOperations2,
              obj.rejectAndReturnRateOfLast7Days,
            );
          },
        },
        // 利润率
        profitRate: {
          title: "利润率",
          type: "computed",
          compute: (obj, context) => {
            const profit = context?.profitCalculator?.calculateProfit(
              obj.brandSN,
              obj.costPrice,
              obj.finalPrice,
              obj.userOperations1,
              obj.userOperations2,
              obj.rejectAndReturnRateOfLast7Days,
            );
            return profit && obj.costPrice
              ? Number((profit / obj.costPrice).toFixed(5))
              : undefined;
          },
        },
        // 直通车
        directTrainPrice: {
          title: "直通车",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 黄金促
        goldPrice: {
          title: "黄金促",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 黄金限量
        goldLimit: {
          title: "黄金限量",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // TOP3
        top3: {
          title: "TOP3",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 白金限量
        silverLimit: {
          title: "白金限量",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 可售库存
        sellableInventory: {
          title: "可售库存",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        // 可售天数
        sellableDays: {
          title: "可售天数",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        // 是否断码
        isOutOfStock: {
          title: "是否断码",
          type: "string",
          validators: [{ type: "enum", params: { values: ["是", undefined] } }],
        },
        // 合计库存
        totalInventory: {
          title: "合计库存",
          type: "computed",
          compute: (obj) => {
            obj.finishedGoodsMainInventory +
              obj.finishedGoodsIncomingInventory +
              obj.finishedGoodsFinishingInventory +
              obj.finishedGoodsOversoldInventory +
              obj.finishedGoodsPrepareInventory +
              obj.finishedGoodsReturnInventory +
              obj.finishedGoodsPurchaseInventory +
              obj.generalGoodsMainInventory +
              obj.generalGoodsIncomingInventory +
              obj.generalGoodsFinishingInventory +
              obj.generalGoodsOversoldInventory +
              obj.generalGoodsPrepareInventory +
              obj.generalGoodsReturnInventory +
              obj.generalGoodsPurchaseInventory;
          },
        },
        // 成品合计
        finishedGoodsTotalInventory: {
          title: "成品合计",
          type: "computed",
          compute: (obj) => {
            obj.finishedGoodsMainInventory +
              obj.finishedGoodsIncomingInventory +
              obj.finishedGoodsFinishingInventory +
              obj.finishedGoodsOversoldInventory +
              obj.finishedGoodsPrepareInventory +
              obj.finishedGoodsReturnInventory +
              obj.finishedGoodsPurchaseInventory;
          },
        },
        // 成品库存明细
        finishedGoodsMainInventory: {
          title: "成品主仓",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsIncomingInventory: {
          title: "成品进货",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsFinishingInventory: {
          title: "成品后整",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsOversoldInventory: {
          title: "成品超卖",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsPrepareInventory: {
          title: "成品备货",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsReturnInventory: {
          title: "成品销退",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsPurchaseInventory: {
          title: "成品在途",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        // 通货合计
        generalGoodsTotalInventory: {
          title: "通货合计",
          type: "computed",
          compute: (obj) => {
            obj.generalGoodsMainInventory +
              obj.generalGoodsIncomingInventory +
              obj.generalGoodsFinishingInventory +
              obj.generalGoodsOversoldInventory +
              obj.generalGoodsPrepareInventory +
              obj.generalGoodsReturnInventory +
              obj.generalGoodsPurchaseInventory;
          },
        },
        // 通货库存明细
        generalGoodsMainInventory: {
          title: "通货主仓",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsIncomingInventory: {
          title: "通货进货",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsFinishingInventory: {
          title: "通货后整",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsOversoldInventory: {
          title: "通货超卖",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsPrepareInventory: {
          title: "通货备货",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsReturnInventory: {
          title: "通货销退",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsPurchaseInventory: {
          title: "通货在途",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        // 品牌SN
        brandSN: {
          title: "品牌SN",
          type: "string",
          validators: [{ type: "required" }],
        },
        // MID
        MID: {
          title: "MID",
          type: "string",
          validators: [
            {
              type: "pattern",
              params: {
                regex: /^69\d{17}$/,
                description: "69开头的19位数字",
              },
            },
          ],
        },
        // P_SPU
        P_SPU: {
          title: "P_SPU",
          type: "string",
          validators: [
            {
              type: "pattern",
              params: {
                regex: /^SPU-[A-F0-9]{16}$/,
                description: "格式: SPU-16位大写十六进制",
              },
            },
          ],
        },
        // 三级品类
        thirdLevelCategory: {
          title: "三级品类",
          type: "string",
          validators: [
            {
              type: "enum",
              params: {
                values: [
                  "儿童文胸",
                  "儿童配饰配件",
                  "儿童礼服/演出服",
                  "儿童羽绒服",
                  "儿童棉服",
                  "儿童外套/夹克/风衣",
                  "儿童防晒服/皮肤衣",
                  "儿童毛衣/线衫",
                  "儿童套装",
                  "儿童卫衣",
                  "儿童T恤/POLO衫",
                  "儿童背心",
                  "儿童衬衫",
                  "儿童马甲",
                  "儿童裙子",
                  "儿童旗袍/汉服",
                  "儿童披肩/斗篷",
                  "儿童裤子",
                  "儿童保暖内衣/套装",
                  "儿童睡衣家居服",
                  "儿童内裤",
                  "儿童打底裤/连袜裤",
                  "儿童袜子",
                  "儿童雨衣/雨具",
                  "儿童泳装泳具",
                  "反穿衣/画画衣",
                  "儿童冲锋衣",
                  "儿童户外服",
                  "婴幼T恤",
                  "婴幼衬衫",
                  "婴幼家居服",
                  "婴幼内衣内裤",
                  "婴幼羽绒服",
                  "婴幼棉服",
                  "婴幼披风/斗篷",
                  "肚兜/肚围/护脐带",
                  "婴幼套装",
                  "哈衣/爬服/连体服",
                  "婴幼裤子",
                  "婴幼裙子",
                  "婴幼外套/风衣",
                  "婴幼毛衣/线衫",
                  "婴幼背心/马甲",
                  "婴幼卫衣/运动服",
                  "婴儿服饰礼盒",
                  "婴儿袜子",
                  "婴幼配饰",
                  "婴幼抱被/抱毯",
                  "婴幼泳装泳具",
                  "婴幼防晒服/皮肤衣",
                  "婴幼礼服/演出服",
                  "婴幼旗袍/汉服",
                ],
              },
            },
          ],
        },
        // 市场价
        tagPrice: {
          title: "市场价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        // 唯品价
        vipshopPrice: {
          title: "唯品价",
          type: "number",
          validators: [{ type: "positive" }],
        },
      },
      // 默认排序：按首次上架时间倒序
      defaultSort: (a, b) => {
        const dateA = Date.parse(a.firstListingTime) || 0;
        const dateB = Date.parse(b.firstListingTime) || 0;
        return dateB - dateA;
      },

      uniqueKey: "itemNumber",
    };

    // ========== 2. 商品价格实体 ==========
    this.PRODUCT_PRICE = {
      worksheet: "商品价格",
      requiredFields: ["itemNumber"],
      fields: {
        itemNumber: { title: "货号", type: "string" },
        designNumber: { title: "设计号", type: "string" },
        picture: { title: "图片", type: "string" },
        costPrice: {
          title: "成本价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        lowestPrice: {
          title: "最低价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        silverPrice: {
          title: "白金价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        userOperations1: {
          title: "中台1",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        userOperations2: {
          title: "中台2",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
      },
      uniqueKey: "itemNumber",
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
