// ============================================================================
// 数据配置中心（精简版）
// 功能：集中管理所有业务实体的字段定义、验证规则、计算逻辑
// 特点：纯配置类，无冗余常量，直接使用字符串标识类型
// ============================================================================

class DataConfig {
  static _instance = null;

  constructor() {
    if (DataConfig._instance) {
      return DataConfig._instance;
    }

    // ========== 1. 商品主实体配置 ==========
    this.PRODUCT = {
      worksheet: "货号总表",
      fields: {
        // 唯一标识
        itemNumber: {
          title: "货号",
          type: "string",
          persist: true,
          unique: true,
          validators: [{ type: "required", message: "货号不能为空" }],
        },
        brandSN: {
          title: "品牌SN",
          type: "string",
          persist: true,
          validators: [{ type: "required" }],
        },
        styleNumber: { title: "款号", type: "string", persist: true },
        designNumber: { title: "设计号", type: "string", persist: true },
        color: { title: "颜色", type: "string", persist: true },
        P_SPU: {
          title: "P_SPU",
          type: "string",
          persist: true,
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
        MID: {
          title: "MID",
          type: "string",
          persist: true,
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

        // 分类
        thirdLevelCategory: {
          title: "三级品类",
          type: "string",
          persist: true,
        },
        fourthLevelCategory: {
          title: "四级品类",
          type: "string",
          persist: true,
        },
        operationClassification: {
          title: "运营分类",
          type: "string",
          persist: true,
        },

        // 时间
        firstListingTime: {
          title: "首次上架时间",
          type: "string",
          persist: true,
        },

        // 价格
        costPrice: {
          title: "成本价",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        lowestPrice: {
          title: "最低价",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        silverPrice: {
          title: "白金价",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        vipshopPrice: { title: "唯品价", type: "number", persist: true },
        finalPrice: { title: "到手价", type: "number", persist: true },

        // 运营
        userOperations1: {
          title: "中台1",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        userOperations2: {
          title: "中台2",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },

        // 状态
        itemStatus: {
          title: "商品状态",
          type: "string",
          persist: true,
          validators: [
            {
              type: "enum",
              params: { values: ["商品上线", "部分上线", "商品下线"] },
            },
          ],
        },
        offlineReason: {
          title: "下线原因",
          type: "string",
          persist: true,
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
        marketingPositioning: {
          title: "营销定位",
          type: "string",
          persist: true,
          validators: [
            {
              type: "enum",
              params: { values: ["引流款", "利润款", "清仓款"] },
            },
          ],
        },
        stockingMode: {
          title: "备货模式",
          type: "string",
          persist: true,
          validators: [
            {
              type: "enum",
              params: { values: ["现货", "通版通货", "专版通货"] },
            },
          ],
        },

        // 库存
        sellableInventory: { title: "可售库存", type: "number", persist: true },
        sellableDays: { title: "可售天数", type: "number", persist: true },
        isOutOfStock: { title: "是否断码", type: "string", persist: true },

        // 成品库存明细
        finishedGoodsMainInventory: {
          title: "成品主仓",
          type: "number",
          persist: true,
        },
        finishedGoodsIncomingInventory: {
          title: "成品进货",
          type: "number",
          persist: true,
        },
        finishedGoodsFinishingInventory: {
          title: "成品后整",
          type: "number",
          persist: true,
        },
        finishedGoodsOversoldInventory: {
          title: "成品超卖",
          type: "number",
          persist: true,
        },
        finishedGoodsPrepareInventory: {
          title: "成品备货",
          type: "number",
          persist: true,
        },
        finishedGoodsReturnInventory: {
          title: "成品销退",
          type: "number",
          persist: true,
        },
        finishedGoodsPurchaseInventory: {
          title: "成品在途",
          type: "number",
          persist: true,
        },
        finishedGoodsTotalInventory: {
          title: "成品合计",
          type: "number",
          persist: true,
        },

        // 通货库存明细
        generalGoodsMainInventory: {
          title: "通货主仓",
          type: "number",
          persist: true,
        },
        generalGoodsIncomingInventory: {
          title: "通货进货",
          type: "number",
          persist: true,
        },
        generalGoodsFinishingInventory: {
          title: "通货后整",
          type: "number",
          persist: true,
        },
        generalGoodsOversoldInventory: {
          title: "通货超卖",
          type: "number",
          persist: true,
        },
        generalGoodsPrepareInventory: {
          title: "通货备货",
          type: "number",
          persist: true,
        },
        generalGoodsReturnInventory: {
          title: "通货销退",
          type: "number",
          persist: true,
        },
        generalGoodsPurchaseInventory: {
          title: "通货在途",
          type: "number",
          persist: true,
        },
        generalGoodsTotalInventory: {
          title: "通货合计",
          type: "number",
          persist: true,
        },

        totalInventory: { title: "合计库存", type: "number", persist: true },

        // 销售数据
        salesQuantityOfLast7Days: {
          title: "近7天销售量",
          type: "number",
          persist: true,
        },
        salesAmountOfLast7Days: {
          title: "近7天销售额",
          type: "number",
          persist: true,
        },
        unitPriceOfLast7Days: {
          title: "近7天件单价",
          type: "number",
          persist: true,
        },
        exposureUVOfLast7Days: {
          title: "近7天曝光UV",
          type: "number",
          persist: true,
        },
        productDetailsUVOfLast7Days: {
          title: "近7天商详UV",
          type: "number",
          persist: true,
        },
        addToCartUVOfLast7Days: {
          title: "近7天加购UV",
          type: "number",
          persist: true,
        },
        customerCountOfLast7Days: {
          title: "近7天客户数",
          type: "number",
          persist: true,
        },
        rejectAndReturnCountOfLast7Days: {
          title: "近7天拒退数",
          type: "number",
          persist: true,
        },
        clickThroughRateOfLast7Days: {
          title: "近7天点击率",
          type: "number",
          persist: true,
        },
        addToCartRateOfLast7Days: {
          title: "近7天加购率",
          type: "number",
          persist: true,
        },
        purchaseRateOfLast7Days: {
          title: "近7天转化率",
          type: "number",
          persist: true,
        },
        rejectAndReturnRateOfLast7Days: {
          title: "近7天拒退率",
          type: "number",
          persist: true,
        },
        styleSalesOfLast7Days: {
          title: "近7天款销量",
          type: "number",
          persist: true,
        },

        totalSales: { title: "销量总计", type: "number", persist: true },

        // ========== 计算字段（不持久化）==========
        link: {
          title: "链接",
          type: "computed",
          persist: false,
          compute: (obj) =>
            obj.MID
              ? `https://detail.vip.com/detail-1234-${obj.MID}.html`
              : undefined,
        },
        salesAge: {
          title: "售龄",
          type: "computed",
          persist: false,
          compute: (obj) => {
            if (!obj.firstListingTime) return undefined;
            const ts = Date.parse(obj.firstListingTime);
            return isNaN(ts)
              ? undefined
              : Math.floor((Date.now() - ts) / 86400000);
          },
        },
        activityStatus: {
          title: "活动状态",
          type: "computed",
          persist: false,
          compute: (obj) => {
            if (obj.vipshopPrice && obj.finalPrice) {
              return obj.vipshopPrice > obj.finalPrice ? "活动中" : "未提报";
            }
            return "(未知)";
          },
        },
        isPriceBroken: {
          title: "是否破价",
          type: "computed",
          persist: false,
          compute: (obj) => {
            if (obj.finalPrice && obj.lowestPrice) {
              return obj.lowestPrice > obj.finalPrice ? "是" : undefined;
            }
            return "(未知)";
          },
        },
        firstOrderPrice: {
          title: "首单价",
          type: "computed",
          persist: false,
          compute: (obj) =>
            obj.finalPrice
              ? obj.finalPrice - (obj.userOperations1 || 0)
              : undefined,
        },
        superVipPrice: {
          title: "超V价",
          type: "computed",
          persist: false,
          compute: (obj, context) => {
            if (!obj.finalPrice || !context?.brandConfig) return undefined;
            const brand = context.brandConfig[obj.brandSN];
            if (!brand) return undefined;

            let discount =
              obj.finalPrice > 50
                ? Math.round(obj.finalPrice * brand.vipDiscountRate)
                : Number((obj.finalPrice * brand.vipDiscountRate).toFixed(1));

            return (
              obj.finalPrice -
              discount -
              (obj.userOperations1 || 0) -
              (obj.userOperations2 || 0)
            );
          },
        },
        profit: {
          title: "利润",
          type: "computed",
          persist: false,
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
        profitRate: {
          title: "利润率",
          type: "computed",
          persist: false,
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
      },

      // 默认排序：按首次上架时间倒序
      defaultSort: (a, b) => {
        const dateA = Date.parse(a.firstListingTime) || 0;
        const dateB = Date.parse(b.firstListingTime) || 0;
        return dateB - dateA;
      },

      uniqueKey: "itemNumber",
    };

    // ========== 2. 商品价格实体配置 ==========
    this.PRODUCT_PRICE = {
      worksheet: "商品价格",
      fields: {
        itemNumber: {
          title: "货号",
          type: "string",
          persist: true,
          unique: true,
          validators: [{ type: "required" }],
        },
        designNumber: { title: "设计号", type: "string", persist: true },
        picture: { title: "图片", type: "string", persist: true },
        costPrice: {
          title: "成本价",
          type: "number",
          persist: true,
          validators: [{ type: "required" }, { type: "positive" }],
        },
        lowestPrice: {
          title: "最低价",
          type: "number",
          persist: true,
          validators: [{ type: "required" }, { type: "positive" }],
        },
        silverPrice: {
          title: "白金价",
          type: "number",
          persist: true,
          validators: [{ type: "required" }, { type: "positive" }],
        },
        userOperations1: {
          title: "中台1",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        userOperations2: {
          title: "中台2",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
      },
      uniqueKey: "itemNumber",
    };

    // ========== 3. 常态商品实体配置 ==========
    this.REGULAR_PRODUCT = {
      worksheet: "常态商品",
      fields: {
        productCode: { title: "条码", type: "string", persist: true },
        itemNumber: { title: "货号", type: "string", persist: true },
        styleNumber: { title: "款号", type: "string", persist: true },
        color: { title: "颜色", type: "string", persist: true },
        size: { title: "尺码", type: "string", persist: true },
        thirdLevelCategory: {
          title: "三级品类",
          type: "string",
          persist: true,
        },
        brandSN: { title: "品牌SN", type: "string", persist: true },
        brand: { title: "品牌名称", type: "string", persist: true },
        sizeStatus: { title: "尺码状态", type: "string", persist: true },
        itemStatus: { title: "商品状态", type: "string", persist: true },
        vipshopPrice: { title: "唯品价", type: "number", persist: true },
        finalPrice: { title: "到手价", type: "number", persist: true },
        sellableInventory: { title: "可售库存", type: "number", persist: true },
        sellableDays: { title: "可售天数", type: "number", persist: true },
        P_SPU: { title: "P_SPU", type: "string", persist: true },
        MID: { title: "MID", type: "string", persist: true },
      },
    };

    // ========== 4. 库存实体配置 ==========
    this.INVENTORY = {
      worksheet: "商品库存",
      fields: {
        productCode: {
          title: "商品编码",
          type: "string",
          persist: true,
          validators: [{ type: "required" }],
        },
        mainInventory: {
          title: "数量",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        incomingInventory: {
          title: "进货仓库存",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        finishingInventory: {
          title: "后整车间",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        oversoldInventory: {
          title: "超卖车间",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        prepareInventory: {
          title: "备货车间",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        returnInventory: {
          title: "销退仓库存",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        purchaseInventory: {
          title: "采购在途数",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
      },
      uniqueKey: "productCode",
    };

    // ========== 5. 组合商品实体配置 ==========
    this.COMBO_PRODUCT = {
      worksheet: "组合商品",
      fields: {
        productCode: {
          title: "组合商品实体编码",
          type: "string",
          persist: true,
        },
        subProductCode: { title: "商品编码", type: "string", persist: true },
        subProductQuantity: {
          title: "数量",
          type: "number",
          persist: true,
          validators: [{ type: "positive" }],
        },
      },
    };

    // ========== 6. 商品销售实体配置 ==========
    this.PRODUCT_SALES = {
      worksheet: "商品销售",
      fields: {
        salesDate: { title: "日期", type: "string", persist: true },
        itemNumber: { title: "货号", type: "string", persist: true },
        exposureUV: { title: "曝光UV", type: "number", persist: true },
        productDetailsUV: { title: "商详UV", type: "number", persist: true },
        addToCartUV: { title: "加购UV", type: "number", persist: true },
        customerCount: { title: "客户数", type: "number", persist: true },
        rejectAndReturnCount: {
          title: "拒退件数",
          type: "number",
          persist: true,
        },
        salesQuantity: { title: "销售量", type: "number", persist: true },
        salesAmount: { title: "销售额", type: "number", persist: true },
        firstListingTime: {
          title: "首次上架时间",
          type: "string",
          persist: true,
        },
      },
    };

    // ========== 7. 系统记录实体配置 ==========
    this.SYSTEM_RECORD = {
      worksheet: "系统记录",
      fields: {
        updateDateOfLast7Days: {
          title: "近7天数据更新日期",
          type: "string",
          persist: true,
        },
        updateDateOfProductPrice: {
          title: "商品价格更新日期",
          type: "string",
          persist: true,
        },
        updateDateOfRegularProduct: {
          title: "常态商品更新日期",
          type: "string",
          persist: true,
        },
        updateDateOfInventory: {
          title: "商品库存更新日期",
          type: "string",
          persist: true,
        },
        updateDateOfProductSales: {
          title: "商品销售更新日期",
          type: "string",
          persist: true,
        },
      },
    };

    // ========== 8. 品牌配置实体（从工作表读取）==========
    this.BRAND_CONFIG = {
      worksheet: "品牌配置",
      fields: {
        brandSN: {
          title: "品牌SN",
          type: "string",
          persist: true,
          validators: [{ type: "required" }],
        },
        brandName: { title: "品牌名称", type: "string", persist: true },
        packagingFee: {
          title: "打包费",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        shippingCost: {
          title: "运费",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        returnProcessingFee: {
          title: "退货整理费",
          type: "number",
          persist: true,
          validators: [{ type: "nonNegative" }],
        },
        vipDiscountRate: {
          title: "超V折扣率",
          type: "number",
          persist: true,
          validators: [{ type: "range", params: { min: 0, max: 1 } }],
        },
        vipDiscountBearingRatio: {
          title: "超V承担比",
          type: "number",
          persist: true,
          validators: [{ type: "range", params: { min: 0, max: 1 } }],
        },
        platformCommission: {
          title: "平台扣点",
          type: "number",
          persist: true,
          validators: [{ type: "range", params: { min: 0, max: 1 } }],
        },
        brandCommission: {
          title: "品牌扣点",
          type: "number",
          persist: true,
          validators: [{ type: "range", params: { min: 0, max: 1 } }],
        },
      },
      uniqueKey: "brandSN",
    };

    DataConfig._instance = this;
  }

  static getInstance() {
    if (!DataConfig._instance) {
      DataConfig._instance = new DataConfig();
    }
    return DataConfig._instance;
  }

  // 获取所有实体配置
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
    };
  }

  // 获取指定实体配置
  get(entityName) {
    return this[entityName] || this.getAll()[entityName];
  }
}
