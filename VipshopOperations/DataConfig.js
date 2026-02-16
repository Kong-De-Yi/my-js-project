// ============================================================================
// DataConfig.js - 数据配置中心
// 功能：集中管理所有业务实体的字段定义，支持统一的主键配置
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
      uniqueKey: "itemNumber",
      fields: {
        itemNumber: { title: "货号", type: "string" },
        styleNumber: { title: "款号", type: "string" },
        color: { title: "颜色", type: "string" },
        link: {
          title: "链接",
          type: "computed",
          compute: (obj) =>
            obj.MID
              ? `https://detail.vip.com/detail-1234-${obj.MID}.html`
              : undefined,
        },
        firstListingTime: {
          title: "首次上架时间",
          type: "string",
          validators: [{ type: "date" }],
        },
        salesAge: {
          title: "售龄",
          type: "computed",
          compute: (obj) => {
            if (!obj.firstListingTime) return undefined;
            const ts = Date.parse(obj.firstListingTime);
            return Math.floor((Date.now() - ts) / 86400000);
          },
        },
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
        picture: { title: "图片", type: "string" },
        designNumber: { title: "设计号", type: "string" },
        generalGoodsStyleNumber: { title: "通货款号", type: "string" },
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
        mainSalesSeason: {
          title: "主销季节",
          type: "string",
          validators: [
            { type: "enum", params: { values: ["春秋", "夏", "冬", "四季"] } },
          ],
        },
        applicableGender: {
          title: "适用性别",
          type: "string",
          validators: [
            { type: "enum", params: { values: ["男童", "女童", "中性"] } },
          ],
        },
        fourthLevelCategory: { title: "四级品类", type: "string" },
        operationClassification: { title: "运营分类", type: "string" },
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
        marketingMemorandum: { title: "营销备忘录", type: "string" },
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
        finalPrice: {
          title: "到手价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        firstOrderPrice: {
          title: "首单价",
          type: "computed",
          compute: (obj) =>
            obj.finalPrice
              ? obj.finalPrice - (obj.userOperations1 || 0)
              : undefined,
        },
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
        isPriceBroken: {
          title: "是否破价",
          type: "computed",
          compute: (obj) => {
            if (obj.finalPrice && obj.lowestPrice) {
              return obj.lowestPrice > obj.finalPrice ? "是" : "否";
            }
            return undefined;
          },
        },
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
        directTrainPrice: {
          title: "直通车",
          type: "number",
          validators: [{ type: "positive" }],
        },
        goldPrice: {
          title: "黄金促",
          type: "number",
          validators: [{ type: "positive" }],
        },
        goldLimit: {
          title: "黄金限量",
          type: "number",
          validators: [{ type: "positive" }],
        },
        top3: {
          title: "TOP3",
          type: "number",
          validators: [{ type: "positive" }],
        },
        silverLimit: {
          title: "白金限量",
          type: "number",
          validators: [{ type: "positive" }],
        },
        sellableInventory: {
          title: "可售库存",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        sellableDays: {
          title: "可售天数",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        isOutOfStock: {
          title: "是否断码",
          type: "string",
        },
        totalInventory: {
          title: "合计库存",
          type: "computed",
          compute: (obj) => {
            return (
              (obj.finishedGoodsMainInventory || 0) +
              (obj.finishedGoodsIncomingInventory || 0) +
              (obj.finishedGoodsFinishingInventory || 0) +
              (obj.finishedGoodsOversoldInventory || 0) +
              (obj.finishedGoodsPrepareInventory || 0) +
              (obj.finishedGoodsReturnInventory || 0) +
              (obj.finishedGoodsPurchaseInventory || 0) +
              (obj.generalGoodsMainInventory || 0) +
              (obj.generalGoodsIncomingInventory || 0) +
              (obj.generalGoodsFinishingInventory || 0) +
              (obj.generalGoodsOversoldInventory || 0) +
              (obj.generalGoodsPrepareInventory || 0) +
              (obj.generalGoodsReturnInventory || 0) +
              (obj.generalGoodsPurchaseInventory || 0)
            );
          },
        },
        finishedGoodsTotalInventory: {
          title: "成品合计",
          type: "computed",
          compute: (obj) => {
            return (
              (obj.finishedGoodsMainInventory || 0) +
              (obj.finishedGoodsIncomingInventory || 0) +
              (obj.finishedGoodsFinishingInventory || 0) +
              (obj.finishedGoodsOversoldInventory || 0) +
              (obj.finishedGoodsPrepareInventory || 0) +
              (obj.finishedGoodsReturnInventory || 0) +
              (obj.finishedGoodsPurchaseInventory || 0)
            );
          },
        },
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
        generalGoodsTotalInventory: {
          title: "通货合计",
          type: "computed",
          compute: (obj) => {
            return (
              (obj.generalGoodsMainInventory || 0) +
              (obj.generalGoodsIncomingInventory || 0) +
              (obj.generalGoodsFinishingInventory || 0) +
              (obj.generalGoodsOversoldInventory || 0) +
              (obj.generalGoodsPrepareInventory || 0) +
              (obj.generalGoodsReturnInventory || 0) +
              (obj.generalGoodsPurchaseInventory || 0)
            );
          },
        },
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
        brandSN: {
          title: "品牌SN",
          type: "string",
          validators: [{ type: "required" }],
        },
        MID: {
          title: "MID",
          type: "string",
          validators: [
            {
              type: "pattern",
              params: { regex: /^69\d{17}$/, description: "69开头的19位数字" },
            },
          ],
        },
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
        tagPrice: {
          title: "市场价",
          type: "number",
          validators: [{ type: "positive" }],
        },
        vipshopPrice: {
          title: "唯品价",
          type: "number",
          validators: [{ type: "positive" }],
        },
      },
      defaultSort: (a, b) => {
        const dateA = Date.parse(a.firstListingTime) || 0;
        const dateB = Date.parse(b.firstListingTime) || 0;
        return dateB - dateA;
      },
    };

    // ========== 2. 商品价格实体 ==========
    this.PRODUCT_PRICE = {
      worksheet: "商品价格",
      requiredFields: ["itemNumber"],
      uniqueKey: "itemNumber",
      fields: {
        designNumber: { title: "设计号", type: "string" },
        itemNumber: {
          title: "货号",
          type: "string",
          validators: [{ type: "required" }],
        },
        picture: { title: "图片", type: "string" },
        costPrice: {
          title: "成本价",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        lowestPrice: {
          title: "最低价",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        silverPrice: {
          title: "白金价",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
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
    };

    // ========== 3. 常态商品实体 ==========
    this.REGULAR_PRODUCT = {
      worksheet: "常态商品",
      requiredFields: ["productCode", "itemNumber"],
      uniqueKey: "productCode",
      fields: {
        productCode: {
          title: "条码",
          type: "string",
          validators: [{ type: "required" }],
        },
        itemNumber: {
          title: "货号",
          type: "string",
          validators: [{ type: "required" }],
        },
        styleNumber: {
          title: "款号",
          type: "string",
          validators: [{ type: "required" }],
        },
        color: {
          title: "颜色",
          type: "string",
          validators: [{ type: "required" }],
        },
        size: {
          title: "尺码",
          type: "string",
          validators: [{ type: "required" }],
        },
        thirdLevelCategory: {
          title: "三级品类",
          type: "string",
          validators: [{ type: "required" }],
        },
        brandSN: {
          title: "品牌SN",
          type: "string",
          validators: [{ type: "required" }],
        },
        sizeStatus: {
          title: "尺码状态",
          type: "string",
          validators: [
            { type: "required" },
            { type: "enum", params: { values: ["尺码上线", "尺码下线"] } },
          ],
        },
        itemStatus: {
          title: "商品状态",
          type: "string",
          validators: [
            { type: "required" },
            {
              type: "enum",
              params: { values: ["商品上线", "部分上线", "商品下线"] },
            },
          ],
        },
        tagPrice: {
          title: "市场价",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        vipshopPrice: {
          title: "唯品价",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        finalPrice: {
          title: "到手价",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        sellableInventory: {
          title: "可售库存",
          type: "number",
          validators: [{ type: "required" }, { type: "nonNegative" }],
        },
        sellableDays: {
          title: "可售天数",
          type: "number",
          validators: [{ type: "required" }, { type: "nonNegative" }],
        },
        MID: {
          title: "MID",
          type: "string",
          validators: [
            {
              type: "pattern",
              params: { regex: /^69\d{17}$/, description: "69开头的19位数字" },
            },
          ],
        },
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
      },
    };

    // ========== 4. 库存实体 ==========
    this.INVENTORY = {
      worksheet: "商品库存",
      requiredFields: ["productCode"],
      uniqueKey: "productCode",
      fields: {
        productCode: {
          title: "商品编码",
          type: "string",
          validators: [{ type: "required" }],
        },
        mainInventory: {
          title: "数量",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        incomingInventory: {
          title: "进货仓库存",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        finishingInventory: {
          title: "后整车间",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        oversoldInventory: {
          title: "超卖车间",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        prepareInventory: {
          title: "备货车间",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        returnInventory: {
          title: "销退仓库存",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
        purchaseInventory: {
          title: "采购在途数",
          type: "number",
          validators: [{ type: "nonNegative" }],
        },
      },
    };

    // ========== 5. 组合商品实体 ==========
    this.COMBO_PRODUCT = {
      worksheet: "组合商品",
      requiredFields: ["productCode", "subProductCode"],
      uniqueKey: {
        fields: ["productCode", "subProductCode"],
        message: "组合商品编码与子商品编码组合必须唯一",
      },
      fields: {
        productCode: {
          title: "组合商品实体编码",
          type: "string",
          validators: [{ type: "required" }],
        },
        subProductCode: {
          title: "商品编码",
          type: "string",
          validators: [{ type: "required" }],
        },
        subProductQuantity: {
          title: "数量",
          type: "number",
          validators: [
            { type: "required" },
            { type: "range", params: { min: 1 } },
          ],
        },
      },
    };

    // ========== 6. 商品销售实体（增强版 - 添加索引字段）==========
    this.PRODUCT_SALES = {
      worksheet: "商品销售",
      requiredFields: ["salesDate", "itemNumber"],
      uniqueKey: {
        fields: ["itemNumber", "salesDate"],
        message: "同一货号同一天的销售数据只能有一条",
      },
      fields: {
        // ----- 基础字段 -----
        salesDate: {
          title: "日期",
          type: "string",
          validators: [{ type: "date" }],
        },
        itemNumber: { title: "货号", type: "string" },
        exposureUV: { title: "曝光UV", type: "number" },
        productDetailsUV: { title: "商详UV", type: "number" },
        addToCartUV: { title: "加购UV", type: "number" },
        customerCount: { title: "客户数", type: "number" },
        rejectAndReturnCount: { title: "拒退件数", type: "number" },
        salesQuantity: { title: "销售量", type: "number" },
        salesAmount: { title: "销售额", type: "number" },
        firstListingTime: { title: "首次上架时间", type: "string" },

        // ========== 索引字段（计算字段，用于快速查询）==========

        /**
         * 所属年份
         * 用途：快速筛选某年的销售数据
         */
        salesYear: {
          title: "所属年份",
          type: "computed",
          description: "从销售日期中提取年份，用于快速筛选",
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = validationEngine.parseDate(obj.salesDate);
            return date ? date.getFullYear() : undefined;
          },
        },

        /**
         * 所属月份
         * 用途：快速筛选某月的销售数据
         */
        salesMonth: {
          title: "所属月份",
          type: "computed",
          description: "从销售日期中提取月份（1-12），用于快速筛选",
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = validationEngine.parseDate(obj.salesDate);
            return date ? date.getMonth() + 1 : undefined;
          },
        },

        /**
         * 所属年份的第几周
         * 用途：快速筛选某周的销售数据
         */
        salesWeekOfYear: {
          title: "所属周数",
          type: "computed",
          description: "ISO周数（1-53），用于快速筛选",
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = validationEngine.parseDate(obj.salesDate);
            if (!date) return undefined;

            const d = new Date(date);
            d.setHours(0, 0, 0, 0);
            d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
            const week1 = new Date(d.getFullYear(), 0, 4);
            const weekNum =
              1 +
              Math.round(
                ((d - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7,
              );
            return weekNum;
          },
        },

        /**
         * 所属季度
         * 用途：快速筛选某季度的销售数据
         */
        salesQuarter: {
          title: "所属季度",
          type: "computed",
          description: "季度（1-4），用于快速筛选",
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = validationEngine.parseDate(obj.salesDate);
            if (!date) return undefined;
            const month = date.getMonth() + 1;
            return Math.ceil(month / 3);
          },
        },

        /**
         * 年月组合
         * 用途：快速筛选某年月的销售数据
         * 格式：YYYY-MM
         */
        yearMonth: {
          title: "年月",
          type: "computed",
          description: "年月组合（YYYY-MM），用于快速筛选",
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = validationEngine.parseDate(obj.salesDate);
            if (!date) return undefined;
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, "0");
            return `${year}-${month}`;
          },
        },

        /**
         * 年周组合
         * 用途：快速筛选某年周的销售数据
         * 格式：YYYY-WW
         */
        yearWeek: {
          title: "年周",
          type: "computed",
          description: "年周组合（YYYY-WW），用于快速筛选",
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = validationEngine.parseDate(obj.salesDate);
            if (!date) return undefined;
            const year = date.getFullYear();
            const week = String(this._getISOWeekNumber(date)).padStart(2, "0");
            return `${year}-${week}`;
          },
        },

        /**
         * 是否为今年
         */
        isCurrentYear: {
          title: "是否今年",
          type: "computed",
          compute: (obj) => {
            if (!obj.salesYear) return false;
            const currentYear = new Date().getFullYear();
            return obj.salesYear === currentYear;
          },
        },

        /**
         * 是否为去年
         */
        isLastYear: {
          title: "是否去年",
          type: "computed",
          compute: (obj) => {
            if (!obj.salesYear) return false;
            const currentYear = new Date().getFullYear();
            return obj.salesYear === currentYear - 1;
          },
        },

        /**
         * 是否为前年
         */
        isYearBeforeLast: {
          title: "是否前年",
          type: "computed",
          compute: (obj) => {
            if (!obj.salesYear) return false;
            const currentYear = new Date().getFullYear();
            return obj.salesYear === currentYear - 2;
          },
        },

        /**
         * 是否为本月
         */
        isCurrentMonth: {
          title: "是否本月",
          type: "computed",
          compute: (obj) => {
            if (!obj.salesDate) return false;
            const date = validationEngine.parseDate(obj.salesDate);
            if (!date) return false;
            const today = new Date();
            return (
              date.getFullYear() === today.getFullYear() &&
              date.getMonth() === today.getMonth()
            );
          },
        },

        /**
         * 是否为上周
         */
        isLastWeek: {
          title: "是否上周",
          type: "computed",
          compute: (obj) => {
            if (!obj.salesDate) return false;
            const date = validationEngine.parseDate(obj.salesDate);
            if (!date) return false;

            const today = new Date();
            const lastWeekStart = new Date(today);
            lastWeekStart.setDate(today.getDate() - today.getDay() - 6);
            lastWeekStart.setHours(0, 0, 0, 0);

            const lastWeekEnd = new Date(lastWeekStart);
            lastWeekEnd.setDate(lastWeekStart.getDate() + 6);
            lastWeekEnd.setHours(23, 59, 59, 999);

            return date >= lastWeekStart && date <= lastWeekEnd;
          },
        },

        /**
         * 距今天数
         * 用途：快速筛选最近N天的数据
         */
        daysSinceSale: {
          title: "距今天数",
          type: "computed",
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = validationEngine.parseDate(obj.salesDate);
            if (!date) return undefined;

            const today = new Date();
            today.setHours(0, 0, 0, 0);
            date.setHours(0, 0, 0, 0);

            const diffTime = today - date;
            return Math.floor(diffTime / (1000 * 60 * 60 * 24));
          },
        },
      },
    };

    // ========== 7. 系统记录实体 ==========
    this.SYSTEM_RECORD = {
      worksheet: "系统记录",
      uniqueKey: "recordId",
      fields: {
        recordId: { title: "记录ID", type: "string", persist: false },
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
      uniqueKey: "brandSN",
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

    // ========== 9. 导入数据实体 ==========
    this.IMPORT_DATA = {
      worksheet: "导入数据",
      uniqueKey: null,
      fields: {},
    };

    // ========== 10. 报表模板配置实体 ==========
    this.REPORT_TEMPLATE = {
      worksheet: "报表配置",
      requiredFields: ["templateName", "fieldName"],
      uniqueKey: ["templateName", "fieldName"],
      fields: {
        templateName: { title: "模板名称", type: "string" },
        fieldName: { title: "字段", type: "string" },
        columnTitle: { title: "标题", type: "string" },
        columnWidth: { title: "宽度", type: "number" },
        isVisible: { title: "显示", type: "string" },
        displayOrder: { title: "顺序", type: "number" },
        numberFormat: { title: "格式", type: "string" },
        titleColor: { title: "标题颜色", type: "string", persist: false },
        description: { title: "说明", type: "string" },
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

  parseUniqueKey(uniqueKey) {
    if (!uniqueKey) {
      return { fields: [], message: "", isComposite: false };
    }

    if (typeof uniqueKey === "string") {
      return {
        fields: [uniqueKey],
        message: `字段【${uniqueKey}】的值必须唯一`,
        isComposite: false,
      };
    }

    if (Array.isArray(uniqueKey)) {
      const fieldNames = uniqueKey.map((f) => `【${f}】`).join("、");
      return {
        fields: uniqueKey,
        message: `${fieldNames}的组合值必须唯一`,
        isComposite: true,
      };
    }

    if (typeof uniqueKey === "object") {
      return {
        fields: uniqueKey.fields || [],
        message: uniqueKey.message || "组合值必须唯一",
        isComposite: true,
      };
    }

    return { fields: [], message: "", isComposite: false };
  }

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

  get(entityName) {
    return this[entityName] || this.getAll()[entityName];
  }

  getFieldTitles(entityName) {
    const entity = this.get(entityName);
    if (!entity) return {};

    const titles = {};
    Object.entries(entity.fields).forEach(([key, config]) => {
      titles[key] = config.title || key;
    });
    return titles;
  }

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
}

const dataConfig = DataConfig.getInstance();
