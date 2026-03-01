/**
 * 数据配置 - 定义所有业务实体的字段结构、验证规则和导入导出配置
 *
 * @class DataConfig
 * @description 作为系统的配置中心，提供以下功能：
 * - 定义所有业务实体的字段结构（字段名、标题、类型、验证规则、默认值）
 * - 配置实体的导入/导出规则（可导入、导入模式、必填字段）
 * - 配置实体的唯一键（单字段/复合主键）
 * - 定义计算字段的计算逻辑
 * - 提供实体配置的查询接口
 *
 * 字段类型说明：
 * - string: 字符串类型
 * - number: 数字类型
 * - date: 日期类型
 * - computed: 计算字段（不持久化，运行时通过 compute 函数计算）
 *
 * 验证规则类型：
 * - required: 必填项验证
 * - enum: 枚举值验证
 * - pattern: 正则表达式验证
 * - range: 数值范围验证
 * - nonNegative: 非负数验证
 * - positive: 正数验证
 * - number: 数字格式验证
 * - date: 日期格式验证
 * - year: 年份范围验证
 * - month: 月份验证（1-12）
 * - week: 周数验证（1-53）
 *
 * 导入模式：
 * - overwrite: 覆盖模式（直接替换目标工作表全部数据）
 * - append: 追加模式（基于主键进行新增或更新）
 *
 * 该类采用单例模式，确保全局只有一个配置实例。
 *
 * @example
 * // 获取配置实例
 * const config = DataConfig.getInstance();
 *
 * // 获取Product实体配置
 * const productConfig = config.get("Product");
 *
 * // 获取可导入的实体列表
 * const importable = config.getImportableEntities();
 *
 * // 解析唯一键
 * const uniqueKey = config.parseUniqueKey(["itemNumber", "salesDate"]);
 */
class DataConfig {
  /** @type {DataConfig} 单例实例 */
  static _instance = null;

  /**
   * 创建数据配置实例
   * @private
   */
  constructor() {
    if (DataConfig._instance) {
      return DataConfig._instance;
    }

    // 类型转换器
    this._converter = Converter.getInstance();

    this.APP_NAME = "商品运营表";

    // ========== 1. 货号总表实体配置 ==========
    this.PRODUCT = {
      worksheet: "货号总表",
      uniqueKey: "itemNumber",

      fields: {
        itemNumber: {
          title: "货号",
          type: "string",
          validators: [{ type: "required" }],
        },
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
          type: "date",
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
              case obj.vipshopPrice:
                return "未提报";
              case obj.directTrainPrice:
                return "直通车";
              case obj.goldPrice:
                return "黄金促";
              case obj.silverPrice:
                return "白金促";
              case obj.top3:
                return "TOP3";
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
          validators: [{ type: "year" }],
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
          default: "利润款",
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
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        userOperations2: {
          title: "中台2",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        vipshopPrice: {
          title: "唯品价",
          type: "number",
          validators: [{ type: "positive" }],
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
        rejectAndReturnRate: {
          title: "拒退率",
          type: "number",
          validators: [{ type: "range", params: { min: 0, max: 1 } }],
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
              obj.rejectAndReturnRate,
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
              obj.rejectAndReturnRate,
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
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        sellableDays: {
          title: "可售天数",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        isOutOfStock: {
          title: "是否断码",
          type: "string",
        },
        totalInventory: {
          title: "合计库存",
          type: "computed",
          compute: (obj) =>
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
            obj.generalGoodsPurchaseInventory,
        },
        finishedGoodsTotalInventory: {
          title: "成品合计",
          type: "computed",
          compute: (obj) =>
            obj.finishedGoodsMainInventory +
            obj.finishedGoodsIncomingInventory +
            obj.finishedGoodsFinishingInventory +
            obj.finishedGoodsOversoldInventory +
            obj.finishedGoodsPrepareInventory +
            obj.finishedGoodsReturnInventory +
            obj.finishedGoodsPurchaseInventory,
        },
        finishedGoodsMainInventory: {
          title: "成品主仓",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsIncomingInventory: {
          title: "成品进货",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsFinishingInventory: {
          title: "成品后整",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsOversoldInventory: {
          title: "成品超卖",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsPrepareInventory: {
          title: "成品备货",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsReturnInventory: {
          title: "成品销退",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        finishedGoodsPurchaseInventory: {
          title: "成品在途",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsTotalInventory: {
          title: "通货合计",
          type: "computed",
          compute: (obj) =>
            obj.generalGoodsMainInventory +
            obj.generalGoodsIncomingInventory +
            obj.generalGoodsFinishingInventory +
            obj.generalGoodsOversoldInventory +
            obj.generalGoodsPrepareInventory +
            obj.generalGoodsReturnInventory +
            obj.generalGoodsPurchaseInventory,
        },
        generalGoodsMainInventory: {
          title: "通货主仓",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsIncomingInventory: {
          title: "通货进货",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsFinishingInventory: {
          title: "通货后整",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsOversoldInventory: {
          title: "通货超卖",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsPrepareInventory: {
          title: "通货备货",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsReturnInventory: {
          title: "通货销退",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        generalGoodsPurchaseInventory: {
          title: "通货在途",
          type: "number",
          default: 0,
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
      },

      defaultSort: (a, b) => {
        const dateA = Date.parse(a.firstListingTime) || 0;
        const dateB = Date.parse(b.firstListingTime) || 0;
        return dateB - dateA;
      },
    };

    // ========== 2. 常态商品实体 ==========
    this.REGULAR_PRODUCT = {
      worksheet: "常态商品",

      canImport: true,
      importMode: "overwrite",
      requiredTitles: [
        "条码",
        "货号",
        "款号",
        "颜色",
        "尺码",
        "三级品类",
        "品牌SN",
        "尺码状态",
        "商品状态",
        "市场价",
        "唯品价",
        "到手价",
        "可售库存",
        "可售天数",
        "商品ID",
        "P_SPU",
      ],
      importDate: "importDateOfRegularProduct",

      canUpdate: true,
      updateDate: "updateDateOfRegularProduct",

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
          validators: [
            { type: "required" },
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
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        sellableDays: {
          title: "可售天数",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        MID: {
          title: "商品ID",
          type: "string",
          validators: [
            { type: "required" },
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
            { type: "required" },
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

    // ========== 3. 商品价格实体 ==========
    this.PRODUCT_PRICE = {
      worksheet: "商品价格",

      canImport: true,
      importMode: "append",
      requiredTitles: ["货号", "成本价", "最低价", "白金价", "中台1", "中台2"],
      importDate: "importDateOfProductPrice",

      canUpdate: true,
      updateDate: "updateDateOfProductPrice",

      uniqueKey: "itemNumber",

      fields: {
        itemNumber: {
          title: "货号",
          type: "string",
          validators: [{ type: "required" }],
        },

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
          default: 0,
          validators: [{ type: "nonNegative" }],
        },

        userOperations2: {
          title: "中台2",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
      },
    };

    // ========== 4. 库存实体 ==========
    this.INVENTORY = {
      worksheet: "商品库存",

      canImport: true,
      importMode: "overwrite",
      requiredTitles: [
        "商品编码",
        "数量",
        "进货仓库存",
        "后整车间",
        "超卖车间",
        "备货车间",
        "销退仓库存",
        "采购在途数",
      ],
      importDate: "importDateOfInventory",

      canUpdate: true,
      updateDate: "updateDateOfInventory",

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
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        incomingInventory: {
          title: "进货仓库存",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        finishingInventory: {
          title: "后整车间",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        oversoldInventory: {
          title: "超卖车间",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        prepareInventory: {
          title: "备货车间",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        returnInventory: {
          title: "销退仓库存",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        purchaseInventory: {
          title: "采购在途数",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
      },
    };

    // ========== 5. 组合商品实体 ==========
    this.COMBO_PRODUCT = {
      worksheet: "组合商品",

      canImport: true,
      importMode: "overwrite",
      requiredTitles: ["组合商品实体编码", "商品编码", "数量"],
      importDate: "importDateOfComboProduct",

      uniqueKey: {
        fields: ["productCode", "subProductCode"],
        message: "组合商品实体编码与子商品编码组合必须唯一",
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
          default: 1,
          validators: [{ type: "range", params: { min: 1 } }],
        },
      },
    };

    // ========== 6. 商品销售实体==========
    this.PRODUCT_SALES = {
      worksheet: "商品销售",

      canImport: true,
      importMode: "append",
      requiredTitles: [
        "日期",
        "货号",
        "曝光UV",
        "商详UV",
        "加购UV(加购用户数)",
        "客户数",
        "拒退件数",
        "销售量",
        "销售额",
        "首次上架时间",
      ],
      importDate: "importDateOfProductSales",

      canUpdate: true,
      updateDate: "updateDateOfProductSales",

      uniqueKey: {
        fields: ["itemNumber", "salesDate"],
        message: "同一货号同一天的销售数据只能有一条",
      },

      fields: {
        // ----- 基础字段 -----
        salesDate: {
          title: "日期",
          type: "date",
          validators: [{ type: "required" }, { type: "date" }],
        },
        itemNumber: {
          title: "货号",
          type: "string",
          validators: [{ type: "required" }],
        },
        exposureUV: {
          title: "曝光UV",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        productDetailsUV: {
          title: "商详UV",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        addToCartUV: {
          title: "加购UV(加购用户数)",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        customerCount: {
          title: "客户数",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        rejectAndReturnCount: {
          title: "拒退件数",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        salesQuantity: {
          title: "销售量",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        salesAmount: {
          title: "销售额",
          type: "number",
          default: 0,
          validators: [{ type: "nonNegative" }],
        },
        firstListingTime: {
          title: "首次上架时间",
          type: "date",
          validators: [{ type: "date" }],
        },

        // ========== 索引字段（计算字段，用于快速查询）==========

        /**
         * 所属年份
         * 用途：快速筛选某年的销售数据
         */
        salesYear: {
          title: "所属年份",
          type: "computed",
          persist: false,
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = this._converter.parseDate(obj.salesDate);
            return date ? date.getFullYear() : undefined;
          },
        },

        yearMonth: {
          title: "年月",
          type: "computed",
          persist: false,
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = this._converter.parseDate(obj.salesDate);
            if (!date) return undefined;
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, "0");
            return `${year}-${month}`;
          },
        },

        yearWeek: {
          title: "年周",
          type: "computed",
          persist: false,
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = this._converter.parseDate(obj.salesDate);
            if (!date) return undefined;
            const year = date.getFullYear();
            const week = String(this._getISOWeekNumber(date)).padStart(2, "0");
            return `${year}-${week}`;
          },
        },

        /**
         * 距今天数
         * 用途：快速筛选最近N天的数据
         */
        daysSinceSale: {
          title: "距今天数",
          type: "computed",
          persist: false,
          compute: (obj) => {
            if (!obj.salesDate) return undefined;
            const date = this._converter.parseDate(obj.salesDate);
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

      uniqueKey: "recordDate",

      fields: {
        recordDate: {
          title: "记录日期",
          type: "computed",
          compute: () => {
            return new Date();
          },
        },

        importDateOfRegularProduct: {
          title: "常态商品导入日期",
          type: "string",
        },
        importDateOfProductPrice: { title: "商品价格导入日期", type: "string" },
        importDateOfComboProduct: { title: "组合商品导入日期", type: "string" },
        importDateOfInventory: { title: "商品库存导入日期", type: "string" },
        importDateOfProductSales: { title: "商品销售导入日期", type: "string" },

        updateDateOfRegularProduct: {
          title: "常态商品更新日期",
          type: "string",
        },
        updateDateOfProductPrice: { title: "商品价格更新日期", type: "string" },
        updateDateOfInventory: { title: "商品库存更新日期", type: "string" },
        updateDateOfProductSales: { title: "商品销售更新日期", type: "string" },
      },
    };

    // ========== 8. 品牌配置实体 ==========
    this.BRAND_CONFIG = {
      worksheet: "品牌配置",

      uniqueKey: "brandSN",

      fields: {
        brandSN: {
          title: "品牌SN",
          type: "string",
          validators: [{ type: "required" }],
        },
        brandName: {
          title: "品牌名称",
          type: "string",
          validators: [{ type: "required" }],
        },
        packagingFee: {
          title: "打包费",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        shippingCost: {
          title: "运费",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        returnProcessingFee: {
          title: "退货整理费",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        vipDiscountRate: {
          title: "超V折扣率",
          type: "number",
          validators: [{ type: "required" }, { type: "nonNegative" }],
        },
        vipDiscountBearingRatio: {
          title: "超V承担比例",
          type: "number",
          validators: [{ type: "required" }, { type: "nonNegative" }],
        },
        platformCommission: {
          title: "平台扣点",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
        brandCommission: {
          title: "品牌扣点",
          type: "number",
          validators: [{ type: "required" }, { type: "positive" }],
        },
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
      uniqueKey: ["templateName", "fieldName"],

      fields: {
        templateName: { title: "模板名称", type: "string" },
        fieldName: { title: "字段", type: "string" },
        columnTitle: { title: "显示标题", type: "string" },
        columnWidth: { title: "显示宽度", type: "number" },
        isVisible: { title: "是否显示", type: "string" },
        displayOrder: { title: "显示顺序", type: "number" },
        numberFormat: { title: "显示格式", type: "string" },
        titleColor: { title: "标题颜色", type: "number" },
        description: { title: "说明", type: "string" },
      },
    };

    DataConfig._instance = this;
  }

  /**
   * 获取数据配置的单例实例
   * @static
   * @returns {DataConfig} 数据配置实例
   */
  static getInstance() {
    if (!DataConfig._instance) {
      DataConfig._instance = new DataConfig();
    }
    return DataConfig._instance;
  }

  /**
   * 解析实体唯一键配置
   * @param {string|string[]|Object} uniqueKey - 唯一键配置
   * @returns {Object} 解析结果
   * @returns {string[]} return.fields - 唯一键字段数组
   * @returns {string} return.message - 唯一键冲突时的错误信息
   * @returns {boolean} return.isComposite - 是否为复合主键
   *
   * @description
   * 支持三种配置格式：
   * 1. 字符串：单字段主键，如 "itemNumber"
   * 2. 数组：复合主键，如 ["itemNumber", "salesDate"]
   * 3. 对象：带自定义错误信息的复合主键，如 {
   *      fields: ["itemNumber", "salesDate"],
   *      message: "同一货号同一天的销售数据只能有一条"
   *    }
   *
   * @example
   * parseUniqueKey("itemNumber")
   * // 返回 { fields: ["itemNumber"], message: "字段【itemNumber】的值必须唯一", isComposite: false }
   *
   * parseUniqueKey(["itemNumber", "salesDate"])
   * // 返回 { fields: ["itemNumber", "salesDate"], message: "【itemNumber】、【salesDate】的组合值必须唯一", isComposite: true }
   */
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

  /**
   * 获取所有实体的配置对象
   * @returns {Object.<string, Object>} 实体名称到配置的映射
   *
   * @description
   * 返回的配置对象包含以下实体：
   * - Product: 货号总表
   * - RegularProduct: 常态商品
   * - ProductPrice: 商品价格
   * - ComboProduct: 组合商品
   * - Inventory: 商品库存
   * - ProductSales: 商品销售
   * - SystemRecord: 系统记录
   * - BrandConfig: 品牌配置
   * - ImportData: 导入数据
   * - ReportTemplate: 报表模板
   */
  getAll() {
    return {
      Product: this.PRODUCT,
      RegularProduct: this.REGULAR_PRODUCT,
      ProductPrice: this.PRODUCT_PRICE,
      ComboProduct: this.COMBO_PRODUCT,
      Inventory: this.INVENTORY,
      ProductSales: this.PRODUCT_SALES,
      SystemRecord: this.SYSTEM_RECORD,
      BrandConfig: this.BRAND_CONFIG,
      ImportData: this.IMPORT_DATA,
      ReportTemplate: this.REPORT_TEMPLATE,
    };
  }

  /**
   * 获取应用名称
   * @returns {string} 应用名称（用于工作簿名称验证）
   */
  getAppName() {
    return this.APP_NAME;
  }

  /**
   * 获取指定实体的配置对象
   * @param {string} entityName - 实体名称
   * @returns {Object|null} 实体配置对象，不存在时返回null
   *
   * @description
   * 实体配置对象结构：
   * - worksheet: {string} 工作表名称
   * - uniqueKey: {string|Array|Object} 唯一键配置
   * - canImport: {boolean} 是否可导入
   * - importMode: {string} 导入模式（overwrite/append）
   * - requiredTitles: {string[]} 必填字段标题列表
   * - importDate: {string} 导入日期字段名
   * - canUpdate: {boolean} 是否可更新
   * - updateDate: {string} 更新日期字段名
   * - fields: {Object} 字段配置映射
   * - defaultSort: {Function} 默认排序函数
   *
   * @example
   * const config = dataConfig.get("Product");
   * console.log(config.worksheet); // "货号总表"
   * console.log(config.fields.itemNumber.title); // "货号"
   */
  get(entityName) {
    return this[entityName] || this.getAll()[entityName];
  }

  /**
   * 获取所有可导入的实体名称
   * @returns {string[]} 可导入实体名称数组
   * @description
   * 筛选条件：实体配置中 canImport === true
   * 可导入实体包括：
   * - RegularProduct（常态商品）
   * - ProductPrice（商品价格）
   * - Inventory（商品库存）
   * - ComboProduct（组合商品）
   * - ProductSales（商品销售）
   */
  getImportableEntities() {
    const importableEntities = [];

    // 获取可导入业务实体
    for (const [key, value] of Object.entries(this.getAll())) {
      if (value?.canImport === true) {
        importableEntities.push(key);
      }
    }

    return importableEntities;
  }

  /**
   * 获取所有可更新的实体名称
   * @returns {string[]} 可更新实体名称数组
   * @description
   * 筛选条件：实体配置中 canUpdate === true
   * 可更新实体用于记录最后更新时间
   */
  getUpdatableEntities() {
    const updatableEntities = [];

    // 获取可导入业务实体
    for (const [key, value] of Object.entries(this.getAll())) {
      if (value?.canUpdate === true) {
        updatableEntities.push(key);
      }
    }

    return updatableEntities;
  }

  /**
   * 获取实体的字段与标题映射
   * @param {string} entityName - 实体名称
   * @returns {Object.<string, string>} 字段名到标题的映射
   *
   * @example
   * getFieldTitles("Product")
   * // 返回 {
   * //   itemNumber: "货号",
   * //   styleNumber: "款号",
   * //   color: "颜色",
   * //   ...
   * // }
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
   * 根据标题查找对应的字段名
   * @param {string} entityName - 实体名称
   * @param {string} title - 字段标题
   * @returns {string|null} 字段名，未找到时返回null
   * @description
   * 用于从Excel表头反向查找对应的字段名
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

  /**
   * 获取ISO周数
   * @private
   * @param {Date} date - 日期对象
   * @returns {number} ISO周数（1-53）
   * @description
   * 用于计算商品销售中的年周字段
   * ISO周数定义：
   * - 一周从周一开始
   * - 一年的第一周是包含该年第一个周四的那一周
   */
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
