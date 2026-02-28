class ProfitCalculator {
  static _instance = null;

  constructor(repository) {
    if (ProfitCalculator._instance) {
      return ProfitCalculator._instance;
    }

    this._repository = repository || Repository.getInstance();

    ProfitCalculator._instance = this;
  }

  // 单例模式获取转化器对象
  static getInstance(repository) {
    if (!ProfitCalculator._instance) {
      ProfitCalculator._instance = new ProfitCalculator(repository);
    }
    return ProfitCalculator._instance;
  }

  // 获取品牌配置
  getBrandConfig(brandSN) {
    const brandMap = this._repository.getBrandConfigMap();
    const brand = brandMap[brandSN];

    if (!brand) {
      throw new Error(
        `未找到品牌SN【${brandSN}】的配置信息，请检查【品牌配置】工作表`,
      );
    }

    return brand;
  }

  // 计算利润
  calculateProfit(
    brandSN,
    costPrice,
    salesPrice,
    userOperations1 = 0,
    userOperations2 = 0,
    returnRate = 0.3,
  ) {
    // 参数验证
    if (
      !brandSN ||
      costPrice == null ||
      salesPrice == null ||
      typeof costPrice === "boolean" ||
      typeof salesPrice === "boolean"
    ) {
      return null;
    }

    const cost = Number(costPrice);
    const sales = Number(salesPrice);
    const ops1 = Number(userOperations1 || 0);
    const ops2 = Number(userOperations2 || 0);
    let returnRt = Number(returnRate || 0.3);

    if (
      isNaN(cost) ||
      isNaN(sales) ||
      isNaN(ops1) ||
      isNaN(ops2) ||
      isNaN(returnRt)
    ) {
      return null;
    }

    if (cost <= 0 || sales <= 0 || ops1 < 0 || ops2 < 0 || returnRt <= 0) {
      return null;
    }

    // 退货率修正
    if (returnRt < 0.3 || returnRt >= 1) {
      returnRt = 0.3;
    }

    // 获取品牌配置
    const brand = this.getBrandConfig(brandSN);

    // 超V优惠金额
    let vipDiscount = 0;
    if (brand.vipDiscountRate > 0) {
      if (sales > 50) {
        vipDiscount = Math.round(sales * brand.vipDiscountRate);
      } else {
        vipDiscount = Number((sales * brand.vipDiscountRate).toFixed(1));
      }
    }

    // 优惠后价格
    const priceAfterCoupon = Math.max(0, sales - ops1 - ops2);

    // 毛利润
    const grossProfit = priceAfterCoupon - cost;

    // 固定费用
    const fixedCosts = brand.packagingFee + brand.shippingCost;

    // 退货相关成本
    const returnMultiplier = 1 / (1 - returnRt) - 1;
    const returnCosts = returnMultiplier * fixedCosts;
    const returnProcessing = returnRt * brand.returnProcessingFee;

    // 优惠承担成本
    const vipCost = vipDiscount * brand.vipDiscountBearingRatio;

    // 平台佣金
    const platformFee = priceAfterCoupon * brand.platformCommission;

    // 品牌佣金
    const brandCommissionBase =
      priceAfterCoupon * (1 - brand.platformCommission) - vipCost;
    const brandFee = Math.max(0, brandCommissionBase) * brand.brandCommission;

    // 最终利润
    const profit =
      grossProfit -
      fixedCosts -
      returnCosts -
      returnProcessing -
      vipCost -
      platformFee -
      brandFee;

    return Number(profit.toFixed(2));
  }

  // 计算利润率
  calculateProfitRate(
    brandSN,
    costPrice,
    salesPrice,
    userOperations1 = 0,
    userOperations2 = 0,
    returnRate = 0.3,
  ) {
    const profit = this.calculateProfit(
      brandSN,
      costPrice,
      salesPrice,
      userOperations1,
      userOperations2,
      returnRate,
    );

    if (profit == null) return null;

    return Number((profit / Number(costPrice)).toFixed(5));
  }

  // 计算活动价格
  calculateActivityPrice(silverPrice, level) {
    if (silverPrice == undefined || typeof silverPrice === "boolean")
      return null;
    const silver = Number(silverPrice);
    if (!isFinite(silver) || silver <= 0) {
      return null;
    }

    switch (level) {
      case "白金限量":
        return Math.floor(silver * 0.95);
      case "白金促":
        return silver;
      case "TOP3":
        return Number((silver / 0.9 + 0.06).toFixed(1));
      case "黄金促":
        return Number((silver / 0.9 / 0.95 + 0.06 * 2).toFixed(1));
      case "直通车":
        return Number((silver / 0.9 / 0.95 / 0.95 + 0.06 * 3).toFixed(1));
      default:
        return null;
    }
  }

  // 验证利润率是否达标
  validateProfit(
    brandSN,
    costPrice,
    salesPrice,
    marketingPositioning,
    salesAge,
    userOperations1 = 0,
    userOperations2 = 0,
    returnRate = 0.3,
  ) {
    const profit = this.calculateProfit(
      brandSN,
      costPrice,
      salesPrice,
      userOperations1,
      userOperations2,
      returnRate,
    );
    const profitRate = this.calculateProfitRate(
      brandSN,
      costPrice,
      salesPrice,
      userOperations1,
      userOperations2,
      returnRate,
    );

    if (profit == null || profitRate == null) {
      return { valid: false, message: "无法计算利润" };
    }

    const warnings = [];

    // 新品（15天内）
    if (!salesAge || salesAge <= 15) {
      if (profit < 5) {
        warnings.push("新品毛利建议≥5元");
      }
      if (profitRate < 0.5) {
        warnings.push("新品毛利率建议≥50%");
      }
    } else {
      switch (marketingPositioning) {
        case "引流款":
          if (profitRate < 0.05) {
            warnings.push("引流款毛利率建议≥5%");
          }
          break;
        case "利润款":
          if (profit < 5) {
            warnings.push("利润款毛利建议≥5元");
          }
          if (profitRate < 0.35) {
            warnings.push("利润款毛利率建议≥35%");
          }
          break;
        case "清仓款":
          // 清仓款不验证利润
          break;
      }
    }

    return {
      valid: warnings.length === 0,
      profit,
      profitRate,
      warnings,
    };
  }
}
