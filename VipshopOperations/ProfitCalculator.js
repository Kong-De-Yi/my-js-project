/**
 * 利润计算器 - 负责商品利润和利润率的计算
 *
 * @class ProfitCalculator
 * @description 作为利润计算的核心引擎，提供以下功能：
 * - 基于品牌配置的利润计算（考虑超V优惠、平台佣金、品牌佣金等）
 * - 利润率计算
 * - 活动价格计算（基于白金价和活动等级）
 * - 利润达标验证（新品/引流款/利润款/清仓款）
 *
 * 利润计算公式涉及以下因素：
 * 1. 收入：销售价 - 用户操作优惠
 * 2. 成本：成本价
 * 3. 固定费用：包装费 + 运费
 * 4. 退货成本：退货率相关的固定费用和退货处理费
 * 5. 优惠承担：超V优惠中品牌承担的部分
 * 6. 平台佣金：平台抽成
 * 7. 品牌佣金：品牌方抽成
 *
 * 该类采用单例模式，确保全局只有一个利润计算器实例。
 *
 * @example
 * // 获取利润计算器实例
 * const calculator = ProfitCalculator.getInstance(repository);
 *
 * // 计算利润
 * const profit = calculator.calculateProfit(
 *   "A001",  // 品牌SN
 *   50,      // 成本价
 *   100,     // 销售价
 *   5,       // 用户操作优惠1
 *   3,       // 用户操作优惠2
 *   0.3      // 退货率
 * );
 *
 * // 计算活动价格
 * const activityPrice = calculator.calculateActivityPrice(100, "白金限量");
 */
class ProfitCalculator {
  /** @type {ProfitCalculator} 单例实例 */
  static _instance = null;

  /**
   * 创建利润计算器实例
   * @param {Repository} [repository] - 数据仓库实例，若不提供则自动获取
   */
  constructor(repository) {
    if (ProfitCalculator._instance) {
      return ProfitCalculator._instance;
    }

    this._repository = repository || Repository.getInstance();

    ProfitCalculator._instance = this;
  }

  /**
   * 获取利润计算器的单例实例
   * @static
   * @param {Repository} [repository] - 数据仓库实例
   * @returns {ProfitCalculator} 利润计算器实例
   */
  static getInstance(repository) {
    if (!ProfitCalculator._instance) {
      ProfitCalculator._instance = new ProfitCalculator(repository);
    }
    return ProfitCalculator._instance;
  }

  /**
   * 获取品牌配置信息
   * @param {string} brandSN - 品牌编号
   * @returns {Object} 品牌配置对象
   * @throws {Error} 未找到品牌配置时抛出
   * @description
   * 从仓库获取品牌配置映射，返回指定品牌的配置。
   * 品牌配置包含以下字段：
   * - vipDiscountRate: 超V优惠折扣率
   * - vipDiscountBearingRatio: 超V优惠承担比例
   * - packagingFee: 包装费
   * - shippingCost: 运费
   * - returnProcessingFee: 退货处理费
   * - platformCommission: 平台佣金率
   * - brandCommission: 品牌佣金率
   */
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

  /**
   * 计算商品利润
   * @param {string} brandSN - 品牌编号
   * @param {number} costPrice - 成本价
   * @param {number} salesPrice - 销售价
   * @param {number} [userOperations1=0] - 用户操作优惠1（中台操作）
   * @param {number} [userOperations2=0] - 用户操作优惠2（中台操作）
   * @param {number} [returnRate=0.3] - 退货率（默认0.3即30%）
   * @returns {number|null} 计算出的利润（保留两位小数），参数无效时返回null
   *
   * @description
   * 利润计算公式：
   *
   * 1. 超V优惠金额
   *    - 销售价 > 50：vipDiscount = round(销售价 × 超V折扣率)
   *    - 销售价 ≤ 50：vipDiscount = 销售价 × 超V折扣率（保留1位小数）
   *
   * 2. 优惠后价格
   *    priceAfterCoupon = max(0, 销售价 - ops1 - ops2)
   *
   * 3. 毛利润
   *    grossProfit = 优惠后价格 - 成本价
   *
   * 4. 固定费用
   *    fixedCosts = 包装费 + 运费
   *
   * 5. 退货相关成本
   *    - returnMultiplier = 1/(1-退货率) - 1
   *    - returnCosts = returnMultiplier × 固定费用
   *    - returnProcessing = 退货率 × 退货处理费
   *
   * 6. 优惠承担成本
   *    vipCost = 超V优惠 × 超V优惠承担比例
   *
   * 7. 平台佣金
   *    platformFee = 优惠后价格 × 平台佣金率
   *
   * 8. 品牌佣金
   *    brandCommissionBase = 优惠后价格 × (1 - 平台佣金率) - vipCost
   *    brandFee = max(0, brandCommissionBase) × 品牌佣金率
   *
   * 9. 最终利润
   *    profit = 毛利润 - 固定费用 - 退货成本 - 退货处理 - 优惠承担 - 平台佣金 - 品牌佣金
   *
   * @example
   * // 计算利润
   * const profit = calculator.calculateProfit("A001", 50, 100, 5, 3, 0.3);
   * // 返回例如 15.23
   */
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

  /**
   * 计算商品利润率
   * @param {string} brandSN - 品牌编号
   * @param {number} costPrice - 成本价
   * @param {number} salesPrice - 销售价
   * @param {number} [userOperations1=0] - 用户操作优惠1
   * @param {number} [userOperations2=0] - 用户操作优惠2
   * @param {number} [returnRate=0.3] - 退货率
   * @returns {number|null} 利润率（保留5位小数），计算失败时返回null
   * @description
   * 利润率 = 利润 / 成本价
   * 保留5位小数，便于展示百分比（如0.35000表示35%）
   */
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

  /**
   * 计算活动价格
   * @param {number} silverPrice - 白金价
   * @param {string} level - 活动等级
   * @returns {number|null} 活动价格（保留1位小数），参数无效时返回null
   * @description
   * 活动等级与计算公式：
   *
   * - 白金限量：floor(白金价 × 0.95)          // 向下取整
   * - 白金促：白金价
   * - TOP3： (白金价 / 0.9 + 0.06).toFixed(1) // 保留1位小数
   * - 黄金促： (白金价 / 0.9 / 0.95 + 0.06 × 2).toFixed(1)
   * - 直通车： (白金价 / 0.9 / 0.95 / 0.95 + 0.06 × 3).toFixed(1)
   *
   * @example
   * calculateActivityPrice(100, "白金限量") // 返回 95
   * calculateActivityPrice(100, "TOP3")     // 返回 111.2 (100/0.9+0.06)
   */
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

  /**
   * 验证利润率是否达标
   * @param {string} brandSN - 品牌编号
   * @param {number} costPrice - 成本价
   * @param {number} salesPrice - 销售价
   * @param {string} marketingPositioning - 营销定位（引流款/利润款/清仓款）
   * @param {number} salesAge - 售龄（上架天数）
   * @param {number} [userOperations1=0] - 用户操作优惠1
   * @param {number} [userOperations2=0] - 用户操作优惠2
   * @param {number} [returnRate=0.3] - 退货率
   * @returns {Object} 验证结果
   * @returns {boolean} return.valid - 是否达标
   * @returns {number} return.profit - 计算的利润
   * @returns {number} return.profitRate - 计算的利润率
   * @returns {string[]} return.warnings - 不达标的警告信息
   *
   * @description
   * 达标标准：
   *
   * 1. 新品（售龄 ≤ 15天）
   *    - 利润 ≥ 5元
   *    - 利润率 ≥ 50%
   *
   * 2. 引流款
   *    - 利润率 ≥ 5%
   *
   * 3. 利润款
   *    - 利润 ≥ 5元
   *    - 利润率 ≥ 35%
   *
   * 4. 清仓款
   *    - 不验证利润
   *
   * @example
   * const result = calculator.validateProfit(
   *   "A001", 50, 100, "利润款", 30, 5, 3, 0.3
   * );
   * if (!result.valid) {
   *   console.log(result.warnings); // ["利润款毛利建议≥5元"]
   * }
   */
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
