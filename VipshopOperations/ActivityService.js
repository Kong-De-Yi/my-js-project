/**
 * 活动提报服务 - 负责平台活动报名、校验和导出
 *
 * @class ActivityService
 * @description 作为活动提报的业务逻辑核心，提供以下功能：
 * - 数据时效性检查（常态商品5小时内、近7天数据最新、库存今日更新）
 * - 活动价格计算（基于白金价和活动等级）
 * - 多重校验机制：
 *   - 价格数据完整性校验
 *   - 破价校验（白金价低于最低价）
 *   - 同款不同价校验
 *   - 利润合规性校验
 *   - 清仓款特殊校验
 * - 白金限量去重（按SPU）
 * - 生成活动提报表
 *
 * 活动等级说明：
 * - 直通车
 * - 黄金促
 * - TOP3
 * - 白金限量
 *
 * 校验规则：
 * 1. 数据时效性：确保参与活动的数据都是最新的
 * 2. 价格完整性：必须包含成本价、最低价、白金价
 * 3. 破价检查：白金价不能低于历史最低价
 * 4. 同款统一：同款号商品应保持价格一致
 * 5. 利润合规：根据品牌、成本、定位等计算利润是否达标
 *
 * @example * // 创建活动服务实例
 * const activityService = new ActivityService(repository, profitCalculator, excelDAO);
 *
 * // 生成活动提报表
 * try {
 *   const workbook = activityService.signUpActivity();
 *   workbook.Activate();
 *   MsgBox("活动提报表生成成功！");
 * } catch (e) {
 *   MsgBox("生成失败：" + e.message);
 * }
 */
class ActivityService {
  /**
   * 创建活动提报服务实例
   * @param {Repository} repository - 数据仓库实例
   * @param {ProfitCalculator} profitCalculator - 利润计算器实例
   * @param {ExcelDAO} excelDAO - Excel数据访问对象实例
   */
  constructor(repository, profitCalculator, excelDAO) {
    this._repository = repository;
    this._profitCalculator = profitCalculator;
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();
  }

  /**
   * 检查数据的时效性
   * @private
   * @returns {boolean} 所有数据都是最新的
   * @throws {Error} 当以下情况时抛出：
   * - 常态商品超过5小时未更新
   * - 近7天销售数据不是最新
   * - 库存数据不是今日更新
   * - 缺少更新记录或日期格式错误
   *
   * @description
   * 检查规则：
   * 1. 常态商品更新记录必须在5小时内
   * 2. 近7天销售数据必须是最新的（与当前日期匹配）
   * 3. 库存数据必须是今日更新的
   */
  _checkDataFreshness() {
    const systemRecord = this._repository.getSystemRecord();
    const now = new Date();
    const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

    // 1. 常态商品更新检查
    if (!systemRecord.updateDateOfRegularProduct) {
      throw new Error("未找到常态商品更新记录，请先更新常态商品");
    }

    const regularUpdateTime = Date.parse(
      systemRecord.updateDateOfRegularProduct,
    );
    if (isNaN(regularUpdateTime)) {
      throw new Error("常态商品更新日期格式错误");
    }

    const hoursSinceUpdate = (now - regularUpdateTime) / (1000 * 60 * 60);
    if (hoursSinceUpdate > 5) {
      throw new Error("常态商品数据已超过5小时未更新，请重新导入");
    }

    // 2. 近7天数据检查
    if (systemRecord.updateDateOfLast7Days !== this._getLast7DaysRange()) {
      throw new Error("近7天销售数据不是最新，请更新商品销售");
    }

    // 3. 库存数据检查
    if (!systemRecord.updateDateOfInventory) {
      throw new Error("未找到商品库存更新记录，请先更新库存");
    }

    const inventoryUpdateTime = Date.parse(systemRecord.updateDateOfInventory);
    if (isNaN(inventoryUpdateTime)) {
      throw new Error("商品库存更新日期格式错误");
    }

    if (!this._isToday(inventoryUpdateTime)) {
      throw new Error("商品库存数据不是今日更新，请重新导入");
    }

    return true;
  }

  /**
   * 获取近7天的日期范围
   * @private
   * @returns {string} 日期范围字符串，格式：YYYY-MM-DD~YYYY-MM-DD
   * @description
   * 计算规则：
   * - 起始日期：今天往前推7天
   * - 结束日期：昨天
   *
   * @example
   * // 假设今天是2024-03-15
   * _getLast7DaysRange() // 返回 "2024-03-08~2024-03-14"
   */
  _getLast7DaysRange() {
    const today = new Date();
    const start = new Date(today);
    start.setDate(today.getDate() - 7);
    const end = new Date(today);
    end.setDate(today.getDate() - 1);

    const format = (date) =>
      `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;

    return `${format(start)}~${format(end)}`;
  }

  /**
   * 判断时间戳是否代表今天
   * @private
   * @param {number} timestamp - 时间戳
   * @returns {boolean} true=今天，false=不是今天
   */
  _isToday(timestamp) {
    const date = new Date(timestamp);
    const today = new Date();
    return (
      date.getDate() === today.getDate() &&
      date.getMonth() === today.getMonth() &&
      date.getFullYear() === today.getFullYear()
    );
  }

  /**
   * 生成活动提报表
   * @returns {Excel.Workbook} 生成的活动提报表工作簿
   * @throws {Error} 当以下情况时抛出：
   * - 未选择活动等级
   * - 数据时效性检查失败
   * - 没有符合条件的商品
   * - 价格数据校验失败
   * - 存在破价商品
   *
   * @description
   * 活动提报完整流程：
   *
   * 1. 验证活动等级
   *
   * 2. 检查数据时效性
   *
   * 3. 获取筛选条件（默认提报上架商品）
   *
   * 4. 查询待提报商品
   *
   * 5. 刷新商品价格数据（从价格表同步）
   *
   * 6. 校验价格数据完整性
   *
   * 7. 校验破价（白金价是否低于最低价）
   *
   * 8. 校验同款不同价（生成警告信息）
   *
   * 9. 为每个商品计算活动信息：
   *    - 活动价格（基于白金价和活动等级）
   *    - 活动利润
   *    - 活动利润率
   *    - 警告信息汇总
   *    - 白金限量特殊字段
   *
   * 10. 白金限量按SPU去重
   *
   * 11. 输出到新工作簿
   *
   * 12. 添加"活动提报"工作表
   *
   * 13. 写入数据（包含所有校验结果）
   */
  signUpActivity() {
    // 1. 验证活动等级
    const activityLevel = UserForm1.ComboBox3.Value;
    if (!activityLevel) {
      throw new Error("请先选择活动价格等级");
    }

    // 2. 检查数据时效性
    this._checkDataFreshness();

    // 3. 获取筛选条件
    let query = {};

    // 如果UI没有筛选条件，默认提报上架商品
    if (Object.keys(this._getUIQuery()).length === 0) {
      query.itemStatus = ["商品上线", "部分上线"];
    } else {
      query = this._getUIQuery();
    }

    // 4. 获取待提报商品
    let products = this._repository.findProducts();
    products = this._applyQuery(products, query);

    if (products.length === 0) {
      throw new Error("没有符合条件的商品可以提报");
    }

    // 5. 刷新商品价格数据
    this._refreshProductPrices();

    // 6. 校验价格数据完整性
    this._validatePriceData(products);

    // 7. 校验破价
    this._validatePriceBroken(products);

    // 8. 校验同款不同价
    const styleWarnings = this._validateSameStyleDifferentPrice(products);

    // 9. 为每个商品计算活动信息
    const activityProducts = [];

    products.forEach((product) => {
      // 跳过无货号商品
      if (!product.itemNumber) return;

      const warnings = [];

      // 同款不同价警告
      if (styleWarnings[product.styleNumber]) {
        warnings.push("同款不同价");
      }

      // 白金价小于到手价
      if (product.silverPrice < product.finalPrice) {
        warnings.push("白金价小于到手价，请核实价格等级");
      }

      // 验证利润
      const profitValidation = this._profitCalculator.validateProfit(
        product.brandSN,
        product.costPrice,
        product.silverPrice, // 使用白金价验证利润
        product.marketingPositioning,
        product.salesAge,
        product.userOperations1,
        product.userOperations2,
        product.rejectAndReturnRateOfLast7Days,
      );

      if (!profitValidation.valid) {
        warnings.push(...profitValidation.warnings);
      }

      // 清仓款特殊校验
      if (
        product.marketingPositioning === "清仓款" &&
        product.generalGoodsTotalInventory > 1
      ) {
        warnings.push("清仓款请解绑组合装");
      }

      // 计算活动价格
      const activityPrice = this._profitCalculator.calculateActivityPrice(
        product.silverPrice,
        activityLevel,
      );

      // 计算活动利润
      const activityProfit = this._profitCalculator.calculateProfit(
        product.brandSN,
        product.costPrice,
        activityPrice,
        product.userOperations1,
        product.userOperations2,
        product.rejectAndReturnRateOfLast7Days,
      );

      const activityProfitRate =
        activityProfit && product.costPrice
          ? Number((activityProfit / product.costPrice).toFixed(5))
          : undefined;

      // 构建活动商品对象
      activityProducts.push({
        ...product,
        activityLevel,
        activityPrice,
        activityProfit,
        activityProfitRate,
        warnMessage: warnings.join("/"),
        // 白金限量特殊字段
        isLimited: activityLevel === "白金限量" ? "是" : undefined,
        limitedCount: activityLevel === "白金限量" ? 500 : undefined,
        limitedCountForUser: activityLevel === "白金限量" ? 10 : undefined,
        canPromote: activityLevel === "白金限量" ? "是" : undefined,
      });
    });

    // 10. 白金限量按SPU去重
    let outputProducts = activityProducts;

    if (activityLevel === "白金限量") {
      const spuMap = new Map();
      activityProducts.forEach((product) => {
        if (!spuMap.has(product.P_SPU)) {
          spuMap.set(product.P_SPU, product);
        }
      });
      outputProducts = Array.from(spuMap.values());
    }

    // 11. 输出到新工作簿
    const sourceWb = this._excelDAO.getWorkbook();
    sourceWb.Sheets(this._config.get("Product").worksheet).Copy();
    const newWb = ActiveWorkbook;

    // 添加活动提报表
    newWb.Worksheets.Add();
    const sheet = ActiveSheet;
    sheet.Name = "活动提报";

    // 定义输出字段
    const outputFields = [
      "itemNumber",
      "styleNumber",
      "color",
      "P_SPU",
      "MID",
      "activityLevel",
      "activityPrice",
      "activityProfit",
      "activityProfitRate",
      "salesAge",
      "marketingPositioning",
      "isOutOfStock",
      "finishedGoodsTotalInventory",
      "generalGoodsTotalInventory",
      "sellableInventory",
      "sellableDays",
      "clickThroughRateOfLast7Days",
      "addToCartRateOfLast7Days",
      "purchaseRateOfLast7Days",
      "warnMessage",
    ];

    if (activityLevel === "白金限量") {
      outputFields.push(
        "isLimited",
        "limitedCount",
        "limitedCountForUser",
        "canPromote",
      );
    }

    // 构建输出数据
    const outputData = [];

    // 标题行
    const headers = outputFields.map((field) => {
      const fieldConfig = this._config.get("Product").fields[field];
      return fieldConfig?.title || field;
    });
    outputData.push(headers);

    // 数据行
    outputProducts.forEach((product) => {
      const row = outputFields.map((field) => {
        let value = product[field];

        if (typeof value === "number") {
          if (field.includes("Rate") || field.includes("率")) {
            return value;
          }
          return Number(value.toFixed(2));
        }

        return value !== undefined && value !== null ? String(value) : "";
      });

      outputData.push(row);
    });

    // 写入Excel
    sheet.Cells.ClearContents();
    const range = sheet
      .Range("A1")
      .Resize(outputData.length, outputData[0].length);
    range.Value2 = outputData;

    return newWb;
  }

  /**
   * 从UI获取查询条件
   * @private
   * @returns {Object} 查询条件对象
   * @description
   * 从用户表单中读取筛选条件：
   * - CheckBox17：商品上线
   * - CheckBox18：部分上线
   */
  _getUIQuery() {
    const query = {};

    if (UserForm1.CheckBox17.Value) query.itemStatus = ["商品上线"];
    if (UserForm1.CheckBox18.Value) {
      query.itemStatus = query.itemStatus
        ? [...query.itemStatus, "部分上线"]
        : ["部分上线"];
    }

    return query;
  }

  /**
   * 应用查询条件过滤商品
   * @private
   * @param {Object[]} products - 商品数组
   * @param {Object} query - 查询条件
   * @returns {Object[]} 过滤后的商品数组
   */
  _applyQuery(products, query) {
    return products.filter((product) => {
      return Object.entries(query).every(([key, condition]) => {
        if (Array.isArray(condition)) {
          return condition.includes(product[key]);
        }
        return product[key] === condition;
      });
    });
  }

  /**
   * 刷新商品价格数据
   * @private
   * @returns {void}
   * @description
   * 从ProductPrice实体读取最新价格数据，更新到Product实体：
   * - designNumber（款号）
   * - picture（图片）
   * - costPrice（成本价）
   * - lowestPrice（最低价）
   * - silverPrice（白金价）
   * - userOperations1/2（中台操作字段）
   */
  _refreshProductPrices() {
    const priceData = this._repository.findAll("ProductPrice");
    const priceMap = {};

    priceData.forEach((price) => {
      priceMap[price.itemNumber] = price;
    });

    const products = this._repository.findProducts();

    products.forEach((product) => {
      const price = priceMap[product.itemNumber];
      if (price) {
        product.designNumber = price.designNumber;
        product.picture = price.picture;
        product.costPrice = price.costPrice;
        product.lowestPrice = price.lowestPrice;
        product.silverPrice = price.silverPrice;
        product.userOperations1 = price.userOperations1 || 0;
        product.userOperations2 = price.userOperations2 || 0;
      }
    });

    this._repository.save("Product", products);
  }

  /**
   * 校验价格数据的完整性
   * @private
   * @param {Object[]} products - 商品数组
   * @throws {CustomError} 当存在价格数据不完整的商品时抛出
   * @description
   * 校验规则：
   * 1. 商品必须在价格表中存在
   * 2. 成本价、最低价、白金价必须填写完整
   *
   * 错误信息包含：
   * - itemNumber：货号
   * - errReason：异常原因
   */
  _validatePriceData(products) {
    const abnormal = [];

    products.forEach((product) => {
      if (!product.itemNumber) return;

      const price = this._repository.findPriceByItemNumber(product.itemNumber);

      if (!price) {
        abnormal.push({
          itemNumber: product.itemNumber,
          errReason: "商品价格表中未找到该货号",
        });
      } else if (!price.costPrice || !price.lowestPrice || !price.silverPrice) {
        abnormal.push({
          itemNumber: product.itemNumber,
          errReason: "成本价/最低价/白金价未填写完整",
        });
      }
    });

    if (abnormal.length > 0) {
      throw new CustomError(
        "商品价格数据不完整",
        { itemNumber: "货号", errReason: "异常原因" },
        abnormal,
      );
    }
  }

  /**
   * 校验破价情况
   * @private
   * @param {Object[]} products - 商品数组
   * @throws {CustomError} 当存在破价商品时抛出
   * @description
   * 破价定义：白金价低于历史最低价
   * 破价商品不允许参与活动报名。
   */
  _validatePriceBroken(products) {
    const broken = [];

    products.forEach((product) => {
      if (!product.itemNumber) return;

      if (product.silverPrice < product.lowestPrice) {
        broken.push({
          itemNumber: product.itemNumber,
          errReason: "白金价低于最低价，已破价",
        });
      }
    });

    if (broken.length > 0) {
      throw new CustomError(
        "存在破价商品",
        { itemNumber: "货号", errReason: "异常原因" },
        broken,
      );
    }
  }

  /**
   * 校验同款不同价
   * @private
   * @param {Object[]} products - 商品数组
   * @returns {Object} 警告信息映射，键为款号，值为true表示该款存在不同价
   * @description
   * 校验逻辑：
   * 1. 按款号分组
   * 2. 统计每个款号下所有颜色尺码的白金价
   * 3. 如果同一款号存在多个不同的白金价，标记为警告
   *
   * 同款不同价可能会导致：
   * - 活动价格不一致
   * - 用户体验差
   * - 运营管理困难
   */
  _validateSameStyleDifferentPrice(products) {
    const styleMap = new Map();
    const warnings = {};

    products.forEach((product) => {
      if (!product.styleNumber) return;

      if (!styleMap.has(product.styleNumber)) {
        styleMap.set(product.styleNumber, {
          count: 0,
          totalSilverPrice: 0,
          prices: new Set(),
        });
      }

      const style = styleMap.get(product.styleNumber);
      style.count++;
      style.totalSilverPrice += product.silverPrice;
      style.prices.add(product.silverPrice);
    });

    styleMap.forEach((style, styleNumber) => {
      if (style.prices.size > 1) {
        warnings[styleNumber] = true;
      }
    });

    return warnings;
  }
}
