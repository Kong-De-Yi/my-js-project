// ============================================================================
// 活动提报服务
// 功能：平台活动报名、校验、导出
// 特点：独立业务逻辑，与数据访问解耦
// ============================================================================

class ActivityService {
  constructor(repository, profitCalculator, excelDAO) {
    this._repository = repository;
    this._profitCalculator = profitCalculator;
    this._excelDAO = excelDAO;
    this._config = DataConfig.getInstance();
  }

  // 检查数据时效性
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

  // 获取近7天日期范围
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

  // 判断是否是今天
  _isToday(timestamp) {
    const date = new Date(timestamp);
    const today = new Date();
    return (
      date.getDate() === today.getDate() &&
      date.getMonth() === today.getMonth() &&
      date.getFullYear() === today.getFullYear()
    );
  }

  // 生成活动提报表
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

  // 从UI获取查询条件
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

  // 应用查询
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

  // 刷新商品价格
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

  // 校验价格数据完整性
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

  // 校验破价
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

  // 校验同款不同价
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
