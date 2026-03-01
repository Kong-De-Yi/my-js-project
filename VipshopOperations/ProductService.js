/**
 * 产品服务 - 负责产品的数据同步、更新和计算
 *
 * @class ProductService
 * @description 作为产品数据的核心业务逻辑层，提供以下功能：
 * - 从常态商品同步基础信息（款号、颜色、状态等）
 * - 从价格表同步价格信息（成本价、最低价、白金价等）
 * - 从库存表计算成品库存和通货库存
 * - 从销售表同步销售数据和拒退率
 * - 数据过期检查（12小时内未更新视为过期）
 * - 一键更新所有产品数据
 *
 * 该类采用单例模式，确保全局只有一个产品服务实例。
 *
 * @example
 * // 获取产品服务实例
 * const productService = ProductService.getInstance(repository);
 *
 * // 一键更新所有产品数据
 * const results = productService.updateAll();
 *
 * // 生成更新报告
 * const report = productService.generateUpdateReport(results);
 * MsgBox(report);
 */
class ProductService {
  /** @type {ProductService} 单例实例 */
  static _instance = null;

  /**
   * 创建产品服务实例
   * @param {Repository} [repository] - 数据仓库实例，若不提供则自动获取
   */
  constructor(repository) {
    if (ProductService._instance) {
      return ProductService._instance;
    }

    this._repository = repository || Repository.getInstance();
    this._config = DataConfig.getInstance();
    this._importableEntities = this._config.getImportableEntities();
    this._profitCalculator = ProfitCalculator.getInstance();

    this._validationEngine = ValidationEngine.getInstance();

    ProductService._instance = this;
  }

  /**
   * 获取产品服务的单例实例
   * @static
   * @param {Repository} [repository] - 数据仓库实例
   * @returns {ProductService} 产品服务实例
   */
  static getInstance(repository) {
    if (!ProductService._instance) {
      ProductService._instance = new ProductService(repository);
    }
    return ProductService._instance;
  }

  /**
   * 计算常态商品中的断码尺码
   * @private
   * @param {Object[]} regulars - 常态商品数组
   * @returns {string|undefined} 断码尺码，用"/"连接（如"S/M/L"），无断码时返回undefined
   * @description
   * 断码判定条件：
   * - 尺码状态为"尺码上线"
   * - 可售库存为0
   * 结果按尺码数字大小排序（如S、M、L按字母顺序或数字大小）
   */
  _getOutOfStockSizes(regulars) {
    const outOfStockSizes = regulars
      .filter(
        (rp) => rp.sizeStatus === "尺码上线" && rp.sellableInventory === 0,
      )
      .map((rp) => rp.size)
      .filter(Boolean)
      .sort((a, b) => +a - +b);

    return outOfStockSizes.length > 0 ? outOfStockSizes.join("/") : undefined;
  }

  /**
   * 重置产品的常态商品相关字段
   * @private
   * @param {Object} product - 产品对象
   * @description
   * 重置的字段包括：
   * - 基础信息：款号、颜色、三级品类、状态
   * - 价格信息：吊牌价、唯品价、最终价
   * - 库存信息：可售库存、可售天数、断码尺码
   * - 标识信息：MID、P_SPU、品牌SN
   */
  _resetRegularFields(product) {
    product.styleNumber = undefined;
    product.color = undefined;
    product.thirdLevelCategory = undefined;
    product.itemStatus = undefined;
    product.tagPrice = undefined;
    product.vipshopPrice = undefined;
    product.finalPrice = undefined;
    product.sellableInventory = 0;
    product.sellableDays = 0;
    product.isOutOfStock = undefined;
    product.MID = undefined;
    product.P_SPU = undefined;
    product.brandSN = undefined;
  }

  /**
   * 从常态商品更新指定产品的基础信息
   * @private
   * @param {Object} product - 要更新的产品对象
   * @returns {Object} 更新后的产品对象
   * @description
   * 更新逻辑：
   * 1. 查找该货号的所有常态商品
   * 2. 如果没找到，重置所有常态商品相关字段
   * 3. 如果找到，使用第一个尺码的信息更新产品基础信息
   * 4. 累加所有尺码的可售库存
   * 5. 计算断码尺码
   * 6. 如果产品已上线，清空下线原因
   */
  _updateProductFromRegulars(product) {
    const regulars = this._repository.findRegularProductsByItemNumber(
      product.itemNumber,
    );
    if (regulars.length === 0) {
      // 没有找到常态商品则重置
      this._resetRegularFields(product);
      return product;
    }

    const first = regulars[0];

    // 更新基础信息
    product.styleNumber = first.styleNumber;
    product.color = first.color;
    product.thirdLevelCategory = first.thirdLevelCategory;
    product.itemStatus = first.itemStatus;
    product.tagPrice = first.tagPrice;
    product.vipshopPrice = first.vipshopPrice;
    product.finalPrice = first.finalPrice;
    product.sellableDays = first.sellableDays;
    product.isOutOfStock = this._getOutOfStockSizes(regulars);
    product.MID = first.MID;
    product.P_SPU = first.P_SPU;
    product.brandSN = first.brandSN;

    // 计算可售库存
    product.sellableInventory = regulars.reduce(
      (sum, rp) => sum + rp.sellableInventory,
      0,
    );

    // 清空下线原因（如果已上线）
    if (product.itemStatus !== "商品下线") {
      product.offlineReason = undefined;
    }

    return product;
  }

  /**
   * 根据货号从常态商品创建新的产品对象
   * @private
   * @param {string} itemNumber - 货号
   * @returns {Object|null} 新产品对象，无常态商品时返回null
   * @description
   * 创建逻辑：
   * - 使用第一个尺码的信息作为产品基础信息
   * - 默认设置营销定位为"利润款"
   * - 累加所有尺码的可售库存
   * - 计算断码尺码
   */
  _createProductFromRegulars(itemNumber) {
    const regulars =
      this._repository.findRegularProductsByItemNumber(itemNumber);

    if (regulars.length === 0) return null;

    const first = regulars[0];

    return {
      itemNumber,
      styleNumber: first.styleNumber,
      color: first.color,
      itemStatus: first.itemStatus,
      marketingPositioning: "利润款",
      finalPrice: first.finalPrice,
      sellableInventory: regulars.reduce(
        (sum, rp) => sum + rp.sellableInventory,
        0,
      ),
      sellableDays: first.sellableDays,
      isOutOfStock: this._getOutOfStockSizes(regulars),
      brandSN: first.brandSN,
      MID: first.MID,
      P_SPU: first.P_SPU,
      thirdLevelCategory: first.thirdLevelCategory,
      tagPrice: first.tagPrice,
      vipshopPrice: first.vipshopPrice,
    };
  }

  /**
   * 重置产品的价格相关字段
   * @private
   * @param {Object} product - 产品对象
   * @description
   * 重置的字段：
   * - costPrice（成本价）
   * - lowestPrice（最低价）
   * - silverPrice（白金价）
   * - userOperations1/2（用户操作，重置为0）
   */
  _resetPriceFields(product) {
    product.costPrice = undefined;
    product.lowestPrice = undefined;
    product.silverPrice = undefined;
    product.userOperations1 = 0;
    product.userOperations2 = 0;
  }

  /**
   * 从价格表更新指定产品的价格信息
   * @private
   * @param {Object} product - 产品对象
   * @returns {boolean} 是否有字段被更新
   * @description
   * 更新逻辑：
   * 1. 查找货号对应的价格记录
   * 2. 无记录且产品有价格信息时，重置价格字段
   * 3. 有记录时更新成本价、最低价、白金价
   * 4. 计算活动价（直通车、黄金促、TOP3、白金限量）
   *    - 活动价基于白金价(silverPrice)计算
   *    - 如果活动价字段已有值则跳过
   */
  _applyPriceToProduct(product) {
    let changed = false;
    const price = this._repository.findPriceByItemNumber(product.itemNumber);
    if (!price) {
      if (
        product.costPrice &&
        product.lowestPrice &&
        product.silverPrice &&
        product.userOperations1 &&
        product.userOperations2
      ) {
        // 没有找到价格信息则重置
        this._resetPriceFields(product);
        changed = true;
        return changed;
      } else {
        return changed;
      }
    }

    if (product.costPrice !== price.costPrice) {
      product.costPrice = price.costPrice;
      changed = true;
    }
    if (product.lowestPrice !== price.lowestPrice) {
      product.lowestPrice = price.lowestPrice;
      changed = true;
    }
    if (product.silverPrice !== price.silverPrice) {
      product.silverPrice = price.silverPrice;
      changed = true;
    }
    if (product.userOperations1 !== price.userOperations1) {
      product.userOperations1 = price.userOperations1;
      changed = true;
    }
    if (product.userOperations2 !== price.userOperations2) {
      product.userOperations2 = price.userOperations2;
      changed = true;
    }

    // 活动价
    const activityPrices = {
      directTrainPrice: "直通车",
      goldPrice: "黄金促",
      top3: "TOP3",
      silverLimit: "白金限量",
    };

    for (const [ap, level] of Object.entries(activityPrices)) {
      if (product[ap]) continue;
      product[ap] = this._profitCalculator.calculateActivityPrice(
        product.silverPrice,
        level,
      );
      changed = true;
    }

    return changed;
  }

  /**
   * 重置产品的所有库存相关字段为0
   * @private
   * @param {Object} product - 产品对象
   * @description
   * 重置的字段分为两大类：
   * - 成品库存：finishedGoods*（7种库存类型）
   * - 通货库存：generalGoods*（7种库存类型）
   * 每种类型包括：主库存、在途库存、加工库存、超卖库存、
   * 备货库存、退货库存、采购库存
   */
  _resetInventoryFields(product) {
    const fields = [
      "finishedGoodsMainInventory",
      "finishedGoodsIncomingInventory",
      "finishedGoodsFinishingInventory",
      "finishedGoodsOversoldInventory",
      "finishedGoodsPrepareInventory",
      "finishedGoodsReturnInventory",
      "finishedGoodsPurchaseInventory",
      "generalGoodsMainInventory",
      "generalGoodsIncomingInventory",
      "generalGoodsFinishingInventory",
      "generalGoodsOversoldInventory",
      "generalGoodsPrepareInventory",
      "generalGoodsReturnInventory",
      "generalGoodsPurchaseInventory",
    ];

    fields.forEach((f) => (product[f] = 0));
  }

  /**
   * 计算产品的成品库存
   * @private
   * @param {Object} product - 产品对象
   * @returns {number} 成品库存总和
   * @description
   * 计算逻辑：
   * 1. 查找该货号的所有常态商品（每个尺码一个）
   * 2. 通过每个常态商品的 productCode 查找库存记录
   * 3. 累加所有7种库存类型到对应的 finishedGoods* 字段
   * 4. 返回所有库存类型的总和
   */
  _calculateFinishedGoods(product) {
    // 查找该货号的所有常态商品
    const regulars = this._repository.findRegularProductsByItemNumber(
      product.itemNumber,
    );

    if (regulars.length === 0) return 0;

    // 累加成品条码的库存
    regulars.forEach((r) => {
      const inv = this._repository.findInventory(r.productCode);
      if (!inv) return;

      product.finishedGoodsMainInventory += inv.mainInventory;
      product.finishedGoodsIncomingInventory += inv.incomingInventory;
      product.finishedGoodsFinishingInventory += inv.finishingInventory;
      product.finishedGoodsOversoldInventory += inv.oversoldInventory;
      product.finishedGoodsPrepareInventory += inv.prepareInventory;
      product.finishedGoodsReturnInventory += inv.returnInventory;
      product.finishedGoodsPurchaseInventory += inv.purchaseInventory;
    });

    return (
      product.finishedGoodsMainInventory +
      product.finishedGoodsIncomingInventory +
      product.finishedGoodsFinishingInventory +
      product.finishedGoodsOversoldInventory +
      product.finishedGoodsPrepareInventory +
      product.finishedGoodsReturnInventory +
      product.finishedGoodsPurchaseInventory
    );
  }

  /**
   * 计算产品的通货库存
   * @private
   * @param {Object} product - 产品对象
   * @returns {number} 通货库存总和
   * @description
   * 计算逻辑：
   * 1. 查找该货号的所有常态商品
   * 2. 查找每个常态商品的组合装
   * 3. 排除以"YH"（赠品）或"FL"（辅料）开头的子商品
   * 4. 根据子商品数量和库存计算分摊库存（库存 / 数量）
   * 5. 累加所有7种库存类型到对应的 generalGoods* 字段
   * 6. 返回所有库存类型的总和
   *
   * @example
   * // 如果一个组合装包含2个相同的子商品
   * // 子商品库存为100，则分摊到产品的通货库存为50
   */
  _calculateGeneralGoods(product) {
    // 查找该货号的所有常态商品
    const regulars = this._repository.findRegularProductsByItemNumber(
      product.itemNumber,
    );

    if (regulars.length === 0) return 0;

    regulars.forEach((r) => {
      const combos = this._repository.findComboProducts(r.productCode);
      if (combos.length === 0) return;
      combos.forEach((combo) => {
        const subPC = combo.subProductCode;

        if (subPC.startsWith("YH") || subPC.startsWith("FL")) return;

        const quantity = combo.subProductQuantity;
        const subInv = this._repository.findInventory(subPC);
        if (!subInv) return;

        product.generalGoodsMainInventory += subInv.mainInventory / quantity;
        product.generalGoodsIncomingInventory +=
          subInv.incomingInventory / quantity;
        product.generalGoodsFinishingInventory +=
          subInv.finishingInventory / quantity;
        product.generalGoodsOversoldInventory +=
          subInv.oversoldInventory / quantity;
        product.generalGoodsPrepareInventory +=
          subInv.prepareInventory / quantity;
        product.generalGoodsReturnInventory +=
          subInv.returnInventory / quantity;
        product.generalGoodsPurchaseInventory +=
          subInv.purchaseInventory / quantity;
      });
    });

    return (
      product.generalGoodsMainInventory +
      product.generalGoodsIncomingInventory +
      product.generalGoodsFinishingInventory +
      product.generalGoodsOversoldInventory +
      product.generalGoodsPrepareInventory +
      product.generalGoodsReturnInventory +
      product.generalGoodsPurchaseInventory
    );
  }

  /**
   * 从销售表更新指定产品的销售信息
   * @private
   * @param {Object} product - 产品对象
   * @param {number} lastNDays - 最近N天
   * @returns {boolean} 是否有字段被更新
   * @description
   * 更新的字段：
   * 1. rejectAndReturnRate（近N天拒退率）
   *    - 计算方式：拒退数量 / 销售数量
   *    - 无销售数据时设为undefined
   * 2. firstListingTime（首次上架时间）
   *    - 如果产品尚无首次上架时间，使用第一条销售记录的首次上架时间
   */
  _applySalesToProduct(product, lastNDays) {
    let changed = false;

    const oldRARR = product.rejectAndReturnRate;

    // 1.获取最近N天的销售数据
    const salesLastNDays = this._repository.findSalesLastNDays(
      product.itemNumber,
      lastNDays,
    );
    if (salesLastNDays.length === 0) {
      if (oldRARR) {
        product.rejectAndReturnRate = undefined;
        changed = true;
        return changed;
      } else {
        return changed;
      }
    }

    // 2.计算近N天的拒退率
    const rr = salesLastNDays.reduce(
      (rd, sd) => {
        rd.salesQuantity += sd.salesQuantity;
        rd.rejectAndReturnCount += sd.rejectAndReturnCount;
        return rd;
      },
      { salesQuantity: 0, rejectAndReturnCount: 0 },
    );

    product.rejectAndReturnRate = rr.salesQuantity
      ? rr.rejectAndReturnCount / rr.salesQuantity
      : undefined;

    if (product.rejectAndReturnRate !== oldRARR) changed = true;

    // 更新首次上架时间
    const flt = salesLastNDays[0].firstListingTime;
    if (!product.firstListingTime && flt) {
      product.firstListingTime = flt;

      changed = true;
    }

    return changed;
  }

  /**
   * 检查业务实体数据是否过期
   * @private
   * @param {string} entityName - 实体名称
   * @returns {boolean} true=已过期，false=未过期
   * @description
   * 过期判定条件（任一满足即过期）：
   * 1. 实体不在可导入列表中 => 过期
   * 2. 实体配置中没有配置 importDate 字段 => 过期
   * 3. 系统记录中没有该实体的导入日期 => 过期
   * 4. 导入时间距离现在超过12小时 => 过期
   */
  _checkDataExpired(entityName) {
    if (!this._importableEntities.includes(entityName)) return true;

    const entityConfig = this._config.get(entityName);
    const importDate = entityConfig?.importDate;
    if (!importDate) return true;

    const systemRecord = this._repository.getSystemRecord();

    // 没有更新日期默认过期
    if (!systemRecord[importDate]) return true;

    const importDateTS = Date.parse(systemRecord[importDate]);
    return new Date() - importDateTS > 12 * 60 * 60 * 1000;
  }

  /**
   * 刷新产品缓存
   * @returns {void}
   * @description 强制从Excel重新读取产品数据，更新缓存和索引
   */
  refreshProduct() {
    this._repository.refresh("Product");
  }

  /**
   * 从常态商品更新产品数据
   * @returns {Object} 更新结果统计
   * @returns {Object} return.regular - 常态商品更新统计
   * @returns {number} return.regular.totalProducts - 当前总产品数
   * @returns {number} return.regular.updatedProducts - 更新的产品数
   * @returns {number} return.regular.newProducts - 新增的货号数
   * @throws {Error} 常态商品数据过期（超过12小时未导入）时抛出
   *
   * @description
   * 更新流程：
   * 1. 检查常态商品数据是否过期
   * 2. 获取所有现有产品
   * 3. 更新现有产品的信息（基于货号匹配）
   * 4. 扫描所有常态商品中的货号，发现新货号则创建新产品
   * 5. 合并新旧数据并保存
   * 6. 更新系统记录中的常态商品更新日期
   */
  updateFromRegularProducts() {
    const result = {
      totalProducts: 0,
      updatedProducts: 0,
      newProducts: 0,
    };

    // 1.检查常态商品是否为最新
    if (this._checkDataExpired("RegularProduct")) {
      throw new Error(`【常态商品】今日尚未导入，请先导入！`);
    }

    // 2.获取所有产品
    const products = this._repository.findProducts();

    result.totalProducts = products.length;

    // 3. 更新现有常态商品
    const updatedProducts = products.map((product) => {
      return this._updateProductFromRegulars(product);
    });
    result.updatedProducts = updatedProducts.length;

    // 4. 添加新货号
    const existingItemNumbers = new Set(products.map((p) => p.itemNumber));
    const AllregularProducts = this._repository.findRegularProducts();
    const AllregularItemNumbers = new Set(
      AllregularProducts.map((rp) => rp.itemNumber),
    );

    const newProducts = [];
    for (const itemNumber of AllregularItemNumbers) {
      if (!existingItemNumbers.has(itemNumber)) {
        const newProduct = this._createProductFromRegulars(itemNumber);
        if (newProduct) {
          newProducts.push(newProduct);
          result.newProducts++;
        }
      }
    }

    // 5.合并数据
    const allProducts = [...updatedProducts, ...newProducts];

    this._repository.save("Product", allProducts);
    this._repository.updateSystemRecord("RegularProduct", "updateDate");

    return { regular: result };
  }

  /**
   * 从价格表更新产品价格
   * @returns {Object} 更新结果统计
   * @returns {Object} return.price - 价格更新统计
   * @returns {number} return.price.updated - 价格有变动的产品数
   * @returns {number} return.price.skipped - 价格无变动的产品数
   * @throws {Error} 商品价格数据过期时抛出
   */
  updateFromPriceData() {
    const result = {
      updated: 0,
      skipped: 0,
    };

    // 1.检查商品价格是否为最新
    if (this._checkDataExpired("ProductPrice")) {
      throw new Error(`【商品价格】今日尚未导入，请先导入！`);
    }

    // 2.获取所有产品
    const products = this._repository.findProducts();

    // 3.更新产品价格
    products.forEach((product) => {
      const changed = this._applyPriceToProduct(product);
      if (changed) {
        result.updated++;
      } else {
        result.skipped++;
      }
    });

    this._repository.save("Product", products);
    this._repository.updateSystemRecord("ProductPrice", "updateDate");

    return { price: result };
  }

  /**
   * 从库存数据更新产品库存
   * @returns {Object} 更新结果统计
   * @returns {Object} return.inventory - 库存更新统计
   * @returns {number} return.inventory.updated - 库存变动的产品数
   * @returns {number} return.inventory.zeroInventory - 零库存产品数
   * @throws {Error} 组合商品或库存数据过期时抛出
   */
  updateFromInventory() {
    const result = {
      updated: 0,
      zeroInventory: 0,
    };

    // 1.检查组合装和商品库存是否更新
    if (this._checkDataExpired("ComboProduct")) {
      throw new Error("【商品库存】需要最新的组合商品，请先导入！");
    }
    if (this._checkDataExpired("Inventory")) {
      throw new Error("【商品库存】今日尚未导入，请先导入！");
    }

    // 2.获取所有产品
    const products = this._repository.findProducts();

    // 3.计算库存
    products.forEach((product) => {
      const before = product.totalInventory;

      // 重置库存
      this._resetInventoryFields(product);
      // 计算成品库存
      const finishedGoodsTI = this._calculateFinishedGoods(product);
      // 计算通货库存
      const GeneralGoodsTI = this._calculateGeneralGoods(product);

      const after = finishedGoodsTI + GeneralGoodsTI;

      if (before !== after) result.updated++;
      if (after === 0) result.zeroInventory++;
    });

    this._repository.save("Product", products);
    this._repository.updateSystemRecord("Inventory", "updateDate");

    return { inventory: result };
  }

  /**
   * 从销售数据更新产品销售信息
   * @returns {Object} 更新结果统计
   * @returns {Object} return.sales - 销售更新统计
   * @returns {number} return.sales.updated - 销售信息变动的产品数
   * @returns {number} return.sales.skipped - 销售信息无变动的产品数
   * @throws {Error} 商品销售数据过期时抛出
   */
  updateFromSalesData() {
    const result = {
      updated: 0,
      skipped: 0,
    };

    // 1.检查商品销售是否为最新
    if (this._checkDataExpired("ProductSales")) {
      throw new Error(`【商品销售】今日尚未导入，请先导入！`);
    }

    // 2.获取所有产品
    const products = this._repository.findProducts();

    // 3.更新销售数据
    products.forEach((product) => {
      const changed = this._applySalesToProduct(product, 30);
      if (changed) {
        result.updated++;
      } else {
        result.skipped++;
      }
    });

    this._repository.save("Product", products);
    this._repository.updateSystemRecord("ProductSales", "updateDate");

    return { sales: result };
  }

  /**
   * 一键更新所有产品数据
   * @returns {Object} 所有更新操作的汇总结果
   * @returns {Array} return.errors - 各步骤的错误信息数组
   * @returns {Object} [return.regular] - 常态商品更新结果
   * @returns {Object} [return.price] - 价格更新结果
   * @returns {Object} [return.inventory] - 库存更新结果
   * @returns {Object} [return.sales] - 销售更新结果
   *
   * @description
   * 按顺序执行所有更新操作，使用try-catch保证某一步失败不影响后续步骤：
   * 1. 更新常态商品（基础信息）
   * 2. 更新商品价格
   * 3. 更新商品库存
   * 4. 更新商品销售
   *
   * 所有步骤执行完毕后返回汇总结果，包含各步骤的统计和错误信息。
   *
   * @example
   * const results = productService.updateAll();
   * if (results.errors.length > 0) {
   *   MsgBox("部分更新失败：\n" + results.errors.join("\n"));
   * } else {
   *   MsgBox("所有数据更新成功！");
   * }
   */
  updateAll() {
    const results = { errors: [] };
    let result = {};

    // 1.更新常态商品
    try {
      result = this.updateFromRegularProducts();
      Object.assign(results, result);
    } catch (e) {
      results.errors.push(e.message);
    }

    // 2.更新商品价格
    try {
      result = this.updateFromPriceData();
      Object.assign(results, result);
    } catch (e) {
      results.errors.push(e.message);
    }

    // 3.更新商品库存
    try {
      result = this.updateFromInventory();
      Object.assign(results, result);
    } catch (e) {
      results.errors.push(e.message);
    }

    // 4.更新商品销售
    try {
      result = this.updateFromSalesData();
      Object.assign(results, result);
    } catch (e) {
      results.errors.push(e.message);
    }

    return results;
  }

  /**
   * 生成格式化的更新报告
   * @param {Object} results - updateAll 方法返回的结果对象
   * @returns {string} 格式化的报告文本，可直接用于消息框显示
   *
   * @description
   * 生成的报告包含以下部分（按实际存在的步骤）：
   * - 常态商品：总产品数、新增货号、更新产品
   * - 商品价格：更新数、跳过数
   * - 商品库存：库存变动数、零库存数
   * - 商品销售：更新数、跳过数
   * - 错误信息：如果有错误，列出所有错误信息
   *
   * @example
   * const results = productService.updateAll();
   * const report = productService.generateUpdateReport(results);
   * MsgBox(report); // 显示完整的更新报告
   */
  generateUpdateReport(results) {
    let report = "========== 数据更新报告 ==========\n\n";

    if (results.regular) {
      report += `【常态商品】\n`;
      report += `  总产品数: ${results.regular.totalProducts}\n`;
      report += `  新增货号: ${results.regular.newProducts}\n`;
      report += `  更新产品: ${results.regular.updatedProducts}\n`;
    }

    if (results.price) {
      report += `\n【商品价格】\n`;
      report += `  更新: ${results.price.updated}\n`;
      report += `  跳过: ${results.price.skipped}\n`;
    }

    if (results.inventory) {
      report += `\n【商品库存】\n`;
      report += `  库存变动: ${results.inventory.updated}\n`;
      report += `  零库存: ${results.inventory.zeroInventory}\n`;
    }

    if (results.sales) {
      report += `\n【商品销售】\n`;
      report += `  更新: ${results.sales.updated}\n`;
      report += `  跳过: ${results.sales.skipped}\n`;
    }

    if (results?.errors?.length > 0) {
      report += `\n【错误信息】\n`;
      results.errors.forEach((err) => (report += `  ${err}\n`));
    }

    return report;
  }
}
