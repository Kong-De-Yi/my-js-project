// ============================================================================
// ProductService.js - 产品业务服务
// 功能：封装所有产品相关的业务逻辑
// ============================================================================

class ProductService {
  constructor(repository, profitCalculator) {
    this._repository = repository;
    this._profitCalculator = profitCalculator;
    this._config = DataConfig.getInstance();
  }

  // 从常态商品更新数据
  updateFromRegularProducts() {
    const result = {
      totalProducts: 0,
      newProducts: 0,
      updatedProducts: 0,
      errors: [],
    };

    try {
      // 1. 刷新数据
      this._repository.refresh("Product");
      const regularProducts = this._repository.findAll("RegularProduct");
      let products = this._repository.findProducts();

      result.totalProducts = products.length;

      // 2. 按货号分组常态商品
      const regularByItem = this._groupRegularProducts(regularProducts);

      // 3. 处理每个产品
      const updatedProducts = products.map((product) => {
        const regulars = regularByItem[product.itemNumber] || [];
        return this._updateProductFromRegulars(product, regulars);
      });

      // 4. 添加新货号
      const existingItemNumbers = new Set(products.map((p) => p.itemNumber));
      const newProducts = [];

      Object.entries(regularByItem).forEach(([itemNumber, regulars]) => {
        if (!existingItemNumbers.has(itemNumber)) {
          const newProduct = this._createProductFromRegulars(
            itemNumber,
            regulars,
          );
          newProducts.push(newProduct);
          result.newProducts++;
        }
      });

      // 5. 合并并保存
      const allProducts = [...updatedProducts, ...newProducts];
      this._repository.save("Product", allProducts);

      result.updatedProducts = updatedProducts.length;

      return result;
    } catch (e) {
      result.errors.push(e.message);
      throw e;
    }
  }

  /**
   * 从价格表更新产品价格
   */
  updateFromPriceData() {
    const result = {
      updated: 0,
      skipped: 0,
      errors: [],
    };

    try {
      this._repository.refresh("ProductPrice");
      const priceData = this._repository.findAll("ProductPrice");
      const products = this._repository.findProducts();

      // 建立价格映射
      const priceMap = {};
      priceData.forEach((p) => {
        if (p.itemNumber) {
          priceMap[p.itemNumber] = p;
        }
      });

      // 更新产品价格
      products.forEach((product) => {
        const price = priceMap[product.itemNumber];
        if (price) {
          const changed = this._applyPriceToProduct(product, price);
          if (changed) result.updated++;
        } else {
          result.skipped++;
        }
      });

      this._repository.save("Product", products);

      return result;
    } catch (e) {
      result.errors.push(e.message);
      throw e;
    }
  }

  /**
   * 从库存数据更新产品库存
   */
  updateFromInventory() {
    const result = {
      updated: 0,
      zeroInventory: 0,
      errors: [],
    };

    try {
      this._repository.refresh("Inventory");
      this._repository.refresh("ComboProduct");

      const products = this._repository.findProducts();
      const inventoryMap = this._buildInventoryMap();
      const comboMap = this._buildComboMap();

      products.forEach((product) => {
        const before = product.totalInventory;
        this._calculateProductInventory(product, inventoryMap, comboMap);
        const after = product.totalInventory;

        if (before !== after) result.updated++;
        if (after === 0) result.zeroInventory++;
      });

      this._repository.save("Product", products);
      this._repository.clear("Inventory");
      this._repository.clear("ComboProduct");

      return result;
    } catch (e) {
      result.errors.push(e.message);
      throw e;
    }
  }

  /**
   * 从销售数据更新产品销售信息
   */
  updateFromSalesData() {
    const result = {
      updated: 0,
      withSales: 0,
      errors: [],
    };

    try {
      this._repository.refresh("ProductSales");

      const products = this._repository.findProducts();
      const salesData = this._repository.findAll("ProductSales");

      // 按货号分组销售数据
      const salesByItem = this._groupSalesData(salesData);

      // 近7天日期范围
      const last7Days = this._getLast7DaysRange();

      products.forEach((product) => {
        const sales = salesByItem[product.itemNumber] || [];
        const changed = this._applySalesToProduct(product, sales, last7Days);

        if (changed) {
          result.updated++;
          if (product.salesQuantityOfLast7Days > 0) {
            result.withSales++;
          }
        }
      });

      this._repository.save("Product", products);

      return result;
    } catch (e) {
      result.errors.push(e.message);
      throw e;
    }
  }

  /**
   * 一键更新所有数据
   */
  async updateAll() {
    const results = {
      regular: null,
      price: null,
      inventory: null,
      sales: null,
      success: true,
      errors: [],
    };

    try {
      // 按顺序执行更新
      results.regular = this.updateFromRegularProducts();
      results.price = this.updateFromPriceData();
      results.inventory = this.updateFromInventory();
      results.sales = this.updateFromSalesData();

      // 更新系统记录
      this._updateSystemRecord();
    } catch (e) {
      results.success = false;
      results.errors.push(e.message);
    }

    return results;
  }

  // ========== 私有辅助方法 ==========

  _groupRegularProducts(regularProducts) {
    const groups = {};
    regularProducts.forEach((rp) => {
      if (!rp.itemNumber) return;
      if (!groups[rp.itemNumber]) {
        groups[rp.itemNumber] = [];
      }
      groups[rp.itemNumber].push(rp);
    });
    return groups;
  }

  _updateProductFromRegulars(product, regulars) {
    if (regulars.length === 0) return product;

    const first = regulars[0];

    // 更新基础信息
    product.thirdLevelCategory = first.thirdLevelCategory;
    product.P_SPU = first.P_SPU;
    product.MID = first.MID;
    product.styleNumber = first.styleNumber;
    product.color = first.color;
    product.itemStatus = first.itemStatus;
    product.vipshopPrice = first.vipshopPrice;
    product.finalPrice = first.finalPrice;
    product.sellableDays = first.sellableDays;

    // 清空下线原因（如果已上线）
    if (product.itemStatus !== "商品下线") {
      product.offlineReason = "";
    }

    // 计算可售库存
    product.sellableInventory = regulars.reduce(
      (sum, rp) => sum + (Number(rp.sellableInventory) || 0),
      0,
    );

    // 计算断码
    const outOfStockSizes = regulars
      .filter(
        (rp) =>
          rp.sizeStatus === "尺码上线" && (rp.sellableInventory || 0) === 0,
      )
      .map((rp) => rp.size)
      .filter(Boolean);

    product.isOutOfStock =
      outOfStockSizes.length > 0 ? outOfStockSizes.join("/") : "";
    return product;
  }

  _createProductFromRegulars(itemNumber, regulars) {
    const first = regulars[0];

    return {
      itemNumber,
      brandSN: first.brandSN,
      brandName: first.brand,
      marketingPositioning: "利润款",
      thirdLevelCategory: first.thirdLevelCategory,
      P_SPU: first.P_SPU,
      MID: first.MID,
      styleNumber: first.styleNumber,
      color: first.color,
      itemStatus: first.itemStatus,
      vipshopPrice: first.vipshopPrice,
      finalPrice: first.finalPrice,
      sellableDays: first.sellableDays,
      sellableInventory: regulars.reduce(
        (sum, rp) => sum + (Number(rp.sellableInventory) || 0),
        0,
      ),
      _rowNumber: Date.now(), // 临时行号
    };
  }

  _applyPriceToProduct(product, price) {
    let changed = false;

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
    if ((product.userOperations1 || 0) !== (price.userOperations1 || 0)) {
      product.userOperations1 = price.userOperations1 || 0;
      changed = true;
    }
    if ((product.userOperations2 || 0) !== (price.userOperations2 || 0)) {
      product.userOperations2 = price.userOperations2 || 0;
      changed = true;
    }

    return changed;
  }

  _buildInventoryMap() {
    const map = {};
    this._repository.findAll("Inventory").forEach((inv) => {
      if (inv.productCode) {
        map[inv.productCode] = inv;
      }
    });
    return map;
  }

  _buildComboMap() {
    const map = {};
    this._repository.findAll("ComboProduct").forEach((combo) => {
      if (!map[combo.productCode]) {
        map[combo.productCode] = [];
      }
      map[combo.productCode].push(combo);
    });
    return map;
  }

  _calculateProductInventory(product, inventoryMap, comboMap) {
    // 重置库存
    this._resetInventoryFields(product);

    // 查找该货号的所有常态商品
    const regulars = this._repository.findRegularProducts({
      itemNumber: product.itemNumber,
    });

    regulars.forEach((regular) => {
      this._calculateFinishedGoods(product, regular, inventoryMap);
      this._calculateGeneralGoods(product, regular, inventoryMap, comboMap);
    });

    // 合计库存
    product.totalInventory =
      product.finishedGoodsTotalInventory + product.generalGoodsTotalInventory;
  }

  _resetInventoryFields(product) {
    const fields = [
      "finishedGoodsMainInventory",
      "finishedGoodsIncomingInventory",
      "finishedGoodsFinishingInventory",
      "finishedGoodsOversoldInventory",
      "finishedGoodsPrepareInventory",
      "finishedGoodsReturnInventory",
      "finishedGoodsPurchaseInventory",
      "finishedGoodsTotalInventory",
      "generalGoodsMainInventory",
      "generalGoodsIncomingInventory",
      "generalGoodsFinishingInventory",
      "generalGoodsOversoldInventory",
      "generalGoodsPrepareInventory",
      "generalGoodsReturnInventory",
      "generalGoodsPurchaseInventory",
      "generalGoodsTotalInventory",
      "totalInventory",
    ];

    fields.forEach((f) => (product[f] = 0));
  }

  _calculateFinishedGoods(product, regular, inventoryMap) {
    const inv = inventoryMap[regular.productCode];
    if (!inv) return;

    product.finishedGoodsMainInventory += Number(inv.mainInventory || 0);
    product.finishedGoodsIncomingInventory += Number(
      inv.incomingInventory || 0,
    );
    product.finishedGoodsFinishingInventory += Number(
      inv.finishingInventory || 0,
    );
    product.finishedGoodsOversoldInventory += Number(
      inv.oversoldInventory || 0,
    );
    product.finishedGoodsPrepareInventory += Number(inv.prepareInventory || 0);
    product.finishedGoodsReturnInventory += Number(inv.returnInventory || 0);
    product.finishedGoodsPurchaseInventory += Number(
      inv.purchaseInventory || 0,
    );

    product.finishedGoodsTotalInventory =
      product.finishedGoodsMainInventory +
      product.finishedGoodsIncomingInventory +
      product.finishedGoodsFinishingInventory +
      product.finishedGoodsOversoldInventory +
      product.finishedGoodsPrepareInventory +
      product.finishedGoodsReturnInventory +
      product.finishedGoodsPurchaseInventory;
  }

  _calculateGeneralGoods(product, regular, inventoryMap, comboMap) {
    const combos = comboMap[regular.productCode] || [];

    combos.forEach((combo) => {
      const subInv = inventoryMap[combo.subProductCode];
      const quantity = Number(combo.subProductQuantity) || 1;

      if (subInv) {
        product.generalGoodsMainInventory +=
          Number(subInv.mainInventory || 0) * quantity;
        product.generalGoodsIncomingInventory +=
          Number(subInv.incomingInventory || 0) * quantity;
        product.generalGoodsFinishingInventory +=
          Number(subInv.finishingInventory || 0) * quantity;
        product.generalGoodsOversoldInventory +=
          Number(subInv.oversoldInventory || 0) * quantity;
        product.generalGoodsPrepareInventory +=
          Number(subInv.prepareInventory || 0) * quantity;
        product.generalGoodsReturnInventory +=
          Number(subInv.returnInventory || 0) * quantity;
        product.generalGoodsPurchaseInventory +=
          Number(subInv.purchaseInventory || 0) * quantity;
      }
    });

    product.generalGoodsTotalInventory =
      product.generalGoodsMainInventory +
      product.generalGoodsIncomingInventory +
      product.generalGoodsFinishingInventory +
      product.generalGoodsOversoldInventory +
      product.generalGoodsPrepareInventory +
      product.generalGoodsReturnInventory +
      product.generalGoodsPurchaseInventory;
  }

  _groupSalesData(salesData) {
    const groups = {};
    salesData.forEach((sale) => {
      if (!sale.itemNumber) return;
      if (!groups[sale.itemNumber]) {
        groups[sale.itemNumber] = [];
      }
      groups[sale.itemNumber].push(sale);
    });
    return groups;
  }

  _applySalesToProduct(product, sales, last7Days) {
    let changed = false;

    // 重置近7天数据
    const oldSales = product.salesQuantityOfLast7Days;

    product.salesQuantityOfLast7Days = 0;
    product.salesAmountOfLast7Days = 0;

    // 计算近7天销售
    sales.forEach((sale) => {
      if (sale.salesDate && sale.salesDate.includes(last7Days)) {
        product.salesQuantityOfLast7Days += Number(sale.salesQuantity || 0);
        product.salesAmountOfLast7Days += Number(sale.salesAmount || 0);
      }
    });

    // 计算件单价
    product.unitPriceOfLast7Days =
      product.salesQuantityOfLast7Days > 0
        ? product.salesAmountOfLast7Days / product.salesQuantityOfLast7Days
        : 0;

    if (oldSales !== product.salesQuantityOfLast7Days) {
      changed = true;
    }

    return changed;
  }

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

  _updateSystemRecord() {
    const systemRecord = this._repository.getSystemRecord();
    const now = validationEngine.formatDate(new Date());

    systemRecord.updateDateOfRegularProduct = now;
    systemRecord.updateDateOfProductPrice = now;
    systemRecord.updateDateOfInventory = now;
    systemRecord.updateDateOfProductSales = now;

    this._repository.save("SystemRecord", [systemRecord]);
  }

  /**
   * 生成更新报告
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
      report += `  更新价格: ${results.price.updated}\n`;
      report += `  无价格: ${results.price.skipped}\n`;
    }

    if (results.inventory) {
      report += `\n【商品库存】\n`;
      report += `  库存变动: ${results.inventory.updated}\n`;
      report += `  零库存: ${results.inventory.zeroInventory}\n`;
    }

    if (results.sales) {
      report += `\n【商品销售】\n`;
      report += `  销量变动: ${results.sales.updated}\n`;
      report += `  有销量: ${results.sales.withSales}\n`;
    }

    if (results.errors.length > 0) {
      report += `\n【错误信息】\n`;
      results.errors.forEach((err) => (report += `  ${err}\n`));
    }

    return report;
  }
}
