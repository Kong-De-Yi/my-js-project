class ProductService {
  constructor(repository) {
    this._repository = repository;
    this._config = DataConfig.getInstance();
    this._validationEngine = ValidationEngine.getInstance();
  }

  // 计算常态商品是否断码
  _getOutOfStockSizes(regulars) {
    const outOfStockSizes = regulars
      .filter(
        (rp) => rp.sizeStatus === "尺码上线" && rp.sellableInventory === 0,
      )
      .map((rp) => rp.size)
      .filter(Boolean)
      .sort((a, b) => +a - +b);

    return outOfStockSizes.length > 0 ? outOfStockSizes.join("/") : "";
  }

  // 从常态商品更新指定产品
  _updateProductFromRegulars(product) {
    const regulars = this._repository.findRegularProductsByItemNumber(
      product.itemNumber,
    );
    if (regulars.length === 0) return product;

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
      product.offlineReason = "";
    }

    return product;
  }

  // 根据货号取常态商品信息创建新的产品
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

  // 根据货号取商品价格信息更新指定产品
  _applyPriceToProduct(product) {
    let changed = false;
    const price = this._repository.findPriceByItemNumber(product.itemNumber);
    if (!price) return changed;

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

    return changed;
  }

  // 指定货号库存清0
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

  // 计算产品的成品库存
  _calculateFinishedGoods(product) {
    // 查找该货号的所有常态商品
    const regulars = this._repository.findRegularProductsByItemNumber(
      product.itemNumber,
    );

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

  // 计算产品的通货库存
  _calculateGeneralGoods(product) {
    // 查找该货号的所有常态商品
    const regulars = this._repository.findRegularProductsByItemNumber(
      product.itemNumber,
    );

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

  // 检查业务实体是否为今日最新
  _checkDataExpired(entityName) {
    if (
      !["ProductPrice", "RegularProduct", "Inventory", "ComboProduct"].includes(
        entityName,
      )
    )
      return false;

    const systemRecord = this._repository.getSystemRecord();
    let importDate = null;

    switch (entityName) {
      case "ProductPrice":
        importDate = systemRecord.importDateOfProductPrice;
        break;

      case "RegularProduct":
        importDate = systemRecord.importDateOfRegularProduct;
        break;

      case "Inventory":
        importDate = systemRecord.importDateOfInventory;
        break;

      case "ComboProduct":
        importDate = systemRecord.importDateOfComboProduct;
        break;
    }

    // 没有更新日期默认过期
    if (!importDate) return true;

    importDate = Date.parse(importDate);
    return new Date() - importDate > 12 * 60 * 60 * 1000;
  }

  // 更新系统记录
  _updateSystemRecord(entityName) {
    if (!["RegularProduct", "ProductPrice", "Inventory"].includes(entityName))
      return;

    const systemRecord = this._repository.getSystemRecord();
    const now = new Date();

    switch (entityName) {
      case "RegularProduct":
        systemRecord.updateDateOfRegularProduct = now;
        break;
      case "ProductPrice":
        systemRecord.updateDateOfProductPrice = now;
        break;
      case "Inventory":
        systemRecord.updateDateOfInventory = now;
        break;
    }

    this._repository.save("SystemRecord", [systemRecord]);
  }

  // 获取最新的商品
  _getLatestProducts() {
    this._repository.refresh("Product");
    return this._repository.findProducts();
  }

  // 从常态商品更新数据
  _updateFromRegularProducts(products) {
    const result = {
      products: [],
      totalProducts: 0,
      newProducts: 0,
      updatedProducts: 0,
    };

    result.totalProducts = products.length;

    // 1.检查常态商品是否为最新
    if (this._checkDataExpired("RegularProduct")) {
      throw new Error(`【常态商品】今日尚未导入，请先导入！`);
    }

    // 2. 更新现有常态商品
    const updatedProducts = products.map((product) => {
      return this._updateProductFromRegulars(product);
    });

    // 3. 添加新货号
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

    // 4.合并数据
    const allProducts = [...updatedProducts, ...newProducts];
    result.updatedProducts = updatedProducts.length;

    result.products = allProducts;

    return result;
  }

  // 从价格表更新产品价格
  _updateFromPriceData(products) {
    const result = {
      products: [],
      updated: 0,
      skipped: 0,
    };

    // 1.检查商品价格是否为最新
    if (this._checkDataExpired("ProductPrice")) {
      throw new Error(`【商品价格】今日尚未导入，请先导入！`);
    }

    // 2.更新产品价格
    products.forEach((product) => {
      const changed = this._applyPriceToProduct(product);
      if (changed) {
        result.updated++;
      } else {
        result.skipped++;
      }
    });

    result.products = products;

    return result;
  }

  // 从库存数据更新产品库存
  _updateFromInventory(products) {
    const result = {
      products: [],
      updated: 0,
      zeroInventory: 0,
    };

    // 1.检查组合装和商品库存是否更新
    if (this._checkDataExpired("ComboProduct")) {
      throw new Error("【组合商品】今日尚未导入，请先导入！");
    }
    if (this._checkDataExpired("Inventory")) {
      throw new Error("【商品库存】今日尚未导入，请先导入！");
    }

    // 2.计算库存
    products.forEach((product) => {
      const before = product.totalInventory;
      // 1.重置库存
      this._resetInventoryFields(product);
      // 2.计算成品库存
      const fgTotalInventory = this._calculateFinishedGoods(product);
      // 3.计算通货库存
      const ggTotalInventory = this._calculateGeneralGoods(product);

      const after = fgTotalInventory + ggTotalInventory;

      if (before !== after) result.updated++;
      if (after === 0) result.zeroInventory++;
    });

    result.products = products;
    return result;
  }

  // 更新常态商品
  updateRegularProduct() {
    const products = this._getLatestProducts();
    const result = this._updateFromRegularProducts(products);
    this._repository.save("Product", result.products);
    this._updateSystemRecord("RegularProduct");
    return result;
  }

  // 更新商品价格
  updateProductPrice() {
    const products = this._getLatestProducts();
    const result = this._updateFromPriceData(products);
    this._repository.save("Product", result.products);
    this._updateSystemRecord("ProductPrice");
    return result;
  }

  // 更新商品库存
  updateInventory() {
    const products = this._getLatestProducts();
    const result = this._updateFromInventory(products);
    this._repository.save("Product", result.products);
    this._updateSystemRecord("Inventory");
    return result;
  }

  // 一键更新
  updateAll() {
    const results = { errors: [] };
    let result = null;
    let products = this._getLatestProducts();

    try {
      result = this._updateFromRegularProducts(products);
      this._updateSystemRecord("RegularProduct");
      results.regular = result;
      products = result.products;
    } catch (e) {
      results.errors.push(e.message);
    }

    try {
      result = this._updateFromPriceData(products);
      this._updateSystemRecord("ProductPrice");
      results.price = result;
      products = result.products;
    } catch (e) {
      results.errors.push(e.message);
    }

    try {
      result = this._updateFromInventory(products);
      this._updateSystemRecord("Inventory");
      results.inventory = result;
      products = result.products;
    } catch (e) {
      results.errors.push(e.message);
    }

    this._repository.save("Product", products);

    return results;
  }

  // 生成更新报告
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

    if (results?.errors?.length > 0) {
      report += `\n【错误信息】\n`;
      results.errors.forEach((err) => (report += `  ${err}\n`));
    }

    return report;
  }
}
