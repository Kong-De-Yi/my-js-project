// ============================================================================
// UI控制器
// 功能：处理用户界面交互，调用业务服务
// 特点：轻量级，只负责UI事件分发，无业务逻辑
// ============================================================================

// 全局服务实例
let _repository = null;
let _profitCalculator = null;
let _reportEngine = null;
let _activityService = null;

// 初始化服务
function _initializeServices() {
  if (_repository) return;

  try {
    const excelDAO = new ExcelDAO();
    _repository = new Repository(excelDAO);
    _profitCalculator = new ProfitCalculator(_repository);
    _reportEngine = new ReportEngine(_repository, excelDAO, _profitCalculator);
    _activityService = new ActivityService(
      _repository,
      _profitCalculator,
      excelDAO,
    );

    // 设置计算上下文
    _repository.setContext({
      profitCalculator: _profitCalculator,
    });
  } catch (e) {
    MsgBox(`系统初始化失败：${e.message}`, 0, "错误");
    throw e;
  }
}

// 显示窗口
function Macro() {
  _initializeServices();

  UserForm1.ComboBox3.AddItem([
    "直通车",
    "黄金等级",
    "TOP3",
    "白金等级",
    "白金限量",
  ]);

  UserForm1.ComboBox5.AddItem(["中台1", "中台2"]);

  UserForm1.ComboBox4.AddItem([
    "上市年份",
    "四级品类",
    "运营分类",
    "下线原因",
    "三级品类",
  ]);

  UserForm1.ComboBox6.AddItem([
    "首次上架时间",
    "成本价",
    "白金价",
    "利润",
    "利润率",
    "近7天件单价",
    "近7天曝光UV",
    "近7天商详UV",
    "近7天加购UV",
    "近7天客户数",
    "近7天拒退数",
    "近7天销售量",
    "近7天销售额",
    "近7天点击率",
    "近7天加购率",
    "近7天转化率",
    "近7天拒退率",
    "近7天款销量",
    "可售库存",
    "可售天数",
    "成品合计",
    "通货合计",
    "合计库存",
    "销量总计",
  ]);

  UserForm1.Show();
}

// 更新常态商品
function UserForm1_CommandButton2_Click() {
  _initializeServices();

  try {
    // 刷新商品主数据缓存
    _repository.refresh("Product");

    // 读取常态商品
    const regularProducts = _repository.findAll("RegularProduct");

    // 获取当前所有商品
    let products = _repository.findProducts();

    // 添加新货号
    regularProducts.forEach((rp) => {
      if (!rp.itemNumber) return;

      const existing = _repository.findProductByItemNumber(rp.itemNumber);
      if (!existing) {
        const newProduct = {
          itemNumber: rp.itemNumber,
          brandSN: rp.brandSN,
          marketingPositioning: "利润款",
          _rowNumber: products.length + 2,
        };
        products.push(newProduct);
      }
    });

    // 更新商品信息
    products.forEach((product) => {
      const regulars = regularProducts.filter(
        (rp) => rp.itemNumber === product.itemNumber,
      );

      if (regulars.length > 0) {
        const firstRegular = regulars[0];
        product.thirdLevelCategory = firstRegular.thirdLevelCategory;
        product.P_SPU = firstRegular.P_SPU;
        product.MID = firstRegular.MID;
        product.styleNumber = firstRegular.styleNumber;
        product.color = firstRegular.color;
        product.itemStatus = firstRegular.itemStatus;
        product.vipshopPrice = firstRegular.vipshopPrice;
        product.finalPrice = firstRegular.finalPrice;
        product.sellableDays = firstRegular.sellableDays;

        // 商品上架清空下线原因
        if (product.itemStatus !== "商品下线") {
          product.offlineReason = undefined;
        }
      }

      // 计算可售库存
      product.sellableInventory = regulars.reduce(
        (sum, rp) => sum + (Number(rp.sellableInventory) || 0),
        0,
      );

      // 计算是否断码
      const outOfStockSizes = regulars
        .filter(
          (rp) =>
            rp.sizeStatus === "尺码上线" && (rp.sellableInventory || 0) === 0,
        )
        .map((rp) => rp.size)
        .filter(Boolean);

      product.isOutOfStock =
        outOfStockSizes.length > 0 ? outOfStockSizes.join("/") : undefined;
    });

    // 保存商品数据
    _repository.save("Product", products);

    // 清空常态商品
    _repository.clear("RegularProduct");

    // 更新系统记录
    const systemRecord = _repository.getSystemRecord();
    systemRecord.updateDateOfRegularProduct = new Date().toString();
    _repository.save("SystemRecord", [systemRecord]);

    MsgBox(`【常态商品】更新成功！`);
  } catch (err) {
    _handleError(err);
  }
}

// 更新商品价格
function UserForm1_CommandButton1_Click() {
  _initializeServices();

  try {
    // 刷新价格数据
    _repository.refresh("ProductPrice");

    // 验证价格数据
    const priceData = _repository.findAll("ProductPrice");
    const validationResult = validationEngine.validateAll(
      priceData,
      DataConfig.getInstance().get("ProductPrice"),
    );

    if (!validationResult.valid) {
      const errorMsg = validationEngine.formatErrors(
        validationResult,
        "商品价格",
      );
      throw new Error(errorMsg);
    }

    // 更新商品价格
    const priceMap = {};
    priceData.forEach((p) => {
      priceMap[p.itemNumber] = p;
    });

    const products = _repository.findProducts();

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

    _repository.save("Product", products);

    // 更新系统记录
    const systemRecord = _repository.getSystemRecord();
    systemRecord.updateDateOfProductPrice = new Date().toString();
    _repository.save("SystemRecord", [systemRecord]);

    MsgBox(`【商品价格】更新成功！`);
  } catch (err) {
    _handleError(err);
  }
}

// 更新商品库存
function UserForm1_CommandButton4_Click() {
  _initializeServices();

  try {
    // 刷新库存和组合商品数据
    _repository.refresh("Inventory");
    _repository.refresh("ComboProduct");

    const products = _repository.findProducts();
    const inventoryMap = {};
    const comboMap = {};

    // 建立库存索引
    _repository.findAll("Inventory").forEach((inv) => {
      inventoryMap[inv.productCode] = inv;
    });

    // 建立组合商品索引
    _repository.findAll("ComboProduct").forEach((combo) => {
      if (!comboMap[combo.productCode]) {
        comboMap[combo.productCode] = [];
      }
      comboMap[combo.productCode].push(combo);
    });

    products.forEach((product) => {
      // 重置库存
      product.finishedGoodsMainInventory = 0;
      product.finishedGoodsIncomingInventory = 0;
      product.finishedGoodsFinishingInventory = 0;
      product.finishedGoodsOversoldInventory = 0;
      product.finishedGoodsPrepareInventory = 0;
      product.finishedGoodsReturnInventory = 0;
      product.finishedGoodsPurchaseInventory = 0;
      product.finishedGoodsTotalInventory = 0;

      product.generalGoodsMainInventory = 0;
      product.generalGoodsIncomingInventory = 0;
      product.generalGoodsFinishingInventory = 0;
      product.generalGoodsOversoldInventory = 0;
      product.generalGoodsPrepareInventory = 0;
      product.generalGoodsReturnInventory = 0;
      product.generalGoodsPurchaseInventory = 0;
      product.generalGoodsTotalInventory = 0;

      product.totalInventory = 0;

      // 查找该货号的所有常态商品
      const regulars = _repository.findRegularProducts({
        itemNumber: product.itemNumber,
      });

      regulars.forEach((regular) => {
        // 成品库存
        const inv = inventoryMap[regular.productCode];
        if (inv) {
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
          product.finishedGoodsPrepareInventory += Number(
            inv.prepareInventory || 0,
          );
          product.finishedGoodsReturnInventory += Number(
            inv.returnInventory || 0,
          );
          product.finishedGoodsPurchaseInventory += Number(
            inv.purchaseInventory || 0,
          );
        }

        // 成品合计
        product.finishedGoodsTotalInventory =
          product.finishedGoodsMainInventory +
          product.finishedGoodsIncomingInventory +
          product.finishedGoodsFinishingInventory +
          product.finishedGoodsOversoldInventory +
          product.finishedGoodsPrepareInventory +
          product.finishedGoodsReturnInventory +
          product.finishedGoodsPurchaseInventory;

        // 通货库存（组合商品）
        const combos = comboMap[regular.productCode] || [];

        combos.forEach((combo) => {
          // 排除YH/FL开头的虚拟商品
          if (
            combo.subProductCode.startsWith("YH") ||
            combo.subProductCode.startsWith("FL")
          ) {
            return;
          }

          const subInv = inventoryMap[combo.subProductCode];
          const quantity = Number(combo.subProductQuantity) || 1;

          if (subInv) {
            // 修复：子商品库存 × 组合数量
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

        // 通货合计
        product.generalGoodsTotalInventory =
          product.generalGoodsMainInventory +
          product.generalGoodsIncomingInventory +
          product.generalGoodsFinishingInventory +
          product.generalGoodsOversoldInventory +
          product.generalGoodsPrepareInventory +
          product.generalGoodsReturnInventory +
          product.generalGoodsPurchaseInventory;
      });

      // 合计库存
      product.totalInventory =
        product.finishedGoodsTotalInventory +
        product.generalGoodsTotalInventory;
    });

    _repository.save("Product", products);

    // 清空库存表和组合商品表
    _repository.clear("Inventory");
    _repository.clear("ComboProduct");

    // 更新系统记录
    const systemRecord = _repository.getSystemRecord();
    systemRecord.updateDateOfInventory = new Date().toString();
    _repository.save("SystemRecord", [systemRecord]);

    MsgBox(`【商品库存】更新成功！`);
  } catch (err) {
    _handleError(err);
  }
}

// 更新商品销售
function UserForm1_CommandButton5_Click() {
  _initializeServices();

  try {
    // 刷新销售数据
    _repository.refresh("ProductSales");

    const dateOfLast7Days = _getLast7DaysRange();
    const systemRecord = _repository.getSystemRecord();
    const needUpdate = systemRecord.updateDateOfLast7Days !== dateOfLast7Days;

    const products = _repository.findProducts();
    const salesData = _repository.findAll("ProductSales");

    // 按货号+日期建立索引
    const salesMap = {};
    salesData.forEach((sale) => {
      const key = `${sale.itemNumber}|${sale.salesDate}`;
      salesMap[key] = sale;
    });

    products.forEach((product) => {
      // 近7天数据
      if (needUpdate) {
        product.exposureUVOfLast7Days = 0;
        product.productDetailsUVOfLast7Days = 0;
        product.addToCartUVOfLast7Days = 0;
        product.customerCountOfLast7Days = 0;
        product.rejectAndReturnCountOfLast7Days = 0;
        product.salesQuantityOfLast7Days = 0;
        product.salesAmountOfLast7Days = 0;
        product.styleSalesOfLast7Days = 0;
      }

      const sevenDaysKey = `${product.itemNumber}|${dateOfLast7Days}`;
      const sevenDaysSale = salesMap[sevenDaysKey];

      if (sevenDaysSale) {
        product.exposureUVOfLast7Days = Number(sevenDaysSale.exposureUV || 0);
        product.productDetailsUVOfLast7Days = Number(
          sevenDaysSale.productDetailsUV || 0,
        );
        product.addToCartUVOfLast7Days = Number(sevenDaysSale.addToCartUV || 0);
        product.customerCountOfLast7Days = Number(
          sevenDaysSale.customerCount || 0,
        );
        product.rejectAndReturnCountOfLast7Days = Number(
          sevenDaysSale.rejectAndReturnCount || 0,
        );
        product.salesQuantityOfLast7Days = Number(
          sevenDaysSale.salesQuantity || 0,
        );
        product.salesAmountOfLast7Days = Number(sevenDaysSale.salesAmount || 0);
        product.firstListingTime = sevenDaysSale.firstListingTime
          ? `'${sevenDaysSale.firstListingTime}`
          : "";
      }

      // 计算率值
      product.unitPriceOfLast7Days = product.salesQuantityOfLast7Days
        ? product.salesAmountOfLast7Days / product.salesQuantityOfLast7Days
        : undefined;

      product.clickThroughRateOfLast7Days = product.exposureUVOfLast7Days
        ? product.productDetailsUVOfLast7Days / product.exposureUVOfLast7Days
        : undefined;

      product.addToCartRateOfLast7Days = product.productDetailsUVOfLast7Days
        ? product.addToCartUVOfLast7Days / product.productDetailsUVOfLast7Days
        : undefined;

      product.purchaseRateOfLast7Days = product.productDetailsUVOfLast7Days
        ? product.customerCountOfLast7Days / product.productDetailsUVOfLast7Days
        : undefined;

      product.rejectAndReturnRateOfLast7Days = product.salesQuantityOfLast7Days
        ? product.rejectAndReturnCountOfLast7Days /
          product.salesQuantityOfLast7Days
        : undefined;

      // 历史销量
      product.totalSales = 0;
    });

    // 计算款销量
    const styleSales = {};
    products.forEach((product) => {
      if (product.styleNumber) {
        styleSales[product.styleNumber] =
          (styleSales[product.styleNumber] || 0) +
          (product.salesQuantityOfLast7Days || 0);
      }
    });

    products.forEach((product) => {
      product.styleSalesOfLast7Days = styleSales[product.styleNumber] || 0;
    });

    _repository.save("Product", products);

    // 更新系统记录
    if (needUpdate) {
      systemRecord.updateDateOfLast7Days = dateOfLast7Days;
    }
    systemRecord.updateDateOfProductSales = new Date().toString();
    _repository.save("SystemRecord", [systemRecord]);

    // 清空销售表
    _repository.clear("ProductSales");

    MsgBox(`【商品销售】更新成功！`);
  } catch (err) {
    _handleError(err);
  }
}

// 一键更新
function UserForm1_CommandButton6_Click() {
  _initializeServices();

  try {
    UserForm1_CommandButton2_Click(); // 常态商品
    UserForm1_CommandButton1_Click(); // 商品价格
    UserForm1_CommandButton4_Click(); // 商品库存
    UserForm1_CommandButton5_Click(); // 商品销售

    MsgBox("一键更新成功！");
  } catch (err) {
    _handleError(err);
  }
}

// 报表输出
function UserForm1_CommandButton13_Click() {
  _initializeServices();

  try {
    // 检查数据更新状态
    const systemRecord = _repository.getSystemRecord();
    const today = new Date();

    const checkDate = (dateStr, name) => {
      if (!dateStr) throw new Error(`【${name}】尚未更新`);
      const ts = Date.parse(dateStr);
      if (isNaN(ts)) throw new Error(`【${name}】更新日期格式错误`);

      const date = new Date(ts);
      if (
        date.getDate() !== today.getDate() ||
        date.getMonth() !== today.getMonth() ||
        date.getFullYear() !== today.getFullYear()
      ) {
        if (MsgBox(`【${name}】今日尚未更新，是否继续？`, 4, "提醒") === 7) {
          throw new Error(`请更新【${name}】后再重试`);
        }
      }
    };

    checkDate(systemRecord.updateDateOfProductPrice, "商品价格");
    checkDate(systemRecord.updateDateOfRegularProduct, "常态商品");
    checkDate(systemRecord.updateDateOfInventory, "商品库存");

    if (systemRecord.updateDateOfLast7Days !== _getLast7DaysRange()) {
      if (
        MsgBox("【近7天商品销售数据】尚未更新，是否继续？", 4, "提醒") === 7
      ) {
        throw new Error("请更新【近7天商品销售数据】后再重试");
      }
    }

    if (!systemRecord.updateDateOfProductSales) {
      if (MsgBox("【商品销售】昨日数据尚未更新，是否继续？", 4, "提醒") === 7) {
        throw new Error("请更新【商品销售】昨日数据后再重试");
      }
    }

    // 生成报表
    const newWb = _reportEngine.generateReport();

    MsgBox("报表输出成功！");
  } catch (err) {
    _handleError(err);
  }
}

// 组合商品
function UserForm1_CheckBox11_Click() {
  if (UserForm1.CheckBox11.Value) {
    UserForm1.TextEdit19.Value = 1;
    UserForm1.TextEdit20.Value = "";
  } else {
    UserForm1.TextEdit19.Value = "";
    UserForm1.TextEdit20.Value = "";
  }
}

// 平台活动提报
function UserForm1_CommandButton9_Click() {
  _initializeServices();

  try {
    const newWb = _activityService.signUpActivity();
    MsgBox("平台活动导入表输出成功！");
  } catch (err) {
    _handleError(err);
  }
}

// 错误处理
function _handleError(err) {
  MsgBox(err.message);

  if (err instanceof CustomError) {
    const wb = Workbooks.Add();
    const excelDAO = new ExcelDAO();
    excelDAO.write("Product", err.data, wb);
  }
}

// 获取近7天日期范围
function _getLast7DaysRange() {
  const today = new Date();
  const start = new Date(today);
  start.setDate(today.getDate() - 7);
  const end = new Date(today);
  end.setDate(today.getDate() - 1);

  const format = (date) =>
    `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;

  return `${format(start)}~${format(end)}`;
}

// 自定义错误类
class CustomError extends Error {
  constructor(message, keyToTitle, data) {
    super(message);
    this.name = "CustomError";
    this.keyToTitle = keyToTitle;
    this.data = data;
  }
}
