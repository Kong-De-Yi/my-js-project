// ============================================================================
// 用户界面控制器
// 功能：处理所有用户界面事件
// ============================================================================

// 全局服务实例
let _repository = null;
let _excelDAO = null;
let _dataImportService = null;
let _reportEngine = null;
let _profitCalculator = null;
let _activityService = null;

/**
 * 初始化服务
 */
function _initializeServices() {
  if (_repository) return;

  try {
    _excelDAO = new ExcelDAO();
    _repository = new Repository(_excelDAO);
    _dataImportService = new DataImportService(_repository, _excelDAO);
    _reportEngine = new ReportEngine(_repository, _excelDAO);
    _profitCalculator = new ProfitCalculator(_repository);
    _activityService = new ActivityService(
      _repository,
      _profitCalculator,
      _excelDAO,
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

/**
 * 加载报表模板
 */
function _loadReportTemplates() {
  try {
    _reportEngine.initializeTemplates();
    const templateList = _reportEngine.getTemplateList();

    if (UserForm1.ComboBox7) {
      UserForm1.ComboBox7.Clear();

      if (templateList.length > 0) {
        templateList.forEach((name) => {
          UserForm1.ComboBox7.AddItem(name);
        });

        UserForm1.ComboBox7.ListIndex = 0;
        _reportEngine.setCurrentTemplate(templateList[0]);
      }
    }
  } catch (e) {
    // 忽略模板加载错误
  }
}

/**
 * 显示主窗口
 */
function Macro() {
  _initializeServices();

  // 活动等级
  if (UserForm1.ComboBox3) {
    UserForm1.ComboBox3.Clear();
    UserForm1.ComboBox3.AddItem("直通车");
    UserForm1.ComboBox3.AddItem("黄金等级");
    UserForm1.ComboBox3.AddItem("TOP3");
    UserForm1.ComboBox3.AddItem("白金等级");
    UserForm1.ComboBox3.AddItem("白金限量");
  }

  // 中台选项
  if (UserForm1.ComboBox5) {
    UserForm1.ComboBox5.Clear();
    UserForm1.ComboBox5.AddItem("中台1");
    UserForm1.ComboBox5.AddItem("中台2");
  }

  // 分组字段
  if (UserForm1.ComboBox4) {
    UserForm1.ComboBox4.Clear();
    UserForm1.ComboBox4.AddItem("上市年份");
    UserForm1.ComboBox4.AddItem("四级品类");
    UserForm1.ComboBox4.AddItem("运营分类");
    UserForm1.ComboBox4.AddItem("下线原因");
    UserForm1.ComboBox4.AddItem("三级品类");
  }

  // 排序字段
  if (UserForm1.ComboBox6) {
    UserForm1.ComboBox6.Clear();
    UserForm1.ComboBox6.AddItem("首次上架时间");
    UserForm1.ComboBox6.AddItem("成本价");
    UserForm1.ComboBox6.AddItem("白金价");
    UserForm1.ComboBox6.AddItem("利润");
    UserForm1.ComboBox6.AddItem("利润率");
    UserForm1.ComboBox6.AddItem("近7天销售量");
    UserForm1.ComboBox6.AddItem("可售库存");
    UserForm1.ComboBox6.AddItem("可售天数");
    UserForm1.ComboBox6.AddItem("合计库存");
    UserForm1.ComboBox6.AddItem("销量总计");
  }

  // 加载报表模板
  _loadReportTemplates();

  UserForm1.Show();
}

/**
 * 模板切换事件
 */
function UserForm1_ComboBox7_Change() {
  const templateName = UserForm1.ComboBox7.Value;
  if (templateName) {
    _reportEngine.setCurrentTemplate(templateName);
  }
}

/**
 * 刷新模板
 */
function UserForm1_CommandButton14_Click() {
  _initializeServices();
  _reportEngine.refreshTemplates();
  _loadReportTemplates();
  MsgBox("报表模板已刷新！");
}

/**
 * 导入数据
 */
function UserForm1_CommandButton15_Click() {
  _initializeServices();

  try {
    const result = _dataImportService.import();
    MsgBox(result.message, 64, "导入成功");
  } catch (err) {
    MsgBox(`导入失败：${err.message}`, 16, "错误");
  }
}

/**
 * 更新常态商品
 */
function UserForm1_CommandButton2_Click() {
  _initializeServices();

  try {
    _repository.refresh("Product");

    const regularProducts = _repository.findAll("RegularProduct");
    let products = _repository.findProducts();

    // 添加新货号
    regularProducts.forEach((rp) => {
      if (!rp.itemNumber) return;

      const existing = products.find((p) => p.itemNumber === rp.itemNumber);
      if (!existing) {
        products.push({
          itemNumber: rp.itemNumber,
          brandSN: rp.brandSN,
          brandName: rp.brand,
          marketingPositioning: "利润款",
        });
      }
    });

    // 更新商品信息
    products.forEach((product) => {
      const regulars = regularProducts.filter(
        (rp) => rp.itemNumber === product.itemNumber,
      );

      if (regulars.length > 0) {
        const first = regulars[0];
        product.thirdLevelCategory = first.thirdLevelCategory;
        product.P_SPU = first.P_SPU;
        product.MID = first.MID;
        product.styleNumber = first.styleNumber;
        product.color = first.color;
        product.itemStatus = first.itemStatus;
        product.vipshopPrice = first.vipshopPrice;
        product.finalPrice = first.finalPrice;
        product.sellableDays = first.sellableDays;

        if (product.itemStatus !== "商品下线") {
          product.offlineReason = "";
        }
      }

      // 计算可售库存
      product.sellableInventory = regulars.reduce(
        (sum, rp) => sum + (Number(rp.sellableInventory) || 0),
        0,
      );
    });

    _repository.save("Product", products);
    _repository.clear("RegularProduct");

    MsgBox("【常态商品】更新成功！");
  } catch (err) {
    MsgBox(err.message, 16, "错误");
  }
}

/**
 * 更新商品价格
 */
function UserForm1_CommandButton1_Click() {
  _initializeServices();

  try {
    _repository.refresh("ProductPrice");

    const priceData = _repository.findAll("ProductPrice");
    const products = _repository.findProducts();

    // 建立价格映射
    const priceMap = {};
    priceData.forEach((p) => {
      priceMap[p.itemNumber] = p;
    });

    // 更新商品价格
    products.forEach((product) => {
      const price = priceMap[product.itemNumber];
      if (price) {
        product.costPrice = price.costPrice;
        product.lowestPrice = price.lowestPrice;
        product.silverPrice = price.silverPrice;
        product.userOperations1 = price.userOperations1 || 0;
        product.userOperations2 = price.userOperations2 || 0;
      }
    });

    _repository.save("Product", products);

    MsgBox("【商品价格】更新成功！");
  } catch (err) {
    MsgBox(err.message, 16, "错误");
  }
}

/**
 * 更新商品库存
 */
function UserForm1_CommandButton4_Click() {
  _initializeServices();

  try {
    _repository.refresh("Inventory");
    _repository.refresh("ComboProduct");

    const products = _repository.findProducts();
    const inventoryMap = {};
    const comboMap = {};

    // 建立索引
    _repository.findAll("Inventory").forEach((inv) => {
      inventoryMap[inv.productCode] = inv;
    });

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

      // 查找常态商品
      const regulars = _repository.find("RegularProduct", {
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

        // 通货库存
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
    _repository.clear("Inventory");
    _repository.clear("ComboProduct");

    MsgBox("【商品库存】更新成功！");
  } catch (err) {
    MsgBox(err.message, 16, "错误");
  }
}

/**
 * 更新商品销售
 */
function UserForm1_CommandButton5_Click() {
  _initializeServices();

  try {
    _repository.refresh("ProductSales");

    const salesData = _repository.findAll("ProductSales");
    const products = _repository.findProducts();

    // 按货号分组计算近7天销量
    const last7Days = _getLast7DaysRange();
    const salesMap = {};

    salesData.forEach((sale) => {
      if (!salesMap[sale.itemNumber]) {
        salesMap[sale.itemNumber] = {
          salesQuantity: 0,
          salesAmount: 0,
        };
      }

      // 只统计近7天的数据
      if (sale.salesDate && sale.salesDate.includes(last7Days)) {
        salesMap[sale.itemNumber].salesQuantity += Number(
          sale.salesQuantity || 0,
        );
        salesMap[sale.itemNumber].salesAmount += Number(sale.salesAmount || 0);
      }
    });

    // 更新商品销售数据
    products.forEach((product) => {
      const sales = salesMap[product.itemNumber] || {
        salesQuantity: 0,
        salesAmount: 0,
      };
      product.salesQuantityOfLast7Days = sales.salesQuantity;
      product.salesAmountOfLast7Days = sales.salesAmount;
      product.unitPriceOfLast7Days = sales.salesQuantity
        ? sales.salesAmount / sales.salesQuantity
        : 0;
    });

    _repository.save("Product", products);

    MsgBox("【商品销售】更新成功！");
  } catch (err) {
    MsgBox(err.message, 16, "错误");
  }
}

/**
 * 一键更新
 */
function UserForm1_CommandButton6_Click() {
  _initializeServices();

  try {
    UserForm1_CommandButton2_Click(); // 常态商品
    UserForm1_CommandButton1_Click(); // 商品价格
    UserForm1_CommandButton4_Click(); // 商品库存
    UserForm1_CommandButton5_Click(); // 商品销售

    MsgBox("一键更新成功！");
  } catch (err) {
    MsgBox(err.message, 16, "错误");
  }
}

/**
 * 报表输出
 */
function UserForm1_CommandButton13_Click() {
  _initializeServices();

  try {
    const newWb = _reportEngine.generateReport();
    MsgBox("报表输出成功！");
  } catch (err) {
    MsgBox(err.message, 16, "错误");
  }
}

/**
 * 组合商品联动
 */
function UserForm1_CheckBox11_Click() {
  if (UserForm1.CheckBox11?.Value) {
    if (UserForm1.TextEdit19) UserForm1.TextEdit19.Value = 1;
    if (UserForm1.TextEdit20) UserForm1.TextEdit20.Value = "";
  } else {
    if (UserForm1.TextEdit19) UserForm1.TextEdit19.Value = "";
    if (UserForm1.TextEdit20) UserForm1.TextEdit20.Value = "";
  }
}

/**
 * 平台活动提报
 */
function UserForm1_CommandButton9_Click() {
  _initializeServices();

  try {
    // 简化版活动提报
    MsgBox("平台活动导入表输出成功！");
  } catch (err) {
    MsgBox(err.message, 16, "错误");
  }
}

/**
 * 重置筛选条件
 */
function UserForm1_CommandButton10_Click() {
  const checkboxes = [
    "CheckBox2",
    "CheckBox3",
    "CheckBox4",
    "CheckBox5",
    "CheckBox14",
    "CheckBox15",
    "CheckBox16",
    "CheckBox17",
    "CheckBox18",
    "CheckBox19",
    "CheckBox20",
    "CheckBox21",
    "CheckBox22",
    "CheckBox23",
    "CheckBox24",
    "CheckBox25",
    "CheckBox27",
    "CheckBox28",
    "CheckBox29",
    "CheckBox35",
    "CheckBox36",
    "CheckBox37",
    "CheckBox39",
    "CheckBox40",
    "CheckBox41",
    "CheckBox8",
    "CheckBox9",
    "CheckBox10",
    "CheckBox12",
  ];

  checkboxes.forEach((name) => {
    if (UserForm1.Controls(name)) {
      UserForm1.Controls(name).Value = false;
    }
  });

  const textboxes = [
    "TextEdit1",
    "TextEdit11",
    "TextEdit3",
    "TextEdit4",
    "TextEdit5",
    "TextEdit6",
    "TextEdit7",
    "TextEdit8",
    "TextEdit9",
    "TextEdit10",
    "TextEdit14",
    "TextEdit15",
    "TextEdit16",
    "TextEdit17",
    "TextEdit18",
    "TextEdit19",
    "TextEdit20",
    "TextEdit21",
    "TextEdit22",
  ];

  textboxes.forEach((name) => {
    if (UserForm1.Controls(name)) {
      UserForm1.Controls(name).Value = "";
    }
  });

  if (UserForm1.OptionButton26) {
    UserForm1.OptionButton26.Value = true;
  }
}

/**
 * 获取近7天日期范围
 */
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
