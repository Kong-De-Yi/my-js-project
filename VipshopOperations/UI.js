// ============================================================================
// UI.js - 简化后的UI控制器
// ============================================================================

// 全局服务实例
let _repository = null;
let _excelDAO = null;
let _dataImportService = null;
let _reportEngine = null;
let _profitCalculator = null;
let _productService = null;
let _activityService = null;

function _initializeServices() {
  if (_repository) return;

  try {
    _excelDAO = new ExcelDAO();
    _repository = new Repository(_excelDAO);
    _profitCalculator = new ProfitCalculator(_repository);
    _productService = new ProductService(_repository, _profitCalculator);
    _dataImportService = new DataImportService(_repository, _excelDAO);
    _reportEngine = new ReportEngine(_repository, _excelDAO);
    _activityService = new ActivityService(
      _repository,
      _profitCalculator,
      _excelDAO,
    );

    // 设置计算上下文
    _repository.setContext({
      profitCalculator: _profitCalculator,
    });

    // 注册索引
    _repository.registerIndexes("Product", indexConfig.getIndexes("Product"));
    _repository.registerIndexes(
      "ProductPrice",
      indexConfig.getIndexes("ProductPrice"),
    );
    _repository.registerIndexes(
      "RegularProduct",
      indexConfig.getIndexes("RegularProduct"),
    );
    _repository.registerIndexes(
      "Inventory",
      indexConfig.getIndexes("Inventory"),
    );
    _repository.registerIndexes(
      "ComboProduct",
      indexConfig.getIndexes("ComboProduct"),
    );
    _repository.registerIndexes(
      "ProductSales",
      indexConfig.getIndexes("ProductSales"),
    );
  } catch (e) {
    MsgBox(`系统初始化失败：${e.message}`, 0, "错误");
    throw e;
  }
}

// 更新常态商品
function UserForm1_CommandButton2_Click() {
  _initializeServices();

  try {
    const result = _productService.updateFromRegularProducts();
    const report = _productService.generateUpdateReport({ regular: result });
    MsgBox(report, 64, "更新成功");
  } catch (err) {
    MsgBox(`更新失败：${err.message}`, 16, "错误");
  }
}

// 更新商品价格
function UserForm1_CommandButton1_Click() {
  _initializeServices();

  try {
    const result = _productService.updateFromPriceData();
    MsgBox(
      `价格更新完成！\n更新: ${result.updated} 个商品\n跳过: ${result.skipped} 个商品`,
      64,
      "更新成功",
    );
  } catch (err) {
    MsgBox(`更新失败：${err.message}`, 16, "错误");
  }
}

// 更新商品库存
function UserForm1_CommandButton4_Click() {
  _initializeServices();

  try {
    const result = _productService.updateFromInventory();
    MsgBox(
      `库存更新完成！\n更新: ${result.updated} 个商品\n零库存: ${result.zeroInventory} 个商品`,
      64,
      "更新成功",
    );
  } catch (err) {
    MsgBox(`更新失败：${err.message}`, 16, "错误");
  }
}

// 更新商品销售
function UserForm1_CommandButton5_Click() {
  _initializeServices();

  try {
    const result = _productService.updateFromSalesData();
    MsgBox(
      `销售更新完成！\n更新: ${result.updated} 个商品\n有销量: ${result.withSales} 个商品`,
      64,
      "更新成功",
    );
  } catch (err) {
    MsgBox(`更新失败：${err.message}`, 16, "错误");
  }
}

// 一键更新
function UserForm1_CommandButton6_Click() {
  _initializeServices();

  try {
    const results = _productService.updateAll();
    const report = _productService.generateUpdateReport(results);
    MsgBox(report, 64, "一键更新完成");
  } catch (err) {
    MsgBox(`更新失败：${err.message}`, 16, "错误");
  }
}

// 导入数据
function UserForm1_CommandButton15_Click() {
  _initializeServices();

  try {
    const result = _dataImportService.import();
    MsgBox(result.message, 64, "导入成功");
  } catch (err) {
    MsgBox(`导入失败：${err.message}`, 16, "错误");
  }
}

// 报表输出
function UserForm1_CommandButton13_Click() {
  _initializeServices();

  try {
    const newWb = _reportEngine.generateReport();
    MsgBox("报表输出成功！", 64, "成功");
  } catch (err) {
    MsgBox(`报表生成失败：${err.message}`, 16, "错误");
  }
}

// ... 其他UI事件处理 ...
