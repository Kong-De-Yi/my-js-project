// 全局服务实例
let _dataImportService = null;
let _productService = null;

// 程序主入口
function Main() {
  _initializeServices();
  UserForm1.Show();
}

// 初始化
function _initializeServices() {
  try {
    // 初始化底层基础服务
    const _excelDAO = new ExcelDAO();
    const _repository = new Repository(_excelDAO);
    const _profitCalculator = new ProfitCalculator(_repository);

    // 注册上下文
    _repository.setContext({
      profitCalculator: _profitCalculator,
    });

    // 注册所有索引
    const indexConfig = IndexConfig.getInstance();

    for (const [entityName, indexConfigs] of Object.entries(
      indexConfig.getAllIndexes(),
    )) {
      _repository.registerIndexes(entityName, indexConfigs);
    }

    // 初始化服务实例
    _dataImportService = new DataImportService(_repository, _excelDAO);
    _productService = new ProductService(_repository);
  } catch (e) {
    MsgBox(`系统初始化失败：${e.message}`, 0, "错误");
    throw e;
  }
}

// 导入数据
function UserForm1_CommandButton6_Click() {
  try {
    const result = _dataImportService.import();

    MsgBox(result.message, 64, "导入成功");
  } catch (err) {
    MsgBox(`导入失败：${err.message}`, 16, "错误");
  }
}

// 更新商品价格
function UserForm1_CommandButton1_Click() {
  try {
    // 刷新货号总表缓存中的数据
    _repository.refresh("Product");

    const result = _productService.updateFromPriceData();
    const updateReport = _productService.generateUpdateReport(result);

    MsgBox(updateReport, 64, "商品价格更新成功");
  } catch (err) {
    MsgBox(`商品价格更新失败：${err.message}`, 16, "错误");
  }
}

// 更新常态商品
function UserForm1_CommandButton2_Click() {
  try {
    // 刷新货号总表缓存中的数据
    _repository.refresh("Product");

    const result = _productService.updateFromRegularProducts();
    const updateReport = _productService.generateUpdateReport(result);

    MsgBox(updateReport, 64, "常态商品更新成功");
  } catch (err) {
    MsgBox(`常态商品更新失败：${err.message}`, 16, "错误");
  }
}

// 更新商品库存
function UserForm1_CommandButton4_Click() {
  try {
    // 刷新货号总表缓存中的数据
    _repository.refresh("Product");

    const result = _productService.updateFromInventory();
    const updateReport = _productService.generateUpdateReport(result);

    MsgBox(updateReport, 64, "商品库存更新成功");
  } catch (err) {
    MsgBox(`商品库存更新失败：${err.message}`, 16, "错误");
  }
}

// 更新商品销售
function UserForm1_CommandButton10_Click() {
  try {
    // 刷新货号总表缓存中的数据
    _repository.refresh("Product");

    const result = _productService.updateFromSalesData();
    const updateReport = _productService.generateUpdateReport(result);

    MsgBox(updateReport, 64, "商品销售更新成功");
  } catch (err) {
    MsgBox(`商品销售更新失败：${err.message}`, 16, "错误");
  }
}

// 一键更新
function UserForm1_CommandButton5_Click() {
  try {
    // 刷新货号总表缓存中的数据
    _repository.refresh("Product");

    const results = _productService.updateAll();
    const updateReport = _productService.generateUpdateReport(results);

    MsgBox(updateReport, 64, "一键更新");
  } catch (err) {
    MsgBox(`一键更新失败：${err.message}`, 16, "错误");
  }
}

// 商品下线选项和下线原因联动
function UserForm1_CheckBox47_Click() {
  UserForm1.CheckBox54.Value = false;
  UserForm1.CheckBox55.Value = false;
  UserForm1.CheckBox56.Value = false;

  if (UserForm1.CheckBox47.Value) {
    UserForm1.CheckBox54.Enabled = true;
    UserForm1.CheckBox55.Enabled = true;
    UserForm1.CheckBox56.Enabled = true;
  } else {
    UserForm1.CheckBox54.Enabled = false;
    UserForm1.CheckBox55.Enabled = false;
    UserForm1.CheckBox56.Enabled = false;
  }
}

// let _repository = null;
// let _excelDAO = null;
// let _dataImportService = null;
// let _reportEngine = null;

// let _activityService = null;

// function _initializeServices() {
//   if (_repository) return;

//   try {
//     _excelDAO = new ExcelDAO();
//     _repository = new Repository(_excelDAO);

//     _productService = new ProductService(_repository, _profitCalculator);
//     _dataImportService = new DataImportService(_repository, _excelDAO);
//     _reportEngine = new ReportEngine(_repository, _excelDAO);
//     _activityService = new ActivityService(
//       _repository,
//       _profitCalculator,
//       _excelDAO,
//     );

//     // 注册所有索引
//     const indexConfig = IndexConfig.getInstance();

//     _repository.registerIndexes("Product", indexConfig.getIndexes("Product"));
//     _repository.registerIndexes(
//       "ProductPrice",
//       indexConfig.getIndexes("ProductPrice"),
//     );
//     _repository.registerIndexes(
//       "RegularProduct",
//       indexConfig.getIndexes("RegularProduct"),
//     );
//     _repository.registerIndexes(
//       "Inventory",
//       indexConfig.getIndexes("Inventory"),
//     );
//     _repository.registerIndexes(
//       "ComboProduct",
//       indexConfig.getIndexes("ComboProduct"),
//     );
//     _repository.registerIndexes(
//       "ProductSales",
//       indexConfig.getIndexes("ProductSales"),
//     );
//     _repository.registerIndexes(
//       "BrandConfig",
//       indexConfig.getIndexes("BrandConfig"),
//     );
//     _repository.registerIndexes(
//       "ReportTemplate",
//       indexConfig.getIndexes("ReportTemplate"),
//     );

//     // 初始化报表模板
//     _reportEngine.initializeTemplates();
//   } catch (e) {
//     MsgBox(`系统初始化失败：${e.message}`, 0, "错误");
//     throw e;
//   }
// }

// // ========== 数据更新事件 ==========

// function UserForm1_CommandButton2_Click() {
//   _initializeServices();

//   try {
//     const result = _productService.updateFromRegularProducts();
//     const report = _productService.generateUpdateReport({ regular: result });
//     MsgBox(report, 64, "更新成功");
//   } catch (err) {
//     MsgBox(`更新失败：${err.message}`, 16, "错误");
//   }
// }

// function UserForm1_CommandButton1_Click() {
//   _initializeServices();

//   try {
//     const result = _productService.updateFromPriceData();
//     MsgBox(
//       `价格更新完成！\n更新: ${result.updated} 个商品\n跳过: ${result.skipped} 个商品`,
//       64,
//       "更新成功",
//     );
//   } catch (err) {
//     MsgBox(`更新失败：${err.message}`, 16, "错误");
//   }
// }

// function UserForm1_CommandButton4_Click() {
//   _initializeServices();

//   try {
//     const result = _productService.updateFromInventory();
//     MsgBox(
//       `库存更新完成！\n更新: ${result.updated} 个商品\n零库存: ${result.zeroInventory} 个商品`,
//       64,
//       "更新成功",
//     );
//   } catch (err) {
//     MsgBox(`更新失败：${err.message}`, 16, "错误");
//   }
// }

// function UserForm1_CommandButton5_Click() {
//   _initializeServices();

//   try {
//     const result = _productService.updateFromSalesData();
//     MsgBox(
//       `销售更新完成！\n更新: ${result.updated} 个商品\n有销量: ${result.withSales} 个商品`,
//       64,
//       "更新成功",
//     );
//   } catch (err) {
//     MsgBox(`更新失败：${err.message}`, 16, "错误");
//   }
// }

// function UserForm1_CommandButton6_Click() {
//   _initializeServices();

//   try {
//     const results = _productService.updateAll();
//     const report = _productService.generateUpdateReport(results);
//     MsgBox(report, 64, "一键更新完成");
//   } catch (err) {
//     MsgBox(`更新失败：${err.message}`, 16, "错误");
//   }
// }

// // ========== 数据导入事件 ==========

// function UserForm1_CommandButton15_Click() {
//   _initializeServices();

//   try {
//     const result = _dataImportService.import();
//     MsgBox(result.message, 64, "导入成功");
//   } catch (err) {
//     MsgBox(`导入失败：${err.message}`, 16, "错误");
//   }
// }

// // ========== 报表相关事件 ==========

// /**
//  * 报表输出（原按钮）
//  */
// function UserForm1_CommandButton13_Click() {
//   _initializeServices();

//   try {
//     _reportEngine.previewReport();
//   } catch (err) {
//     MsgBox(`报表生成失败：${err.message}`, 16, "错误");
//   }
// }

// /**
//  * 初始化报表模板
//  * 新增按钮：CommandButton25
//  */
// function UserForm1_CommandButton25_Click() {
//   _initializeServices();

//   try {
//     _reportEngine.initializeTemplates();
//     MsgBox("报表模板初始化成功！", 64, "成功");
//   } catch (err) {
//     MsgBox(`模板初始化失败：${err.message}`, 16, "错误");
//   }
// }

// /**
//  * 刷新报表模板（重新读取配置）
//  * 新增按钮：CommandButton26
//  */
// function UserForm1_CommandButton26_Click() {
//   _initializeServices();

//   try {
//     _reportEngine.refreshTemplates();
//     MsgBox("报表模板刷新成功！", 64, "成功");
//   } catch (err) {
//     MsgBox(`模板刷新失败：${err.message}`, 16, "错误");
//   }
// }

// /**
//  * 获取模板列表
//  * 新增：ComboBox7 - 模板选择下拉框
//  */
// function UserForm1_ComboBox7_DropButtonClick() {
//   _initializeServices();

//   try {
//     const templates = _reportEngine.getTemplateList();
//     UserForm1.ComboBox7.Clear();
//     templates.forEach((template) => {
//       UserForm1.ComboBox7.AddItem(template);
//     });
//   } catch (err) {
//     MsgBox(`获取模板列表失败：${err.message}`, 16, "错误");
//   }
// }

// /**
//  * 选择模板
//  */
// function UserForm1_ComboBox7_Change() {
//   _initializeServices();

//   const templateName = UserForm1.ComboBox7.Value;
//   if (templateName) {
//     try {
//       _reportEngine.setCurrentTemplate(templateName);
//     } catch (err) {
//       MsgBox(`选择模板失败：${err.message}`, 16, "错误");
//     }
//   }
// }

// /**
//  * 获取可用字段列表（用于调试/配置）
//  * 新增按钮：CommandButton27
//  */
// function UserForm1_CommandButton27_Click() {
//   _initializeServices();

//   try {
//     const fields = _reportEngine.getAvailableFields();

//     // 按分组显示字段
//     const groups = {};
//     fields.forEach((field) => {
//       const group = field.group || "其他";
//       if (!groups[group]) {
//         groups[group] = [];
//       }
//       groups[group].push(`${field.field} - ${field.title}`);
//     });

//     let message = "可用的统计字段：\n\n";
//     Object.entries(groups).forEach(([group, fieldList]) => {
//       message += `【${group}】\n`;
//       fieldList.forEach((f) => {
//         message += `  ${f}\n`;
//       });
//       message += "\n";
//     });

//     MsgBox(message, 64, "可用字段列表");
//   } catch (err) {
//     MsgBox(`获取字段列表失败：${err.message}`, 16, "错误");
//   }
// }

// // ========== 活动提报事件 ==========

// function UserForm1_CommandButton3_Click() {
//   _initializeServices();

//   try {
//     const newWb = _activityService.signUpActivity();
//     newWb.Activate();
//     MsgBox("活动提报表生成成功！", 64, "成功");
//   } catch (err) {
//     MsgBox(`活动提报失败：${err.message}`, 16, "错误");
//   }
// }

// // ========== 表单初始化事件 ==========

// function UserForm1_Initialize() {
//   _initializeServices();

//   // 初始化模板下拉框
//   try {
//     const templates = _reportEngine.getTemplateList();
//     UserForm1.ComboBox7.Clear();
//     templates.forEach((template) => {
//       UserForm1.ComboBox7.AddItem(template);
//     });
//   } catch (err) {
//     // 忽略错误
//   }
// }
