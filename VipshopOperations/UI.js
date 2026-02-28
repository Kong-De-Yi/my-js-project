// 全局服务实例
let _dataImportService = null;
let _productService = null;
let _reportEngine = null;

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

    // 初始化报表模板
    _reportEngine = new ReportEngine(_repository, _excelDAO);
    _reportEngine.initializeTemplates();
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
    _productService.refreshProduct();

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
    _productService.refreshProduct();

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
    _productService.refreshProduct();

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
    _productService.refreshProduct();

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
    _productService.refreshProduct();

    const results = _productService.updateAll();
    const updateReport = _productService.generateUpdateReport(results);

    MsgBox(updateReport, 64, "一键更新");
  } catch (err) {
    MsgBox(`一键更新失败：${err.message}`, 16, "错误");
  }
}

// 从UI获取筛选条件
function _buildQueryFromUI() {
  const query = {};

  // 主销季节
  const seasons = [];
  if (UserForm1.CheckBox37?.Value) seasons.push("春秋");
  if (UserForm1.CheckBox34?.Value) seasons.push("夏");
  if (UserForm1.CheckBox35?.Value) seasons.push("冬");
  if (UserForm1.CheckBox36?.Value) seasons.push("四季");
  if (seasons.length > 0)
    Object.assign(query, { mainSalesSeason: { $in: seasons } });

  // 适用性别
  const genders = [];
  if (UserForm1.CheckBox42?.Value) genders.push("男童");
  if (UserForm1.CheckBox43?.Value) genders.push("女童");
  if (UserForm1.CheckBox44?.Value) genders.push("中性");
  if (genders.length > 0)
    Object.assign(query, { applicableGender: { $in: genders } });

  // 商品状态
  const statuses = [];
  if (UserForm1.CheckBox45?.Value) statuses.push("商品上线");
  if (UserForm1.CheckBox46?.Value) statuses.push("部分上线");
  if (UserForm1.CheckBox47?.Value) statuses.push("商品下线");
  if (statuses.length > 0)
    Object.assign(query, { itemStatus: { $in: statuses } });

  // 下线原因
  const reasons = [];
  if (UserForm1.CheckBox54?.Value)
    reasons.push("新品下架", "过季下架", "更换吊牌", "转移品牌", "清仓淘汰");
  if (UserForm1.CheckBox55?.Value)
    reasons.push("内网撞款", "资质问题", "内在质检");
  if (UserForm1.CheckBox56?.Value) reasons.push(undefined);
  if (reasons.length > 0)
    Object.assign(query, { offlineReason: { $in: reasons } });

  // 售龄
  const ageRange = [];
  if (isFinite(UserForm1.TextEdit31?.Value))
    ageRange[0] = Number(UserForm1.TextEdit31?.Value);
  if (isFinite(UserForm1.TextEdit32?.Value))
    ageRange[1] = Number(UserForm1.TextEdit32?.Value);

  if (ageRange.length === 1) {
    Object.assign(query, {
      salesAge: { $gte: ageRange[0] },
    });
  }

  if (ageRange.length === 2) {
    if (ageRange[0]) {
      Object.assign(query, {
        salesAge: { $between: [ageRange[0], ageRange[1]] },
      });
    } else {
      Object.assign(query, {
        salesAge: { $lte: ageRange[1] },
      });
    }
  }

  // 营销定位
  const positions = [];
  if (UserForm1.CheckBox48?.Value) positions.push("引流款");
  if (UserForm1.CheckBox49?.Value) positions.push("利润款");
  if (UserForm1.CheckBox50?.Value) positions.push("清仓款");
  if (positions.length > 0)
    Object.assign(query, { marketingPositioning: { $in: positions } });

  // 活动状态
  const activities = [];
  if (UserForm1.OptionButton6?.Value) {
    activities.push(
      "直通车",
      "黄金促",
      "黄金限量",
      "白金促",
      "TOP3",
      "白金限量",
    );
  }
  if (UserForm1.OptionButton5?.Value) activities.push(undefined);
  if (activities.length > 0)
    Object.assign(query, { activityStatus: { $in: activities } });

  // 中台
  if (UserForm1.CheckBox40?.Value)
    Object.assign(query, { userOperations1: { $gt: 0 } });
  if (UserForm1.CheckBox41?.Value)
    Object.assign(query, { userOperations2: { $gt: 0 } });

  // 备货模式
  const stockModes = [];
  if (UserForm1.CheckBox23?.Value) stockModes.push("现货");
  if (UserForm1.CheckBox24?.Value) stockModes.push("通版通货");
  if (UserForm1.CheckBox25?.Value) stockModes.push("专版通货");
  if (stockModes.length > 0) query.stockingMode = stockModes;

  // 数值范围
  if (UserForm1.TextEdit1?.Value || UserForm1.TextEdit11?.Value) {
    query.salesAge = [
      UserForm1.TextEdit1?.Value
        ? Number(UserForm1.TextEdit1.Value)
        : undefined,
      UserForm1.TextEdit11?.Value
        ? Number(UserForm1.TextEdit11.Value)
        : undefined,
    ];
  }

  if (UserForm1.TextEdit3?.Value || UserForm1.TextEdit4?.Value) {
    query.profit = [
      UserForm1.TextEdit3?.Value
        ? Number(UserForm1.TextEdit3.Value)
        : undefined,
      UserForm1.TextEdit4?.Value
        ? Number(UserForm1.TextEdit4.Value)
        : undefined,
    ];
  }

  if (UserForm1.TextEdit7?.Value || UserForm1.TextEdit8?.Value) {
    query.sellableInventory = [
      UserForm1.TextEdit7?.Value
        ? Number(UserForm1.TextEdit7.Value)
        : undefined,
      UserForm1.TextEdit8?.Value
        ? Number(UserForm1.TextEdit8.Value)
        : undefined,
    ];
  }

  if (UserForm1.TextEdit14?.Value || UserForm1.TextEdit15?.Value) {
    query.salesQuantityOfLast7Days = [
      UserForm1.TextEdit14?.Value
        ? Number(UserForm1.TextEdit14.Value)
        : undefined,
      UserForm1.TextEdit15?.Value
        ? Number(UserForm1.TextEdit15.Value)
        : undefined,
    ];
  }

  return query;
}

// 从UI获取排序条件
function _bulidSortFromUI() {}

// 生成报表
function UserForm1_CommandButton13_Click() {
  _reportEngine.setCurrentTemplate("库存预警报表");
  _reportEngine.generateReport();
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
