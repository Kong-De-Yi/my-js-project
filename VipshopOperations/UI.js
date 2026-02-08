// 显示窗口
function Macro() {
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

//更新常态商品
function UserForm1_CommandButton2_Click() {
  try {
    SystemRecord.initializeData();
    VipshopGoods.initializeData();
    Main.updateRegularProduct();
  } catch (err) {
    MsgBox(err.message);
    if (err instanceof CustomError) {
      let wb = Workbooks.Add();
      DAO.updateWorksheet("Sheet1", err.data, err.keyToTitle, wb);
    }
    return;
  }
  VipshopGoods.saveVipshopGoods();
  MsgBox("【" + RegularProduct.getWsName() + "】更新成功！");
}

//更新商品价格
function UserForm1_CommandButton1_Click() {
  try {
    SystemRecord.initializeData();
    VipshopGoods.initializeData();
    Main.updateProductPrice();
  } catch (err) {
    MsgBox(err.message);
    if (err instanceof CustomError) {
      let wb = Workbooks.Add();
      DAO.updateWorksheet("Sheet1", err.data, err.keyToTitle, wb);
    }
    return;
  }
  VipshopGoods.saveVipshopGoods();
  MsgBox("【" + ProductPrice.getWsName() + "】更新成功！");
}

//更新商品库存
function UserForm1_CommandButton4_Click() {
  try {
    SystemRecord.initializeData();
    VipshopGoods.initializeData();
    Main.updateInventory();
  } catch (err) {
    MsgBox(err.message);
    if (err instanceof CustomError) {
      let wb = Workbooks.Add();
      DAO.updateWorksheet("Sheet1", err.data, err.keyToTitle, wb);
    }
    return;
  }
  VipshopGoods.saveVipshopGoods();
  MsgBox("【" + Inventory.getWsName() + "】更新成功！");
}

//更新商品销售
function UserForm1_CommandButton5_Click() {
  try {
    SystemRecord.initializeData();
    VipshopGoods.initializeData();
    Main.updateProductSales();
  } catch (err) {
    MsgBox(err.message);
    if (err instanceof CustomError) {
      let wb = Workbooks.Add();
      DAO.updateWorksheet("Sheet1", err.data, err.keyToTitle, wb);
    }
    return;
  }
  VipshopGoods.saveVipshopGoods();
  MsgBox("【" + ProductSales.getWsName() + "】更新成功！");
}

//一键更新
function UserForm1_CommandButton6_Click() {
  try {
    SystemRecord.initializeData();
    VipshopGoods.initializeData();
    Main.updateRegularProduct();
    Main.updateProductPrice();
    Main.updateInventory();
    Main.updateProductSales();
  } catch (err) {
    MsgBox(err.message);
    if (err instanceof CustomError) {
      let wb = Workbooks.Add();
      DAO.updateWorksheet("Sheet1", err.data, err.keyToTitle, wb);
    }
    return;
  }
  VipshopGoods.saveVipshopGoods();
  MsgBox("一键更新成功！");
}

//报表输出
function UserForm1_CommandButton13_Click() {
  try {
    SystemRecord.initializeData();
    Main.outputReport();
  } catch (err) {
    MsgBox(err.message);
    if (err instanceof CustomError) {
      let wb = Workbooks.Add();
      DAO.updateWorksheet("Sheet1", err.data, err.keyToTitle, wb);
    }
    return;
  }
  MsgBox("报表输出成功！");
}
//组合商品
function UserForm1_CheckBox11_Click() {
  if (UserForm1.CheckBox11.Value) {
    UserForm1.TextEdit19.Value = 1;
    UserForm1.TextEdit20.Value = "";
  } else {
    UserForm1.TextEdit19.Value = "";
    UserForm1.TextEdit20.Value = "";
  }
}

//平台活动提报
function UserForm1_CommandButton9_Click() {
  try {
    SystemRecord.initializeData();
    Main.signUpActivity();
  } catch (err) {
    MsgBox(err.message);
    if (err instanceof CustomError) {
      let wb = Workbooks.Add();
      DAO.updateWorksheet("Sheet1", err.data, err.keyToTitle, wb);
    }
    return;
  }
  MsgBox("平台活动导入表输出成功！");
}
