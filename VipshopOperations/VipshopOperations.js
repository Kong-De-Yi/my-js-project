class Main {
  static updateRegularProduct() {
    try {
      RegularProduct.initializeData();
    } catch (err) {
      throw err;
    }

    //向货号总表添加新增的常态货号
    let allRegularProducts = RegularProduct.filterRegularProducts({
      brandSN: VipshopGoods.getBrandSN(),
    });
    allRegularProducts.forEach((item) => {
      let findItem = VipshopGoods.findVipshopGoods({
        itemNumber: item.itemNumber,
      });

      if (!findItem) {
        let newItem = new VipshopGoods(item.itemNumber);
        newItem.marketingPositioning = "利润款";
        VipshopGoods.addVipshopGoods(newItem);
      }
    });

    let allVipshopGoods = VipshopGoods.filterVipshopGoods();
    allVipshopGoods.forEach((item) => {
      let findItem = RegularProduct.findRegularProduct({
        itemNumber: item.itemNumber,
      });
      if (findItem) {
        item.thirdLevelCategory = findItem.thirdLevelCategory;
        item.P_SPU = findItem.P_SPU;
        item.MID = findItem.MID;
        item.styleNumber = findItem.styleNumber;
        item.color = findItem.color;
        item.itemStatus = findItem.itemStatus;
        item.vipshopPrice = Number(findItem.vipshopPrice);
        item.finalPrice = Number(findItem.finalPrice);
        item.sellableDays = findItem.sellableDays;
        //商品上架则清空下线原因
        if (item.itemStatus != "商品下线") {
          item.offlineReason = "";
        }
      }

      let products = RegularProduct.filterRegularProducts({
        itemNumber: item.itemNumber,
      }).sort(RegularProduct.compareBySize);

      //可售库存
      item.sellableInventory = products.reduce(
        (sum, current) => sum + Number(current.sellableInventory),
        0,
      );

      //是否断码
      item.isOutOfStock = products
        .reduce((sizes, current) => {
          if (
            current.sizeStatus == "尺码上线" &&
            current.sellableInventory == 0
          ) {
            sizes.push(current.size);
          }
          return sizes;
        }, [])
        .join("/");
    });

    //清空常态商品
    RegularProduct.clear();

    //更新系统记录
    SystemRecord.getSystemRecord().updateDateOfRegularProduct =
      new Date().toString();
    SystemRecord.updateSystemRecord();
  }

  static updateProductPrice() {
    try {
      ProductPrice.initializeData();
      ProductPrice.checkData();
    } catch (err) {
      throw err;
    }

    let allVipshopGoods = VipshopGoods.filterVipshopGoods();
    allVipshopGoods.forEach((item) => {
      let findItem = ProductPrice.findProductPrice({
        itemNumber: item.itemNumber,
      });
      if (findItem) {
        item.designNumber = findItem.designNumber;
        item.picture = findItem.picture;
        item.costPrice = Number(findItem.costPrice);
        item.lowestPrice = Number(findItem.lowestPrice);
        item.silverPrice = Number(findItem.silverPrice);
        item.userOperations1 = Number(findItem.userOperations1 ?? 0);
        item.userOperations2 = Number(findItem.userOperations2 ?? 0);
      }
    });

    //更新系统记录
    SystemRecord.getSystemRecord().updateDateOfProductPrice =
      new Date().toString();
    SystemRecord.updateSystemRecord();
  }

  static updateInventory() {
    try {
      ComboProduct.initializeData();
      Inventory.initializeData();
    } catch (err) {
      throw err;
    }

    let allVipshopGoods = VipshopGoods.filterVipshopGoods();

    allVipshopGoods.forEach((item) => {
      //成品库存清0
      item.finishedGoodsMainInventory = 0;
      item.finishedGoodsIncomingInventory = 0;
      item.finishedGoodsFinishingInventory = 0;
      item.finishedGoodsOversoldInventory = 0;
      item.finishedGoodsPrepareInventory = 0;
      item.finishedGoodsReturnInventory = 0;
      item.finishedGoodsPurchaseInventory = 0;
      item.finishedGoodsTotalInventory = 0;
      //通货库存清0
      item.generalGoodsMainInventory = 0;
      item.generalGoodsIncomingInventory = 0;
      item.generalGoodsFinishingInventory = 0;
      item.generalGoodsOversoldInventory = 0;
      item.generalGoodsPrepareInventory = 0;
      item.generalGoodsReturnInventory = 0;
      item.generalGoodsPurchaseInventory = 0;
      item.generalGoodsTotalInventory = 0;
      //合计库存清0
      item.totalInventory = 0;

      //找出常态商品中货号对应的条码商品
      let products = RegularProduct.filterRegularProducts({
        itemNumber: item.itemNumber,
      });

      products.forEach((productItem) => {
        //查找条码商品对应的库存
        let findProuctInventory = Inventory.findInventory({
          productCode: productItem.productCode,
        });

        if (findProuctInventory) {
          item.finishedGoodsMainInventory += +findProuctInventory.mainInventory;
          item.finishedGoodsIncomingInventory +=
            +findProuctInventory.incomingInventory;
          item.finishedGoodsFinishingInventory +=
            +findProuctInventory.finishingInventory;
          this.finishedGoodsOversoldInventory +=
            +findProuctInventory.oversoldInventory;
          this.finishedGoodsPrepareInventory +=
            +findProuctInventory.prepareInventory;
          this.finishedGoodsReturnInventory +=
            +findProuctInventory.returnInventory;
          this.finishedGoodsPurchaseInventory +=
            +findProuctInventory.purchaseInventory;
        }
        //成品合计
        item.finishedGoodsTotalInventory =
          item.finishedGoodsMainInventory +
          item.finishedGoodsIncomingInventory +
          item.finishedGoodsFinishingInventory +
          item.finishedGoodsOversoldInventory +
          item.finishedGoodsPrepareInventory +
          item.finishedGoodsReturnInventory +
          item.finishedGoodsPurchaseInventory;

        //通货库存计算
        let findComboProducts = ComboProduct.filterComboProducts({
          productCode: productItem.productCode,
        });

        if (findComboProducts.length != 0) {
          findComboProducts.forEach((comboProduct) => {
            let findSubProductInventory = Inventory.findInventory({
              productCode: comboProduct.subProductCode,
            });

            if (findSubProductInventory) {
              item.generalGoodsMainInventory +=
                +findSubProductInventory.mainInventory /
                +comboProduct.subProductQuantity;
              item.generalGoodsIncomingInventory +=
                +findSubProductInventory.incomingInventory /
                +comboProduct.subProductQuantity;
              item.generalGoodsFinishingInventory +=
                +findSubProductInventory.finishingInventory /
                +comboProduct.subProductQuantity;
              item.generalGoodsOversoldInventory +=
                +findSubProductInventory.oversoldInventory /
                +comboProduct.subProductQuantity;
              item.generalGoodsPrepareInventory +=
                +findSubProductInventory.prepareInventory /
                +comboProduct.subProductQuantity;
              item.generalGoodsReturnInventory +=
                +findSubProductInventory.returnInventory /
                +comboProduct.subProductQuantity;
              item.generalGoodsPurchaseInventory +=
                +findSubProductInventory.purchaseInventory /
                +comboProduct.subProductQuantity;
            }
          });
        }
        //通货合计
        item.generalGoodsTotalInventory =
          item.generalGoodsMainInventory +
          item.generalGoodsIncomingInventory +
          item.generalGoodsFinishingInventory +
          item.generalGoodsOversoldInventory +
          item.generalGoodsPrepareInventory +
          item.generalGoodsReturnInventory +
          item.generalGoodsPurchaseInventory;
      });
      //合计库存
      item.totalInventory =
        item.finishedGoodsTotalInventory + item.generalGoodsTotalInventory;
    });

    //清空组合商品和商品库存
    ComboProduct.clear();
    Inventory.clear();

    //更新系统记录
    SystemRecord.getSystemRecord().updateDateOfInventory =
      new Date().toString();
    SystemRecord.updateSystemRecord();
  }

  static updateProductSales() {
    try {
      ProductSales.initializeData();
    } catch (err) {
      throw err;
    }

    let dateOfLast7Days = Utility.generateStringOfLast7Days();

    let updateDateOfLast7Days =
      SystemRecord.getSystemRecord().updateDateOfLast7Days;
    let needUpdateSystemRecordOfLast7Days = false;
    let needUpdateSystemRecordOfProductSales = false;

    let allVipshopGoods = VipshopGoods.filterVipshopGoods();

    allVipshopGoods.forEach((item) => {
      //近7天数据失效清0
      if (updateDateOfLast7Days != dateOfLast7Days) {
        item.exposureUVOfLast7Days = 0;
        item.productDetailsUVOfLast7Days = 0;
        item.addToCartUVOfLast7Days = 0;
        item.customerCountOfLast7Days = 0;
        item.rejectAndReturnCountOfLast7Days = 0;
        item.salesQuantityOfLast7Days = 0;
        item.salesAmountOfLast7Days = 0;
        item.styleSalesOfLast7Days = 0;
      }

      let findItem = ProductSales.findProductSales({
        itemNumber: item.itemNumber,
        salesDate: dateOfLast7Days,
      });

      if (findItem) {
        //需要更新7天数据更新日期
        needUpdateSystemRecordOfLast7Days = true;
        SystemRecord.getSystemRecord().updateDateOfLast7Days = dateOfLast7Days;

        item.exposureUVOfLast7Days = +findItem.exposureUV;
        item.productDetailsUVOfLast7Days = +findItem.productDetailsUV;
        item.addToCartUVOfLast7Days = +findItem.addToCartUV;
        item.customerCountOfLast7Days = +findItem.customerCount;
        item.rejectAndReturnCountOfLast7Days = +findItem.rejectAndReturnCount;
        item.salesQuantityOfLast7Days = +findItem.salesQuantity;
        item.salesAmountOfLast7Days = +findItem.salesAmount;

        item.firstListingTime = findItem.firstListingTime
          ? "'" + findItem.firstListingTime
          : "";
      }

      item.unitPriceOfLast7Days = !item.salesQuantityOfLast7Days
        ? ""
        : item.salesAmountOfLast7Days / item.salesQuantityOfLast7Days;

      item.clickThroughRateOfLast7Days = !item.exposureUVOfLast7Days
        ? ""
        : item.productDetailsUVOfLast7Days / item.exposureUVOfLast7Days;

      item.addToCartRateOfLast7Days = !item.productDetailsUVOfLast7Days
        ? ""
        : item.addToCartUVOfLast7Days / item.productDetailsUVOfLast7Days;

      item.purchaseRateOfLast7Days = !item.productDetailsUVOfLast7Days
        ? ""
        : item.customerCountOfLast7Days / item.productDetailsUVOfLast7Days;

      item.rejectAndReturnRateOfLast7Days = !item.salesQuantityOfLast7Days
        ? ""
        : item.rejectAndReturnCountOfLast7Days / item.salesQuantityOfLast7Days;

      //更新年，月，日销量
      for (let prop of Object.keys(VipshopGoods.getOptionalKeyToTitle())) {
        let findItem = ProductSales.findProductSales({
          itemNumber: item.itemNumber,
          salesDate: prop.replace(/^\+/, ""),
        });

        if (findItem) {
          item[prop] = findItem.salesQuantity;
          item.firstListingTime = findItem.firstListingTime
            ? "'" + findItem.firstListingTime
            : "";
          //需要更新销量总计
          item.totalSales = "expired";
          //需要更新商品销售系统记录
          if ("+" + findItem.salesDate == Utility.generateStringOfYesterday()) {
            needUpdateSystemRecordOfProductSales = true;
            SystemRecord.getSystemRecord().updateDateOfProductSales =
              new Date().toString();
          }
        }
      }
    });

    //根据需要计算近7天款销量
    let styleSalesOfLast7DaysMap = new Map();
    if (needUpdateSystemRecordOfLast7Days) {
      allVipshopGoods.forEach((item) => {
        let styleSalesOfLast7Days = styleSalesOfLast7DaysMap.get(
          item.styleNumber,
        );

        if (!styleSalesOfLast7Days) {
          styleSalesOfLast7DaysMap.set(
            item.styleNumber,
            item.salesQuantityOfLast7Days,
          );
        } else {
          styleSalesOfLast7DaysMap.set(
            item.styleNumber,
            styleSalesOfLast7Days + item.salesQuantityOfLast7Days,
          );
        }
      });
    }
    //更新近7天款销量和销量合计
    allVipshopGoods.forEach((item) => {
      if (needUpdateSystemRecordOfLast7Days) {
        item.styleSalesOfLast7Days =
          styleSalesOfLast7DaysMap.get(item.styleNumber) ?? 0;
      }
      if (item.totalSales == "expired") {
        item.totalSales = 0;
        for (let prop of Object.keys(
          Utility.generateDateKeyToTitleForTotalSales(),
        )) {
          if (item[prop]) {
            item.totalSales += +item[prop];
          }
        }
      }
    });

    if (
      needUpdateSystemRecordOfLast7Days ||
      needUpdateSystemRecordOfProductSales
    ) {
      SystemRecord.updateSystemRecord();
    }
    //清空商品销售表
    ProductSales.clear();
  }

  static outputReport() {
    //检查数据更新情况
    let systemRecord = SystemRecord.getSystemRecord();
    if (
      systemRecord.updateDateOfLast7Days != Utility.generateStringOfLast7Days()
    ) {
      if (
        MsgBox(
          "【近7天商品销售数据】尚未更新,是否继续生成报表?",
          jsYesNo,
          "提醒",
        ) === jsResultNo
      )
        throw new Error("请更新【近7天商品销售数据】后再重试!");
    }

    if (
      !Utility.isToday(
        new Date(Date.parse(systemRecord.updateDateOfProductPrice)),
      )
    ) {
      if (
        MsgBox(
          "【商品价格】今日尚未更新,是否继续生成报表?",
          jsYesNo,
          "提醒",
        ) === jsResultNo
      )
        throw new Error("请更新【商品价格】后再重试!");
    }

    if (
      !Utility.isToday(
        new Date(Date.parse(systemRecord.updateDateOfRegularProduct)),
      )
    ) {
      if (
        MsgBox(
          "【常态商品】今日尚未更新,是否继续生成报表?",
          jsYesNo,
          "提醒",
        ) === jsResultNo
      )
        throw new Error("请更新【常态商品】后再重试!");
    }

    if (
      !Utility.isToday(new Date(Date.parse(systemRecord.updateDateOfInventory)))
    ) {
      if (
        MsgBox(
          "【商品库存】今日尚未更新,是否继续生成报表?",
          jsYesNo,
          "提醒",
        ) === jsResultNo
      )
        throw new Error("请更新【商品库存】后再重试!");
    }

    if (
      !Utility.isToday(
        new Date(Date.parse(systemRecord.updateDateOfProductSales)),
      )
    ) {
      if (
        MsgBox(
          "【商品销售】昨日数据尚未更新,是否继续生成报表?",
          jsYesNo,
          "提醒",
        ) === jsResultNo
      )
        throw new Error("请更新【商品销售】昨日数据后再重试!");
    }

    //获取货品筛选选项
    let selectOption = Utility.getSelectOption();

    //拆分工作表选项
    let splitByOption = new Map([
      ["", "brandSN"],
      ["上市年份", "listingYear"],
      ["四级品类", "fourthLevelCategory"],
      ["运营分类", "operationClassification"],
      ["下线原因", "offlineReason"],
      ["三级品类", "thirdLevelCategory"],
    ]);
    //排序选项
    let sortOption = new Map([
      ["首次上架时间#true", VipshopGoods.compareByFirstListingTime],
      ["首次上架时间#false", VipshopGoods.compareByFirstListingTimeDesc],
      ["成本价#true", VipshopGoods.compareByCostPrice],
      ["成本价#false", VipshopGoods.compareByCostPriceDesc],
      ["白金价#true", VipshopGoods.compareBySilverPrice],
      ["白金价#false", VipshopGoods.compareBySilverPriceDesc],
      ["利润#true", VipshopGoods.compareByProfit],
      ["利润#false", VipshopGoods.compareByProfitDesc],
      ["利润率#true", VipshopGoods.compareByProfitRate],
      ["利润率#false", VipshopGoods.compareByProfitRateDesc],
      ["近7天件单价#true", VipshopGoods.compareByUnitPriceOfLast7Days],
      ["近7天件单价#false", VipshopGoods.compareByUnitPriceOfLast7DaysDesc],
      ["近7天曝光UV#true", VipshopGoods.compareByExposureUVOfLast7Days],
      ["近7天曝光UV#false", VipshopGoods.compareByExposureUVOfLast7DaysDesc],
      ["近7天商详UV#true", VipshopGoods.compareByProductDetailsUVOfLast7Days],
      [
        "近7天商详UV#false",
        VipshopGoods.compareByProductDetailsUVOfLast7DaysDesc,
      ],
      ["近7天加购UV#true", VipshopGoods.compareByAddToCartUVOfLast7Days],
      ["近7天加购UV#false", VipshopGoods.compareByAddToCartUVOfLast7DaysDesc],
      ["近7天客户数#true", VipshopGoods.compareByCustomerCountOfLast7Days],
      ["近7天客户数#false", VipshopGoods.compareByCustomerCountOfLast7DaysDesc],
      [
        "近7天拒退数#true",
        VipshopGoods.compareByRejectAndReturnCountOfLast7Days,
      ],
      [
        "近7天拒退数#false",
        VipshopGoods.compareByRejectAndReturnCountOfLast7DaysDesc,
      ],
      ["近7天销售量#true", VipshopGoods.compareBySalesQuantityOfLast7Days],
      ["近7天销售量#false", VipshopGoods.compareBySalesQuantityOfLast7DaysDesc],
      ["近7天销售额#true", VipshopGoods.compareBySalesAmountOfLast7Days],
      ["近7天销售额#false", VipshopGoods.compareBySalesAmountOfLast7DaysDesc],
      ["近7天点击率#true", VipshopGoods.compareByClickThroughRateOfLast7Days],
      [
        "近7天点击率#false",
        VipshopGoods.compareByClickThroughRateOfLast7DaysDesc,
      ],
      ["近7天加购率#true", VipshopGoods.compareByAddToCartRateOfLast7Days],
      ["近7天加购率#false", VipshopGoods.compareByAddToCartRateOfLast7DaysDesc],
      ["近7天转化率#true", VipshopGoods.compareByPurchaseRateOfLast7Days],
      ["近7天转化率#false", VipshopGoods.compareByPurchaseRateOfLast7DaysDesc],
      [
        "近7天拒退率#true",
        VipshopGoods.compareByRejectAndReturnRateOfLast7Days,
      ],
      [
        "近7天拒退率#false",
        VipshopGoods.compareByRejectAndReturnRateOfLast7DaysDesc,
      ],
      ["近7天款销量#true", VipshopGoods.compareByStyleSalesOfLast7Days],
      ["近7天款销量#false", VipshopGoods.compareByStyleSalesOfLast7DaysDesc],

      ["可售库存#true", VipshopGoods.compareBySellableInventory],
      ["可售库存#false", VipshopGoods.compareBySellableInventoryDesc],
      ["可售天数#true", VipshopGoods.compareBySellableDays],
      ["可售天数#false", VipshopGoods.compareBySellableDaysDesc],
      ["合计库存#true", VipshopGoods.compareByTotalInventory],
      ["合计库存#false", VipshopGoods.compareByTotalInventoryDesc],
      ["成品合计#true", VipshopGoods.compareByFinishedGoodsTotalInventory],
      ["成品合计#false", VipshopGoods.compareByFinishedGoodsTotalInventoryDesc],
      ["通货合计#true", VipshopGoods.compareByGeneralGoodsTotalInventory],
      ["通货合计#false", VipshopGoods.compareByGeneralGoodsTotalInventoryDesc],
      ["销量总计#true", VipshopGoods.compareByTotalSales],
      ["销量总计#false", VipshopGoods.compareByTotalSalesDesc],
    ]);

    let keyToTitle = VipshopGoods.getFullKeyToTitle();

    //加载货号总表数据
    VipshopGoods.initializeData();
    //根据筛选条件获取货号总表数据
    let allVipshopGoods =
      VipshopGoods.filterVipshopGoodsByMultiCondition(selectOption);

    let sortBy = sortOption.get(
      UserForm1.ComboBox6.Value + "#" + UserForm1.OptionButton26.Value,
    );
    allVipshopGoods.sort(sortBy);

    let splitBy = splitByOption.get(UserForm1.ComboBox4.Value);

    let splitByMap = new Map();

    allVipshopGoods.forEach((item) => {
      if (item[splitBy]) {
        let splitBySet = splitByMap.get(item[splitBy]);
        if (splitBySet) {
          if (!splitBySet.has(item.styleNumber)) {
            splitBySet.add(item.styleNumber);
          }
        } else {
          splitByMap.set(item[splitBy], new Set());
          splitByMap.get(item[splitBy]).add(item.styleNumber);
        }
      }
    });

    Workbooks(DAO._wbName).Sheets(VipshopGoods.getWsName()).Copy();
    let newWb = ActiveWorkbook;

    for (let entry of splitByMap) {
      let outputData = [];
      for (let value of entry[1]) {
        let outputItems = VipshopGoods.filterVipshopGoods({
          styleNumber: value,
        });

        outputData.push(...outputItems);
        outputData.push([]);
      }
      let worksheetCount = newWb.Worksheets.Count;
      newWb
        .Sheets(VipshopGoods.getWsName())
        .Copy(null, newWb.Worksheets(worksheetCount));
      let newSt = ActiveSheet;
      newSt.Name = String(entry[0]).replace(/\//g, "");

      DAO.updateWorksheet(
        String(entry[0]).replace(/\//g, ""),
        outputData,
        keyToTitle,
        newWb,
      );

      //隐藏非必要列
      if (UserForm1.CheckBox38.Value) {
        newSt.Columns("A:E").EntireColumn.Hidden = true;
        newSt.Columns("I:L").EntireColumn.Hidden = true;
        newSt.Columns("S:S").EntireColumn.Hidden = true;
        newSt.Columns("Z:AC").EntireColumn.Hidden = true;
        newSt.Columns("AE:AF").EntireColumn.Hidden = true;
        newSt.Columns("BA:BF").EntireColumn.Hidden = true;
        newSt.Columns("BI:BN").EntireColumn.Hidden = true;
        newSt.Columns("BQ:CE").EntireColumn.Hidden = true;
      }
      //冻结窗口
      if (UserForm1.CheckBox33.Value) {
        newSt.Activate();
        let win = ActiveWindow;
        win.FreezePanes = false;
        win.SplitRow = 1; // 冻结第1行
        win.SplitColumn = 34; // 冻结前34列（A-AH）
        win.FreezePanes = true;
      }
    }
  }

  static checkUpdateBeforeSignUp() {
    //常态商品过期强制更新
    const systemRecord = SystemRecord.getSystemRecord();

    // 1. 检查更新日期是否存在
    if (!systemRecord || !systemRecord.updateDateOfRegularProduct) {
      throw new Error("未找到常态商品更新日期");
    }

    const updateTimestamp = Date.parse(systemRecord.updateDateOfRegularProduct);
    const updateDate = new Date(updateTimestamp);
    const now = new Date();

    // 3. 计算时间差（小时）
    const diffMs = now - updateDate;
    const diffHours = diffMs / (1000 * 60 * 60);

    if (diffHours > 5) {
      throw new Error(
        "活动提报需要最新的常态商品数据,请导入最新的常态商品数据后重试！",
      );
    }

    if (
      systemRecord.updateDateOfLast7Days != Utility.generateStringOfLast7Days()
    ) {
      throw new Error(
        "活动提报需要最新的近7天商品销售数据,请导入最新的商品销售数据后重试!",
      );
    }

    if (
      !Utility.isToday(new Date(Date.parse(systemRecord.updateDateOfInventory)))
    ) {
      throw new Error(
        "活动提报需要最新的商品库存数据,请导入最新的商品库存数据后重试!",
      );
    }
  }

  static signUpActivity() {
    /*1.确保常态商品和近7天数据为最新
      10.检查是否破价
      3.默认只提报上架商品
      4.可单独筛选商品提报
      7.商品价格表中有价格信息，并且价格信息填写正常
      5.校验同款不同价
       2.验证利润
       8.白金价和到手价不一致问题

      6.检验同款不同券
 
  */

    //验证用户选择的活动等级
    if (!UserForm1.ComboBox3.Value) {
      throw new Error("请先选择活动价格等级");
    }

    //活动前检查数据是否过期
    this.checkUpdateBeforeSignUp();

    //加载货号总表数据
    VipshopGoods.initializeData();

    //获取需要提报的商品
    let selectOption = Utility.getSelectOption();
    if (Object.keys(selectOption).length === 0) {
      selectOption = { itemStatus: ["商品上线", "部分上线"] };
    }

    let requiredSignUpVipshopGoods =
      VipshopGoods.filterVipshopGoodsByMultiCondition(selectOption);

    if (requiredSignUpVipshopGoods.length == 0) {
      throw new Error("没有需要提报的商品");
    }

    //更新商品价格数据
    this.updateProductPrice();

    let allVipshopGoods = VipshopGoods.filterVipshopGoods();

    //检查所有商品是否都存在商品价格信息并验证是否已破价
    let abnormalPriceVipshopGoods = [];

    for (const item of allVipshopGoods) {
      if (!item.itemNumber) continue; //忽略无货号的商品

      let findItem = ProductPrice.findProductPrice({
        itemNumber: item.itemNumber,
      });
      if (findItem) {
        if (item.silverPrice < item.lowestPrice) {
          item.errReason = "商品白金价低于最低价，已破价！";
          abnormalPriceVipshopGoods.push(item);
        }
      } else {
        item.errReason = "商品价格中未找到该货号！";
        abnormalPriceVipshopGoods.push(item);
      }
    }

    if (abnormalPriceVipshopGoods.length != 0) {
      throw new CustomError(
        "请检查商品价格数据！",
        { itemNumber: "货号", errReason: "异常原因" },
        abnormalPriceVipshopGoods,
      );
    }

    //统计是否存在同款不同价
    let styleSilverPriceMap = new Map();
    for (const item of allVipshopGoods) {
      if (!item.itemNumber) continue; //忽略无货号的商品
      if (styleSilverPriceMap.has(item.styleNumber)) {
        let vipshopStyle = styleSilverPriceMap.get(item.styleNumber);
        vipshopStyle.vipshopGoodsCount++;
        vipshopStyle.totalSilverPrice += item.silverPrice;
        vipshopStyle.styleTotalSales += item.totalSales;

        let averageSilverPrice =
          vipshopStyle.totalSilverPrice / vipshopStyle.vipshopGoodsCount;

        if (item.silverPrice != averageSilverPrice) {
          vipshopStyle.warnMessage = "同款不同价";
        }
      } else {
        styleSilverPriceMap.set(item.styleNumber, {
          vipshopGoodsCount: 1,
          totalSilverPrice: item.silverPrice,
          totalSales: item.totalSales,
        });
      }
    }

    //校验货号维度
    for (const item of allVipshopGoods) {
      if (!item.itemNumber) continue; //忽略无货号的商品

      item.warnMessage = [];

      //检查同款不同价
      if (styleSilverPriceMap.get(item.styleNumber).warnMessage) {
        item.warnMessage.push("同款不同价");
      }

      //检查白金价是否小于到手价
      if (item.silverPrice < item.finalPrice) {
        item.warnMessage.push("白金价小于到手价，需核实到手价价格等级");
      }

      //验证白金价活动利润:1.新品毛利率50% 2.利润款毛利率>35%，毛利>5元 3. 引流款毛利率和毛利>0
      let silverPriceProfit = ProductPrice.calProfit(
        item.brandSN,
        item.costPrice,
        item.silverPrice,
        item.userOperations1,
        item.userOperations2,
        item.rejectAndReturnRateOfLast7Days,
      );
      let silverPriceProfitRate = Number(
        (silverPriceProfit / item.costPrice).toFixed(5),
      );

      //判断是否能正常计算利润
      if (!silverPriceProfit) {
        item.warnMessage.push("无法计算白金价毛利和毛利率");
      }

      //新品
      if (!item.salesAge || item.salesAge <= 15) {
        if (silverPriceProfit < 5) {
          item.warnMessage.push("售龄15天内的新品白金价毛利建议在5元以上");
        }
        if (silverPriceProfitRate < 0.5) {
          item.warnMessage.push("售龄15天内的新品白金价毛利率建议在50%以上");
        }
      } else {
        switch (item.marketingPositioning) {
          case "引流款":
            if (silverPriceProfitRate < 0.05) {
              item.warnMessage.push("引流款的白金价毛利率建议在5%以上");
            }
            break;

          case "清仓款":
            if (item.generalGoodsTotalInventory > 1) {
              item.warnMessage.push("清仓款请解绑组合装");
            }
            break;

          default:
            if (silverPriceProfit < 5) {
              item.warnMessage.push("利润款的白金价毛利建议在5元以上");
            }
            if (silverPriceProfitRate < 0.35) {
              item.warnMessage.push("利润款的白金价毛利率建议在35%以上");
            }
        }
      }

      //计算活动价格
      const productActivityPrice = ProductPrice.calProductActivityPrice(
        item.silverPrice,
      );

      switch (UserForm1.ComboBox3.Value) {
        case "直通车":
          item.activityLevel = "直通车";
          item.activityPrice = productActivityPrice.firstActPrice;
          break;
        case "黄金等级":
          item.activityLevel = "黄金等级";
          item.activityPrice = productActivityPrice.secondActPrice;
          break;
        case "TOP3":
          item.activityLevel = "TOP3";
          item.activityPrice = productActivityPrice.thirdActPrice;
          break;
        case "白金等级":
          item.activityLevel = "白金等级";
          item.activityPrice = productActivityPrice.fourthActPrice;
          break;
        case "白金限量":
          item.activityLevel = "白金限量";
          item.isLimited = "是";
          item.limitedCount = 500;
          item.limitedCountForUser = 10;
          item.canPromote = "是";
          item.activityPrice = productActivityPrice.fifthActPrice;
      }

      item.activityProfit = ProductPrice.calProfit(
        item.brandSN,
        item.costPrice,
        item.activityPrice,
        item.userOperations1,
        item.userOperations2,
        item.rejectAndReturnRateOfLast7Days,
      );
    }
    item.activityProfitRate = Number(
      (item.activityProfit / item.costPrice).toFixed(5),
    );

    let keyToTitle = {
      activityLevel: "活动等级",
      activityPrice: "活动价",
      activityProfit: "活动毛利",
      activityProfitRate: "活动毛利率",
      salesAge: "售龄",
      isOutOfStock: "是否断码",
      finishedGoodsTotalInventory: "成品库存",
      generalGoodsTotalInventory: "通货库存",
      clickThroughRateOfLast7Days: "近7天点击率",
      addToCartRateOfLast7Days: "近7天加购率",
      purchaseRateOfLast7Days: "近7天转化率",
    };
    let activityKeyToTitle = {};
    let outputData = [];

    if (UserForm1.ComboBox3.Value == "白金限量") {
      for (let item of requiredSignUpVipshopGoods) {
        let outputItem = VipshopGoods.findVipshopGoods({
          styleNumber: item.styleNumber,
        });
        let warnMessage = [];
        let findItems = VipshopGoods.filterVipshopGoods({
          styleNumber: item.styleNumber,
        });
        for (let item of findItems) {
          if (item.warnMessage.length != 0) {
            for (let value of item.warnMessage) {
              if (!warnMessage.includes(value) && value) {
                warnMessage.push(value);
              }
            }
          }
        }
        if (warnMessage.length != 0) {
          outputItem.styleWarnMessage = warnMessage.join("/");
        }

        let styleTotalSales = styleSilverPriceMap.get(
          outputItem.styleNumber,
        ).styleTotalSales;

        if (styleTotalSales) {
          outputItem.styleTotalSales = styleTotalSales;
        }
        if (!outputData.find((item) => item.P_SPU == outputItem.P_SPU)) {
          outputData.push(outputItem);
        }
      }

      activityKeyToTitle = {
        styleTotalSales: "销量总计",
        styleWarnMessage: "活动警告",
        signUpNumber: "报名编号",
        P_SPU: "SPU_ID",
        sizeNumber: "尺码ID",
        activityPrice: "快抢价",
        isLimited: "是否限量",
        limitedCount: "限量件数",
        limitedCountForUser: "每用户限购数",
        canPromote: "是否可外推",
        applyReason: "申请理由",
      };
    }

    Object.assign(keyToTitle, activityKeyToTitle);

    Workbooks(DAO._wbName).Sheets(VipshopGoods.getWsName()).Copy();
    let newWb = ActiveWorkbook;

    newWb.Worksheets.Add();
    let newSt = ActiveSheet;
    newSt.Name = "活动提报";

    DAO.updateWorksheet("活动提报", outputData, keyToTitle, newWb);
  }
}

//货号总表
class VipshopGoods {
  static _brandSN = "10000708";
  static _wsName = "货号总表";

  static _keyToTitle = {
    listingYear: "上市年份",
    mainSalesSeason: "主销季节",
    applicableGender: "适用性别",
    fourthLevelCategory: "四级品类",
    operationClassification: "运营分类",
    stockingMode: "备货模式",
    offlineReason: "下线原因",
    marketingPositioning: "营销定位",
    marketingMemorandum: "营销备忘录",
    thirdLevelCategory: "三级品类",
    P_SPU: "P_SPU",
    MID: "MID",
    designNumber: "设计号",
    styleNumber: "款号",
    itemNumber: "货号",
    picture: "图片",
    color: "颜色",
    link: "链接",
    firstListingTime: "首次上架时间",
    salesAge: "售龄",
    itemStatus: "商品状态",
    activityStatus: "活动状态",
    isPriceBroken: "是否破价",
    costPrice: "成本价",
    lowestPrice: "最低价",
    silverPrice: "白金价",
    userOperations1: "中台1",
    userOperations2: "中台2",
    vipshopPrice: "唯品价",
    finalPrice: "到手价",
    firstOrderPrice: "首单价",
    superVipPrice: "超V价",
    profit: "利润",
    profitRate: "利润率",
    unitPriceOfLast7Days: "近7天件单价",
    exposureUVOfLast7Days: "近7天曝光UV",
    productDetailsUVOfLast7Days: "近7天商详UV",
    addToCartUVOfLast7Days: "近7天加购UV",
    customerCountOfLast7Days: "近7天客户数",
    rejectAndReturnCountOfLast7Days: "近7天拒退数",
    salesQuantityOfLast7Days: "近7天销售量",
    salesAmountOfLast7Days: "近7天销售额",
    clickThroughRateOfLast7Days: "近7天点击率",
    addToCartRateOfLast7Days: "近7天加购率",
    purchaseRateOfLast7Days: "近7天转化率",
    rejectAndReturnRateOfLast7Days: "近7天拒退率",
    styleSalesOfLast7Days: "近7天款销量",
    sellableInventory: "可售库存",
    sellableDays: "可售天数",
    isOutOfStock: "是否断码",
    totalInventory: "合计库存",
    finishedGoodsTotalInventory: "成品合计",
    finishedGoodsMainInventory: "成品主仓",
    finishedGoodsIncomingInventory: "成品进货",
    finishedGoodsFinishingInventory: "成品后整",
    finishedGoodsOversoldInventory: "成品超卖",
    finishedGoodsPrepareInventory: "成品备货",
    finishedGoodsReturnInventory: "成品销退",
    finishedGoodsPurchaseInventory: "成品在途",
    generalGoodsTotalInventory: "通货合计",
    generalGoodsMainInventory: "通货主仓",
    generalGoodsIncomingInventory: "通货进货",
    generalGoodsFinishingInventory: "通货后整",
    generalGoodsOversoldInventory: "通货超卖",
    generalGoodsPrepareInventory: "通货备货",
    generalGoodsReturnInventory: "通货销退",
    generalGoodsPurchaseInventory: "通货在途",
    totalSales: "销量总计",
  };

  static _optionalKeyToTitle = {};
  static _data = [];

  static initializeData() {
    this._optionalKeyToTitle = Utility.generateDateKeyToTitle();
    this._data = DAO.readWorksheet(
      this._wsName,
      this,
      this._keyToTitle,
      this._optionalKeyToTitle,
    );

    //货号去重
    let duplicates = Utility.findDuplicatesByProperty(this._data, "itemNumber");
    if (duplicates.length != 0) {
      throw new CustomError(
        this._wsName +
          "中存在重复的货号：【" +
          duplicates +
          "】，请核查后重试！",
        { itemNumber: "货号", errReason: "错误原因" },
        duplicates,
      );
    }
  }

  constructor(itemNumber) {
    this.itemNumber = itemNumber;
    this.brandSN = VipshopGoods._brandSN;
  }

  toString() {
    return this.itemNumber;
  }

  static getOptionalKeyToTitle() {
    return this._optionalKeyToTitle;
  }

  static getFullKeyToTitle() {
    return Object.assign({}, this._keyToTitle, this._optionalKeyToTitle);
  }

  static getBrandSN() {
    return this._brandSN;
  }
  static getWsName() {
    return this._wsName;
  }

  set listingYear(value) {
    value = value === undefined ? undefined : Number(value);

    if (
      [2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030].includes(value) ||
      value === undefined
    ) {
      this._listingYear = value;
    } else {
      throw new Error("货号总表中【上市年份】输入了非法值，请检查！");
    }
  }
  get listingYear() {
    return this._listingYear;
  }

  set mainSalesSeason(value) {
    value = value === undefined ? undefined : String(value);

    if (["春秋", "夏", "冬", "四季"].includes(value) || value === undefined) {
      this._mainSalesSeason = value;
    } else {
      throw new Error("货号总表中【主销季节】输入了非法值，请检查！");
    }
  }
  get mainSalesSeason() {
    return this._mainSalesSeason;
  }

  set applicableGender(value) {
    value = value === undefined ? undefined : String(value);

    if (["男童", "女童", "中性"].includes(value) || value === undefined) {
      this._applicableGender = value;
    } else {
      throw new Error("货号总表中【适用性别】输入了非法值，请检查！");
    }
  }
  get applicableGender() {
    return this._applicableGender;
  }

  set fourthLevelCategory(value) {
    this._fourthLevelCategory = value === undefined ? undefined : String(value);
  }
  get fourthLevelCategory() {
    return this._fourthLevelCategory;
  }

  set operationClassification(value) {
    this._operationClassification =
      value === undefined ? undefined : String(value);
  }
  get operationClassification() {
    return this._operationClassification;
  }

  set stockingMode(value) {
    value = value === undefined ? undefined : String(value);

    if (
      ["现货", "通版通货", "专版通货"].includes(value) ||
      value === undefined
    ) {
      this._stockingMode = value;
    } else {
      throw new Error("货号总表中【备货模式】输入了非法值，请检查！");
    }
  }
  get stockingMode() {
    return this._stockingMode;
  }

  set offlineReason(value) {
    value = value === undefined ? undefined : String(value);

    if (
      [
        "新品下架",
        "过季下架",
        "更换吊牌",
        "转移品牌",
        "清仓淘汰",
        "内网撞款",
        "资质问题",
        "内在质检",
      ].includes(value) ||
      value === undefined
    ) {
      this._offlineReason = value;
    } else {
      throw new Error("货号总表中【下线原因】输入了非法值，请检查！");
    }
  }
  get offlineReason() {
    return this._offlineReason;
  }

  set marketingPositioning(value) {
    value = value === undefined ? undefined : String(value);

    if (["引流款", "利润款", "清仓款"].includes(value) || value === undefined) {
      this._marketingPositioning = value;
    } else {
      throw new Error("货号总表中【营销定位】输入了非法值，请检查！");
    }
  }
  get marketingPositioning() {
    return this._marketingPositioning;
  }

  set marketingMemorandum(value) {
    this._marketingMemorandum = value === undefined ? undefined : String(value);
  }
  get marketingMemorandum() {
    return this._marketingMemorandum;
  }

  set thirdLevelCategory(value) {
    this._thirdLevelCategory = value === undefined ? undefined : String(value);
  }
  get thirdLevelCategory() {
    return this._thirdLevelCategory;
  }

  set P_SPU(value) {
    this._P_SPU = value === undefined ? undefined : String(value);
  }
  get P_SPU() {
    return this._P_SPU;
  }

  set MID(value) {
    this._MID = value === undefined ? undefined : String(value);
  }
  get MID() {
    return this._MID;
  }

  set designNumber(value) {
    this._designNumber = value === undefined ? undefined : String(value);
  }
  get designNumber() {
    return this._designNumber;
  }

  set styleNumber(value) {
    this._styleNumber = value === undefined ? undefined : String(value);
  }
  get styleNumber() {
    return this._styleNumber;
  }

  set itemNumber(value) {
    this._itemNumber = value === undefined ? undefined : String(value);
  }
  get itemNumber() {
    return this._itemNumber;
  }

  set color(value) {
    this._color = value === undefined ? undefined : String(value);
  }
  get color() {
    return this._color;
  }

  set link(value) {}
  get link() {
    return this._MID
      ? "https://detail.vip.com/detail-1234-" + this._MID + ".html"
      : undefined;
  }

  set firstListingTime(value) {
    this._firstListingTime = value === undefined ? undefined : String(value);
  }
  get firstListingTime() {
    return this._firstListingTime;
  }

  set salesAge(value) {}
  get salesAge() {
    let firstListingTime = Date.parse(this._firstListingTime);

    if (!isNaN(firstListingTime)) {
      return Math.floor(
        (Date.now() - firstListingTime) / (24 * 60 * 60 * 1000),
      );
    }

    return undefined;
  }

  set itemStatus(value) {
    this._itemStatus = value === undefined ? undefined : String(value);
  }
  get itemStatus() {
    return this._itemStatus;
  }

  set activityStatus(value) {}
  get activityStatus() {
    if (this._vipshopPrice && this._finalPrice) {
      return this._vipshopPrice > this._finalPrice ? "活动中" : "未提报";
    }
    return "(未知)";
  }

  set isPriceBroken(value) {}
  get isPriceBroken() {
    if (this._finalPrice && this._lowestPrice) {
      return this._lowestPrice > this._finalPrice ? "是" : "";
    }
    return "(未知)";
  }

  set costPrice(value) {
    if (value === undefined || typeof value === "boolean") {
      this._costPrice = undefined;
      return;
    }
    this._costPrice = isFinite(value) ? Number(value) : undefined;
  }
  get costPrice() {
    return this._costPrice;
  }

  get firstOrderPrice() {
    if (this.finalPrice) return +this.finalPrice - +(this.userOperations1 ?? 0);
  }
  set firstOrderPrice(value) {}

  get superVipPrice() {
    let vipDiscountRate = ProductPrice.getVipDiscountRate(this.brandSN);

    let vipDiscountAmount =
      this.finalPrice > 50
        ? Math.round(this.finalPrice * vipDiscountRate)
        : Number((this.finalPrice * vipDiscountRate).toFixed(1)); //超V优惠金额

    return (
      this.finalPrice -
      vipDiscountAmount -
      +(this.userOperations1 ?? 0) -
      +(this.userOperations2 ?? 0)
    );
  }
  set superVipPrice(value) {}

  get profit() {
    return ProductPrice.calProfit(
      this.brandSN,
      +this.costPrice,
      +this.finalPrice,
      +(this.userOperations1 ?? 0),
      +(this.userOperations2 ?? 0),
      this.rejectAndReturnRateOfLast7Days
        ? +this.rejectAndReturnRateOfLast7Days
        : undefined,
    );
  }
  set profit(value) {}

  get profitRate() {
    return ProductPrice.calProfitRate(
      this.brandSN,
      +this.costPrice,
      +this.finalPrice,
      +(this.userOperations1 ?? 0),
      +(this.userOperations2 ?? 0),
      this.rejectAndReturnRateOfLast7Days
        ? +this.rejectAndReturnRateOfLast7Days
        : undefined,
    );
  }
  set profitRate(value) {}

  //查找单个商品
  static findVipshopGoods(querys) {
    if (querys) {
      return this._data.find((item) => {
        for (let entry of Object.entries(querys)) {
          if (item[entry[0]] != entry[1] || !item[entry[0]]) {
            return false;
          }
        }
        return true;
      });
    }
    return undefined;
  }

  //查找符合条件的商品
  static filterVipshopGoods(querys) {
    let filterItems = this._data;
    if (querys) {
      filterItems = this._data.filter((item) => {
        for (let entry of Object.entries(querys)) {
          if (item[entry[0]] != entry[1] || !item[entry[0]]) {
            return false;
          }
        }
        return true;
      });
    }

    return filterItems;
  }
  //多条件查询
  static filterVipshopGoodsByMultiCondition(querys) {
    let VipShopGoodsForQuerys = this._data;
    for (let query of Object.entries(querys)) {
      if (!query[1]) continue;

      switch (query[0]) {
        case "mainSalesSeason":
        case "applicableGender":
        case "itemStatus":
        case "marketingPositioning":
        case "stockingMode":
          VipShopGoodsForQuerys = query[1].reduce((result, current) => {
            let filteredVipshopGoods = VipShopGoodsForQuerys.filter(
              (item) => item[query[0]] == current,
            );
            result.push(...filteredVipshopGoods);
            return result;
          }, []);
          break;

        case "offlineReason":
          VipShopGoodsForQuerys = query[1].reduce((result, current) => {
            let filteredVipshopGoods = VipShopGoodsForQuerys.filter(
              (item) =>
                item[query[0]] == current && item.itemStatus == "商品下线",
            );
            result.push(...filteredVipshopGoods);
            return result;
          }, []);
          break;

        case "userOperations1":
        case "userOperations2":
        case "isOutOfStock":
          VipShopGoodsForQuerys = VipShopGoodsForQuerys.filter(
            (item) => item[query[0]],
          );
          break;

        case "activityStatus":
          VipShopGoodsForQuerys = VipShopGoodsForQuerys.filter(
            (item) => item[query[0]] == query[1],
          );
          break;

        case "isPriceBroken":
          VipShopGoodsForQuerys = VipShopGoodsForQuerys.filter(
            (item) => item[query[0]] == "是",
          );
          break;

        case "salesAge":
        case "profit":
        case "profitRate":
        case "sellableInventory":
        case "sellableDays":
        case "finishedGoodsTotalInventory":
        case "generalGoodsTotalInventory":
        case "totalInventory":
        case "salesQuantityOfLast7Days":
          let start =
            query[0] == "profitRate" ? query[1][0] / 100 : query[1][0];
          let end = query[0] == "profitRate" ? query[1][1] / 100 : query[1][1];

          if (start) {
            if (end) {
              VipShopGoodsForQuerys = VipShopGoodsForQuerys.filter(
                (item) => +item[query[0]] >= +start && +item[query[0]] <= +end,
              );
            } else {
              VipShopGoodsForQuerys = VipShopGoodsForQuerys.filter(
                (item) => +item[query[0]] >= +start,
              );
            }
          } else {
            VipShopGoodsForQuerys = VipShopGoodsForQuerys.filter(
              (item) => +item[query[0]] <= +end,
            );
          }
          break;

        case "topProductsBySales":
          VipShopGoodsForQuerys.sort(
            VipshopGoods.compareBySalesQuantityOfLast7DaysDesc,
          );
          VipShopGoodsForQuerys = VipShopGoodsForQuerys.slice(0, query[1]);
      }
    }
    return VipShopGoodsForQuerys;
  }

  //添加新的商品
  static addVipshopGoods(...items) {
    this._data.push(...items);
  }

  //保存货号总表
  static saveVipshopGoods() {
    //默认按上架时间降序排序
    this._data.sort(VipshopGoods.compareByFirstListingTimeDesc);

    DAO.updateWorksheet(this._wsName, this._data, this.getFullKeyToTitle());
  }

  //首次上架时间升序
  static compareByFirstListingTime(VipshopGoodsA, VipshopGoodsB) {
    // 解析日期为当天开始时间的时间戳
    const parseDate = (dateStr) => {
      const timestamp = Date.parse(dateStr);
      if (isNaN(timestamp)) return 0;

      const date = new Date(timestamp);
      date.setHours(0, 0, 0, 0);
      return date.getTime();
    };

    const firstListingTimeA = parseDate(VipshopGoodsA.firstListingTime);
    const firstListingTimeB = parseDate(VipshopGoodsB.firstListingTime);

    // 先比较日期
    if (firstListingTimeA !== firstListingTimeB) {
      return firstListingTimeA - firstListingTimeB;
    }

    // 日期相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //首次上架时间降序
  static compareByFirstListingTimeDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByFirstListingTime(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //成本价升序
  static compareByCostPrice(VipshopGoodsA, VipshopGoodsB) {
    const costPriceA = Number(VipshopGoodsA.costPrice) || 0;
    const costPriceB = Number(VipshopGoodsB.costPrice) || 0;

    if (costPriceA !== costPriceB) {
      return costPriceA - costPriceB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //成本价降序
  static compareByCostPriceDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByCostPrice(VipshopGoodsA, VipshopGoodsB);
  }

  //白金价升序
  static compareBySilverPrice(VipshopGoodsA, VipshopGoodsB) {
    const silverPriceA = Number(VipshopGoodsA.silverPrice) || 0;
    const silverPriceB = Number(VipshopGoodsB.silverPrice) || 0;

    if (silverPriceA !== silverPriceB) {
      return silverPriceA - silverPriceB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //白金价降序
  static compareBySilverPriceDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareBySilverPrice(VipshopGoodsA, VipshopGoodsB);
  }

  //利润升序
  static compareByProfit(VipshopGoodsA, VipshopGoodsB) {
    const profitA = Number(VipshopGoodsA.profit) || 0;
    const profitB = Number(VipshopGoodsB.profit) || 0;

    if (profitA !== profitB) {
      return profitA - profitB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //利润降序
  static compareByProfitDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByProfit(VipshopGoodsA, VipshopGoodsB);
  }

  //利润率升序
  static compareByProfitRate(VipshopGoodsA, VipshopGoodsB) {
    const profitRateA = Number(VipshopGoodsA.profitRate) || 0;
    const profitRateB = Number(VipshopGoodsB.profitRate) || 0;

    if (profitRateA !== profitRateB) {
      return profitRateA - profitRateB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //利润率降序
  static compareByProfitRateDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByProfitRate(VipshopGoodsA, VipshopGoodsB);
  }

  //近7天件单价升序
  static compareByUnitPriceOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const unitPriceOfLast7DaysA =
      Number(VipshopGoodsA.unitPriceOfLast7Days) || 0;
    const unitPriceOfLast7DaysB =
      Number(VipshopGoodsB.unitPriceOfLast7Days) || 0;

    if (unitPriceOfLast7DaysA !== unitPriceOfLast7DaysB) {
      return unitPriceOfLast7DaysA - unitPriceOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天件单价降序
  static compareByUnitPriceOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByUnitPriceOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天曝光UV升序
  static compareByExposureUVOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const exposureA = Number(VipshopGoodsA.exposureUVOfLast7Days) || 0;
    const exposureB = Number(VipshopGoodsB.exposureUVOfLast7Days) || 0;

    if (exposureA !== exposureB) {
      return exposureA - exposureB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天曝光UV降序
  static compareByExposureUVOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByExposureUVOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天商详UV升序
  static compareByProductDetailsUVOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const productDetailsUVOfLast7DaysA =
      Number(VipshopGoodsA.productDetailsUVOfLast7Days) || 0;
    const productDetailsUVOfLast7DaysB =
      Number(VipshopGoodsB.productDetailsUVOfLast7Days) || 0;

    if (productDetailsUVOfLast7DaysA !== productDetailsUVOfLast7DaysB) {
      return productDetailsUVOfLast7DaysA - productDetailsUVOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天商详UV降序
  static compareByProductDetailsUVOfLast7DaysDesc(
    VipshopGoodsA,
    VipshopGoodsB,
  ) {
    return -VipshopGoods.compareByProductDetailsUVOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天加购UV升序
  static compareByAddToCartUVOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const addToCartUVOfLast7DaysA =
      Number(VipshopGoodsA.addToCartUVOfLast7Days) || 0;
    const addToCartUVOfLast7DaysB =
      Number(VipshopGoodsB.addToCartUVOfLast7Days) || 0;

    if (addToCartUVOfLast7DaysA !== addToCartUVOfLast7DaysB) {
      return addToCartUVOfLast7DaysA - addToCartUVOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天加购UV降序
  static compareByAddToCartUVOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByAddToCartUVOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天客户数升序
  static compareByCustomerCountOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const customerCountOfLast7DaysA =
      Number(VipshopGoodsA.customerCountOfLast7Days) || 0;
    const customerCountOfLast7DaysB =
      Number(VipshopGoodsB.customerCountOfLast7Days) || 0;

    if (customerCountOfLast7DaysA !== customerCountOfLast7DaysB) {
      return customerCountOfLast7DaysA - customerCountOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天客户数降序
  static compareByCustomerCountOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByCustomerCountOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天拒退数升序
  static compareByRejectAndReturnCountOfLast7Days(
    VipshopGoodsA,
    VipshopGoodsB,
  ) {
    const rejectAndReturnCountOfLast7DaysA =
      Number(VipshopGoodsA.rejectAndReturnCountOfLast7Days) || 0;
    const rejectAndReturnCountOfLast7DaysB =
      Number(VipshopGoodsB.rejectAndReturnCountOfLast7Days) || 0;

    if (rejectAndReturnCountOfLast7DaysA !== rejectAndReturnCountOfLast7DaysB) {
      return (
        rejectAndReturnCountOfLast7DaysA - rejectAndReturnCountOfLast7DaysB
      );
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天拒退数降序
  static compareByRejectAndReturnCountOfLast7DaysDesc(
    VipshopGoodsA,
    VipshopGoodsB,
  ) {
    return -VipshopGoods.compareByRejectAndReturnCountOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天销售量升序
  static compareBySalesQuantityOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const salesQuantityOfLast7DaysA =
      Number(VipshopGoodsA.salesQuantityOfLast7Days) || 0;
    const salesQuantityOfLast7DaysB =
      Number(VipshopGoodsB.salesQuantityOfLast7Days) || 0;

    if (salesQuantityOfLast7DaysA !== salesQuantityOfLast7DaysB) {
      return salesQuantityOfLast7DaysA - salesQuantityOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天销售量降序
  static compareBySalesQuantityOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareBySalesQuantityOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天销售额升序
  static compareBySalesAmountOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const salesAmountOfLast7DaysA =
      Number(VipshopGoodsA.salesAmountOfLast7Days) || 0;
    const salesAmountOfLast7DaysB =
      Number(VipshopGoodsB.salesAmountOfLast7Days) || 0;

    if (salesAmountOfLast7DaysA !== salesAmountOfLast7DaysB) {
      return salesAmountOfLast7DaysA - salesAmountOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天销售额降序
  static compareBySalesAmountOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareBySalesAmountOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天点击率升序
  static compareByClickThroughRateOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const clickThroughRateOfLast7DaysA =
      Number(VipshopGoodsA.clickThroughRateOfLast7Days) || 0;
    const clickThroughRateOfLast7DaysB =
      Number(VipshopGoodsB.clickThroughRateOfLast7Days) || 0;

    if (clickThroughRateOfLast7DaysA !== clickThroughRateOfLast7DaysB) {
      return clickThroughRateOfLast7DaysA - clickThroughRateOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天点击率降序
  static compareByClickThroughRateOfLast7DaysDesc(
    VipshopGoodsA,
    VipshopGoodsB,
  ) {
    return -VipshopGoods.compareByClickThroughRateOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天加购率升序
  static compareByAddToCartRateOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const addToCartRateOfLast7DaysA =
      Number(VipshopGoodsA.addToCartRateOfLast7Days) || 0;
    const addToCartRateOfLast7DaysB =
      Number(VipshopGoodsB.addToCartRateOfLast7Days) || 0;

    if (addToCartRateOfLast7DaysA !== addToCartRateOfLast7DaysB) {
      return addToCartRateOfLast7DaysA - addToCartRateOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天加购率降序
  static compareByAddToCartRateOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByAddToCartRateOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天转化率升序
  static compareByPurchaseRateOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const purchaseRateOfLast7DaysA =
      Number(VipshopGoodsA.purchaseRateOfLast7Days) || 0;
    const purchaseRateOfLast7DaysB =
      Number(VipshopGoodsB.purchaseRateOfLast7Days) || 0;

    if (purchaseRateOfLast7DaysA !== purchaseRateOfLast7DaysB) {
      return purchaseRateOfLast7DaysA - purchaseRateOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天转化率降序
  static compareByPurchaseRateOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByPurchaseRateOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天拒退率升序
  static compareByRejectAndReturnRateOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const rejectAndReturnRateOfLast7DaysA =
      Number(VipshopGoodsA.rejectAndReturnRateOfLast7Days) || 0;
    const rejectAndReturnRateOfLast7DaysB =
      Number(VipshopGoodsB.rejectAndReturnRateOfLast7Days) || 0;

    if (rejectAndReturnRateOfLast7DaysA !== rejectAndReturnRateOfLast7DaysB) {
      return rejectAndReturnRateOfLast7DaysA - rejectAndReturnRateOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天拒退率降序
  static compareByRejectAndReturnRateOfLast7DaysDesc(
    VipshopGoodsA,
    VipshopGoodsB,
  ) {
    return -VipshopGoods.compareByRejectAndReturnRateOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //近7天款销量升序
  static compareByStyleSalesOfLast7Days(VipshopGoodsA, VipshopGoodsB) {
    const styleSalesOfLast7DaysA =
      Number(VipshopGoodsA.styleSalesOfLast7Days) || 0;
    const styleSalesOfLast7DaysB =
      Number(VipshopGoodsB.styleSalesOfLast7Days) || 0;

    if (styleSalesOfLast7DaysA !== styleSalesOfLast7DaysB) {
      return styleSalesOfLast7DaysA - styleSalesOfLast7DaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //近7天款销量降序
  static compareByStyleSalesOfLast7DaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByStyleSalesOfLast7Days(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //可售库存升序
  static compareBySellableInventory(VipshopGoodsA, VipshopGoodsB) {
    const sellableInventoryA = Number(VipshopGoodsA.sellableInventory) || 0;
    const sellableInventoryB = Number(VipshopGoodsB.sellableInventory) || 0;

    if (sellableInventoryA !== sellableInventoryB) {
      return sellableInventoryA - sellableInventoryB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //可售库存降序
  static compareBySellableInventoryDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareBySellableInventory(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //可售天数升序
  static compareBySellableDays(VipshopGoodsA, VipshopGoodsB) {
    const sellableDaysA = Number(VipshopGoodsA.sellableDays) || 0;
    const sellableDaysB = Number(VipshopGoodsB.sellableDays) || 0;

    if (sellableDaysA !== sellableDaysB) {
      return sellableDaysA - sellableDaysB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //可售天数降序
  static compareBySellableDaysDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareBySellableDays(VipshopGoodsA, VipshopGoodsB);
  }

  //合计库存升序
  static compareByTotalInventory(VipshopGoodsA, VipshopGoodsB) {
    const totalInventoryA = Number(VipshopGoodsA.totalInventory) || 0;
    const totalInventoryB = Number(VipshopGoodsB.totalInventory) || 0;

    if (totalInventoryA !== totalInventoryB) {
      return totalInventoryA - totalInventoryB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //合计库存降序
  static compareByTotalInventoryDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByTotalInventory(VipshopGoodsA, VipshopGoodsB);
  }

  //成品合计升序
  static compareByFinishedGoodsTotalInventory(VipshopGoodsA, VipshopGoodsB) {
    const finishedGoodsTotalInventoryA =
      Number(VipshopGoodsA.finishedGoodsTotalInventory) || 0;
    const finishedGoodsTotalInventoryB =
      Number(VipshopGoodsB.finishedGoodsTotalInventory) || 0;

    if (finishedGoodsTotalInventoryA !== finishedGoodsTotalInventoryB) {
      return finishedGoodsTotalInventoryA - finishedGoodsTotalInventoryB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //成品合计降序
  static compareByFinishedGoodsTotalInventoryDesc(
    VipshopGoodsA,
    VipshopGoodsB,
  ) {
    return -VipshopGoods.compareByFinishedGoodsTotalInventory(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //通货合计升序
  static compareByGeneralGoodsTotalInventory(VipshopGoodsA, VipshopGoodsB) {
    const generalGoodsTotalInventoryA =
      Number(VipshopGoodsA.generalGoodsTotalInventory) || 0;
    const generalGoodsTotalInventoryB =
      Number(VipshopGoodsB.generalGoodsTotalInventory) || 0;

    if (generalGoodsTotalInventoryA !== generalGoodsTotalInventoryB) {
      return generalGoodsTotalInventoryA - generalGoodsTotalInventoryB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //通货合计降序
  static compareByGeneralGoodsTotalInventoryDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByGeneralGoodsTotalInventory(
      VipshopGoodsA,
      VipshopGoodsB,
    );
  }

  //销量总计升序
  static compareByTotalSales(VipshopGoodsA, VipshopGoodsB) {
    const totalSalesA = Number(VipshopGoodsA.totalSales) || 0;
    const totalSalesB = Number(VipshopGoodsB.totalSales) || 0;

    if (totalSalesA !== totalSalesB) {
      return totalSalesA - totalSalesB;
    }

    // 销量相同再比较款式号（按字符串排序）
    const styleA = String(VipshopGoodsA.styleNumber || "");
    const styleB = String(VipshopGoodsB.styleNumber || "");

    return styleA.localeCompare(styleB);
  }

  //销量总计降序
  static compareByTotalSalesDesc(VipshopGoodsA, VipshopGoodsB) {
    return -VipshopGoods.compareByTotalSales(VipshopGoodsA, VipshopGoodsB);
  }
}

//商品价格
class ProductPrice {
  static _wsName = "商品价格";
  static _keyToTitle = {
    designNumber: "设计号",
    itemNumber: "货号",
    picture: "图片",
    costPrice: "成本价",
    lowestPrice: "最低价",
    silverPrice: "白金价",
    userOperations1: "中台1",
    userOperations2: "中台2",
  };

  static _priceConfig = [
    {
      brandSN: "10016178",
      brand: "LAVI",
      packagingFee: 1.5, //发货打包费
      shippingCost: 1.5, //发货运费
      returnProcessingFee: 1.8, //退货整理费
      vipDiscountRate: 0.05, //超V优惠比例
      vipDiscountBearingRatio: 0.6, //超V优惠商家承担比例
      platformCommission: 0.3, //平台扣点
      brandCommission: 0.07, //品牌回款扣点
    },
    {
      brandSN: "10015023",
      brand: "巴帝巴帝",
      packagingFee: 1.5, //发货打包费
      shippingCost: 1.5, //发货运费
      returnProcessingFee: 1.8, //退货整理费
      vipDiscountRate: 0.05, //超V优惠比例
      vipDiscountBearingRatio: 0.6, //超V优惠商家承担比例
      platformCommission: 0.25, //平台扣点
      brandCommission: 0.12, //品牌回款扣点
    },
    {
      brandSN: "10000708",
      brand: "史努比",
      packagingFee: 1.5, //发货打包费
      shippingCost: 1.5, //发货运费
      returnProcessingFee: 1.8, //退货整理费
      vipDiscountRate: 0.02, //超V优惠比例
      vipDiscountBearingRatio: 0, //超V优惠商家承担比例
      platformCommission: 0.3, //平台扣点
      brandCommission: 0.11, //品牌回款扣点
    },
    {
      brandSN: "10000782",
      brand: "小猪班纳",
      packagingFee: 1.5, //发货打包费
      shippingCost: 1.5, //发货运费
      returnProcessingFee: 1.8, //退货整理费
      vipDiscountRate: 0.05, //超V优惠比例
      vipDiscountBearingRatio: 0.6, //超V优惠商家承担比例
      platformCommission: 0.27, //平台扣点
      brandCommission: 0.1, //品牌回款扣点
    },
  ];

  static _data = [];

  static initializeData() {
    this._data = DAO.readWorksheet(this._wsName, this, this._keyToTitle);

    //货号去重
    let duplicates = Utility.findDuplicatesByProperty(this._data, "itemNumber");
    if (duplicates.length != 0) {
      throw new CustomError(
        this._wsName +
          "中存在重复的货号：【" +
          duplicates +
          "】，请核查后重试！",
        { itemNumber: "货号", errReason: "错误原因" },
        duplicates,
      );
    }
  }

  //检查价格数据录入是否正确
  static checkData() {
    const abnormalProductPrice = [];

    for (const item of this._data) {
      if (!item.itemNumber) continue;

      const {
        costPrice,
        lowestPrice,
        silverPrice,
        userOperations1,
        userOperations2,
      } = item;
      const ops1 = userOperations1 ?? 0;
      const ops2 = userOperations2 ?? 0;

      const requiredProps = [costPrice, lowestPrice, silverPrice];
      const hasRequired = Utility.validateRequired(requiredProps);

      let hasError = false;
      let errReason = "";

      if (!hasRequired) {
        hasError = true;
        errReason = "没有填写成本价/最低价/白金价";
      } else {
        // 只有必填项通过时才验证数值
        const numberProps = [costPrice, lowestPrice, silverPrice, ops1, ops2];
        const costPriceNum = Number(costPrice);
        const lowestPriceNum = Number(lowestPrice);
        const silverPriceNum = Number(silverPrice);
        const ops1Num = Number(ops1);
        const ops2Num = Number(ops2);

        if (
          !Utility.validateIsNumber(numberProps) ||
          costPriceNum <= 0 ||
          lowestPriceNum <= 0 ||
          silverPriceNum <= 0 ||
          ops1Num < 0 ||
          ops2Num < 0
        ) {
          hasError = true;
          errReason = "成本价/最低价/白金价/中台1/中台2数据填写不正确";
        }
      }

      if (hasError) {
        item.errReason = errReason;
        abnormalProductPrice.push(item);
      }
    }

    if (abnormalProductPrice.length > 0) {
      throw new CustomError(
        "商品价格数据异常",
        { itemNumber: "货号", errReason: "错误原因" },
        abnormalProductPrice,
      );
    }
  }

  static getWsName() {
    return this._wsName;
  }
  //查找单个商品价格
  static findProductPrice(querys) {
    if (querys) {
      return this._data.find((item) => {
        for (let entry of Object.entries(querys)) {
          if (item[entry[0]] != entry[1] || !item[entry[0]]) {
            return false;
          }
        }
        return true;
      });
    }
    return undefined;
  }

  //查询品牌超V折扣率
  static getVipDiscountRate(brandSN) {
    let priceInfo = this._priceConfig?.find((item) => item.brandSN == brandSN);
    if (priceInfo) return priceInfo.vipDiscountRate;
    return 0;
  }

  //计算利润
  static calProfit(
    brandSN,
    costPrice,
    salesPrice,
    userOperations1 = 0,
    userOperations2 = 0,
    returnRate = 0.3,
  ) {
    try {
      // 1. 基础参数验证
      const requiredParams = [brandSN, costPrice, salesPrice];
      if (!Utility.validateRequired(requiredParams)) {
        return undefined;
      }

      // 2. 验证是否为正确的数字
      const isNumbers = [
        costPrice,
        salesPrice,
        userOperations1 ?? 0,
        userOperations2 ?? 0,
        returnRate ?? 0.3,
      ];
      if (!Utility.validateIsNumber(isNumbers)) {
        return undefined;
      }

      const params = {
        costPrice: Number(costPrice),
        salesPrice: Number(salesPrice),
        userOperations1: Number(userOperations1 ?? 0),
        userOperations2: Number(userOperations2 ?? 0),
        returnRate: Number(returnRate ?? 0.3),
      };

      // 验证数值范围
      if (
        params.costPrice <= 0 ||
        params.salesPrice <= 0 ||
        params.userOperations1 < 0 ||
        params.userOperations2 < 0 ||
        params.returnRate < 0 ||
        params.returnRate > 1 // 退货率应为0-1之间
      ) {
        return undefined;
      }

      // 退货率修正：如果退货率为100%或者小于30%，重置为默认值
      if (params.returnRate == 1 || params.returnRate < 0.3) {
        params.returnRate = 0.3;
      }

      // 3. 获取价格配置
      const priceInfo = this._priceConfig?.find(
        (item) => item.brandSN === brandSN,
      );
      if (!priceInfo) {
        throw new Error(`没有找到品牌SN【${brandSN}】的价格信息，请核实！`);
      }

      // 4. 提取配置参数
      const {
        packagingFee,
        shippingCost,
        returnProcessingFee,
        vipDiscountRate,
        vipDiscountBearingRatio,
        platformCommission,
        brandCommission,
      } = priceInfo;

      // 验证配置参数
      const configValid = [
        packagingFee,
        shippingCost,
        returnProcessingFee,
        vipDiscountRate,
        vipDiscountBearingRatio,
        platformCommission,
        brandCommission,
      ].every((value) => !isNaN(value) && value >= 0);

      if (!configValid) {
        return undefined;
      }

      // 5. 计算超V优惠金额
      let vipDiscountAmount;
      if (params.salesPrice > 50) {
        vipDiscountAmount = Math.round(params.salesPrice * vipDiscountRate);
      } else {
        vipDiscountAmount =
          Math.round(params.salesPrice * vipDiscountRate * 10) / 10;
      }

      // 6. 计算优惠后价格（确保不为负数）
      const priceAfterCoupon = Math.max(
        0,
        params.salesPrice - params.userOperations1 - params.userOperations2,
      );

      // 7. 分步计算利润

      // a. 基础计算
      const grossProfit = priceAfterCoupon - params.costPrice;

      // b. 固定费用
      const fixedCosts = packagingFee + shippingCost;

      // c. 退货相关成本
      const returnMultiplier = 1 / (1 - params.returnRate) - 1;
      const returnCosts = returnMultiplier * fixedCosts;
      const returnProcessingCost = params.returnRate * returnProcessingFee;

      // d. 优惠承担成本
      const vipDiscountCost = vipDiscountAmount * vipDiscountBearingRatio;

      // e. 平台佣金
      const platformFee = priceAfterCoupon * platformCommission;

      // 计算品牌佣金
      const brandCommissionBase =
        priceAfterCoupon * (1 - platformCommission) - vipDiscountCost;
      const brandFee = Math.max(0, brandCommissionBase) * brandCommission;

      // 8. 最终利润计算
      const profit =
        grossProfit -
        fixedCosts -
        returnCosts -
        returnProcessingCost -
        vipDiscountCost -
        platformFee -
        brandFee;

      // 9. 返回结果（限制小数位数）
      return parseFloat(profit.toFixed(2));
    } catch (error) {
      // 根据错误类型返回不同结果
      if (error.message.includes("没有找到品牌SN")) {
        throw error; // 重新抛出关键业务错误
      }
      return undefined;
    }
  }

  //计算利润率
  static calProfitRate(
    brandSN,
    costPrice,
    salesPrice,
    userOperations1 = 0,
    userOperations2 = 0,
    returnRate = 0.3,
  ) {
    let profit = this.calProfit(
      brandSN,
      costPrice,
      salesPrice,
      userOperations1,
      userOperations2,
      returnRate,
    );
    if (profit) {
      return Number((profit / costPrice).toFixed(5));
    } else {
      return undefined;
    }
  }

  //计算活动价
  static calProductActivityPrice(silverPrice) {
    const fifthActPrice = Math.floor(Number(silverPrice) * 0.95); //5级活动
    const fourthActPrice = Number(silverPrice); //4级活动价
    const thirdActPrice = (fourthActPrice / 0.9 + 0.06).toFixed(1); //3级活动价
    const secondActPrice = (thirdActPrice / 0.95 + 0.06).toFixed(1); //2级活动价
    const firstActPrice = (secondActPrice / 0.95 + 0.06).toFixed(1); //1级活动价
    return {
      firstActPrice,
      secondActPrice,
      thirdActPrice,
      fourthActPrice,
      fifthActPrice,
    };
  }

  toString() {
    return this.itemNumber;
  }
}

//常态商品
class RegularProduct {
  static _wsName = "常态商品";
  static _keyToTitle = {
    MID: "商品ID",
    P_SPU: "P_SPU",
    productCode: "条码",
    itemNumber: "货号",
    styleNumber: "款号",
    color: "颜色",
    size: "尺码",
    thirdLevelCategory: "三级品类",
    brandSN: "品牌SN",
    brand: "品牌名称",
    sizeStatus: "尺码状态",
    itemStatus: "商品状态",
    vipshopPrice: "唯品价",
    finalPrice: "到手价",
    sellableInventory: "可售库存",
    sellableDays: "可售天数",
  };

  static _data = [];

  static initializeData() {
    this._data = DAO.readWorksheet(this._wsName, this, this._keyToTitle);
  }

  static getWsName() {
    return this._wsName;
  }

  //查找单个常态商品
  static findRegularProduct(querys) {
    if (querys) {
      return this._data.find((item) => {
        for (let entry of Object.entries(querys)) {
          if (item[entry[0]] != entry[1] || !item[entry[0]]) {
            return false;
          }
        }
        return true;
      });
    }
    return undefined;
  }

  //查找符合条件的常态商品
  static filterRegularProducts(querys) {
    let filterItems = this._data;
    if (querys) {
      filterItems = this._data.filter((item) => {
        for (let entry of Object.entries(querys)) {
          if (item[entry[0]] != entry[1] || !item[entry[0]]) {
            return false;
          }
        }
        return true;
      });
    }
    return filterItems;
  }

  static compareBySize(productA, productB) {
    const sizeA = Number(productA.size);
    const sizeB = Number(productB.size);

    if (Number.isNaN(sizeA) || Number.isNaN(sizeB)) {
      return 0;
    }

    return sizeA - sizeB;
  }

  //清空常态商品
  static clear() {
    DAO.clearWorksheet(this._wsName);
  }

  toString() {
    return this.productCode;
  }
}

//组合商品
class ComboProduct {
  static _wsName = "组合商品";
  static _keyToTitle = {
    productCode: "组合商品实体编码",
    subProductCode: "商品编码",
    subProductQuantity: "数量",
  };

  static _date = [];

  static initializeData() {
    this._data = DAO.readWorksheet(this._wsName, this, this._keyToTitle);
  }

  //查找符合条件的子商品
  static filterComboProducts(querys) {
    let filterItems = this._data;
    if (querys) {
      filterItems = this._data.filter((item) => {
        for (let entry of Object.entries(querys)) {
          if (
            item[entry[0]] != entry[1] ||
            !item[entry[0]] ||
            item.subProductCode.startsWith("YH") ||
            item.subProductCode.startsWith("FL")
          ) {
            return false;
          }
        }
        return true;
      });
    }

    return filterItems;
  }

  //清空组合商品
  static clear() {
    DAO.clearWorksheet(this._wsName);
  }

  toString() {
    return (
      "组合商品实体编码：【" +
      this.productCode +
      "】,商品编码：【" +
      this.subProductCode +
      "】"
    );
  }
}

//商品库存
class Inventory {
  static _wsName = "商品库存";
  static _keyToTitle = {
    productCode: "商品编码",
    mainInventory: "数量",
    incomingInventory: "进货仓库存",
    finishingInventory: "后整车间",
    oversoldInventory: "超卖车间",
    prepareInventory: "备货车间",
    returnInventory: "销退仓库存",
    purchaseInventory: "采购在途数",
  };

  static _data = [];

  static initializeData() {
    this._data = DAO.readWorksheet(this._wsName, this, this._keyToTitle);
  }

  static getWsName() {
    return this._wsName;
  }

  //查找单个商品的库存
  static findInventory(querys) {
    if (querys) {
      return this._data.find((item) => {
        for (let entry of Object.entries(querys)) {
          if (item[entry[0]] != entry[1] || !item[entry[0]]) {
            return false;
          }
        }
        return true;
      });
    }
    return undefined;
  }

  //清空商品库存
  static clear() {
    DAO.clearWorksheet(this._wsName);
  }

  toString() {
    return this.productCode;
  }
}

//商品销售
class ProductSales {
  static _wsName = "商品销售";
  static _keyToTitle = {
    salesDate: "日期",
    itemNumber: "货号",
    exposureUV: "曝光UV",
    productDetailsUV: "商详UV",
    addToCartUV: "加购UV(加购用户数)",
    customerCount: "客户数",
    rejectAndReturnCount: "拒退件数",
    salesQuantity: "销售量",
    salesAmount: "销售额",
    firstListingTime: "首次上架时间",
  };

  static _data = [];

  static initializeData() {
    this._data = DAO.readWorksheet(this._wsName, this, this._keyToTitle);
  }

  static getWsName() {
    return this._wsName;
  }

  //查找单个商品销售
  static findProductSales(querys) {
    if (querys) {
      return this._data.find((item) => {
        for (let entry of Object.entries(querys)) {
          if (item[entry[0]] != entry[1] || !item[entry[0]]) {
            return false;
          }
        }
        return true;
      });
    }
    return undefined;
  }

  //查找符合条件的商品销售
  static filterProductSales(querys) {
    let filterItems = this._data;
    if (querys) {
      filterItems = this._data.filter((item) => {
        for (let entry of Object.entries(querys)) {
          if (item[entry[0]] != entry[1] || !item[entry[0]]) {
            return false;
          }
        }
        return true;
      });
    }
    return filterItems;
  }

  //清空商品销售
  static clear() {
    DAO.clearWorksheet(this._wsName);
  }

  toString() {
    return "日期：【" + this.salesDate + "】,货号：【" + this.itemNumber + "】";
  }
}

//系统记录
class SystemRecord {
  static _wsName = "系统记录";
  static _keyToTitle = {
    updateDateOfLast7Days: "近7天数据更新日期",
    updateDateOfProductPrice: "商品价格更新日期",
    updateDateOfRegularProduct: "常态商品更新日期",
    updateDateOfInventory: "商品库存更新日期",
    updateDateOfProductSales: "商品销售更新日期",
  };

  static _data = [];

  static initializeData() {
    this._data = DAO.readWorksheet(this._wsName, this, this._keyToTitle);
  }

  constructor() {
    for (let key of Object.keys(SystemRecord._keyToTitle)) {
      this[key] = undefined;
    }
    SystemRecord._data.push(this);
  }

  //返回记录
  static getSystemRecord() {
    return this._data[0] ? this._data[0] : new SystemRecord();
  }

  //更新记录
  static updateSystemRecord() {
    DAO.updateWorksheet(this._wsName, this._data, this._keyToTitle);
  }
}

//自定义错误
class CustomError extends Error {
  constructor(message, keyToTitle, data) {
    super(message);
    this.name = "CustomError";
    this.keyToTitle = keyToTitle;
    this.data = data;
    // 保持正确的调用栈
    Error.captureStackTrace(this, CustomError);
  }
}

class DAO {
  static _wbName = "商品运营表【史努比】"; //当前工作薄的名称

  static getWbName() {
    return this._wbName;
  }

  static readWorksheet(wsName, objType, keyToTitle, optionalKeyToTitle = {}) {
    let data = Workbooks(this._wbName)
      .Sheets(wsName)
      .UsedRange.Value2.filter((subArray) => Boolean(subArray.join("")));
    let titleRow = data.shift();

    if (!titleRow) {
      throw new Error("【" + wsName + "】中没有任何数据，请更新数据后重试！");
    }

    let titleIndex = {};

    for (let entry of Object.entries(keyToTitle)) {
      if (!titleRow.includes(entry[1]))
        throw new Error(
          "【" + wsName + "】中没有找到名称为：" + entry[1] + " 的列，请检查！",
        );

      titleIndex[entry[0]] = titleRow.indexOf(entry[1]);
    }

    for (let entry of Object.entries(optionalKeyToTitle)) {
      if (titleRow.includes(entry[1].replace(/^'/, "")))
        titleIndex[entry[0]] = titleRow.indexOf(entry[1].replace(/^'/, ""));
    }

    return data.map((item) => {
      let obj = new objType();
      for (let key of Object.keys(titleIndex)) {
        obj[key] = item[titleIndex[key]];
      }
      return obj;
    });
  }

  static updateWorksheet(
    wsName,
    data,
    keyToTitle,
    workbook = Workbooks(this._wbName),
  ) {
    let titleRow = [];
    let showData = [];

    for (let title of Object.values(keyToTitle)) {
      titleRow.push(title);
    }
    showData.push(titleRow);

    data.forEach((item) => {
      let dataRow = [];
      for (let key of Object.keys(keyToTitle)) {
        dataRow.push(item[key]);
      }
      showData.push(dataRow);
    });

    workbook.Sheets(wsName).Cells.ClearContents();
    workbook
      .Sheets(wsName)
      .Range("A1")
      .Resize(showData.length, showData[0].length).Value2 = showData;
    if (workbook == Workbooks(this._wbName)) {
      workbook.Save();
    }
  }

  static clearWorksheet(wsName, workbook = Workbooks(this._wbName)) {
    workbook.Sheets(wsName).Cells.ClearContents();
    if (workbook == Workbooks(this._wbName)) {
      workbook.Save();
    }
  }
}
//工具类
class Utility {
  //生成实时日期标题
  static generateDateKeyToTitle() {
    let result = {};

    let today = new Date();
    let thisYear = today.getFullYear();
    let thisMonth = today.getMonth();

    result["+" + (thisYear - 2)] = `'${thisYear - 2}`;
    result["+" + (thisYear - 1)] = `'${thisYear - 1}`;
    result["+" + thisYear] = `'${thisYear}`;

    for (let i = 1; i <= 12; i++) {
      result["+" + (thisYear - 1) + i.toString().padStart(2, "0")] =
        `'${thisYear - 1}` + i.toString().padStart(2, "0");
    }
    for (let i = 1; i <= thisMonth + 1; i++) {
      result["+" + thisYear + i.toString().padStart(2, "0")] =
        `'${thisYear}` + i.toString().padStart(2, "0");
    }
    for (let i = 45; i > 0; i--) {
      const date = new Date();
      date.setDate(today.getDate() - i);
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      result[`+${year}-${month}-${day}`] = `'${month}/${day}`;
    }

    return result;
  }

  //生成近7天日期段字符串
  static generateStringOfLast7Days() {
    let today = new Date();
    let startDate = new Date();
    let endDate = new Date();
    startDate.setDate(today.getDate() - 7);
    endDate.setDate(today.getDate() - 1);

    let startDateYear = startDate.getFullYear();
    let startDateMonth = String(startDate.getMonth() + 1).padStart(2, "0");
    let startDateDay = String(startDate.getDate()).padStart(2, "0");
    let endDateYear = endDate.getFullYear();
    let endDateMonth = String(endDate.getMonth() + 1).padStart(2, "0");
    let endDateDay = String(endDate.getDate()).padStart(2, "0");

    return `${startDateYear}-${startDateMonth}-${startDateDay}~${endDateYear}-${endDateMonth}-${endDateDay}`;
  }

  //生成销量合计日期标题
  static generateDateKeyToTitleForTotalSales() {
    let result = {};

    let today = new Date();
    let thisYear = today.getFullYear();
    let thisMonth = today.getMonth();

    result["+" + (thisYear - 2)] = `'${thisYear - 2}`;
    result["+" + (thisYear - 1)] = `'${thisYear - 1}`;

    for (let i = 1; i <= thisMonth; i++) {
      result["+" + thisYear + i.toString().padStart(2, "0")] =
        `'${thisYear}` + i.toString().padStart(2, "0");
    }
    for (let i = today.getDate() - 1; i > 0; i--) {
      const date = new Date();
      date.setDate(today.getDate() - i);
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      result[`+${year}-${month}-${day}`] = `'${month}/${day}`;
    }

    return result;
  }

  //判断日期是否为今天
  static isToday(date) {
    const today = new Date();

    return (
      date.getDate() === today.getDate() &&
      date.getMonth() === today.getMonth() &&
      date.getFullYear() === today.getFullYear()
    );
  }

  //生成昨天日期的字符串格式
  static generateStringOfYesterday() {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);

    const year = yesterday.getFullYear();
    const month = String(yesterday.getMonth() + 1).padStart(2, "0");
    const day = String(yesterday.getDate()).padStart(2, "0");

    return `+${year}-${month}-${day}`;
  }

  //检查单字段重复项目
  static findDuplicatesByProperty(data, prop) {
    let seen = new Set();
    let duplicates = [];

    data.forEach((item) => {
      let value = item[prop];
      if (seen.has(value) && value) {
        item.errReason = "重复";
        duplicates.push(item);
      } else {
        seen.add(value);
      }
    });

    return duplicates;
  }

  //验证数据不为空
  static validateRequired(requiredValues) {
    if (!Array.isArray(requiredValues) || requiredValues.length === 0) {
      return false;
    }

    return requiredValues.every(
      (value) =>
        value != null && // 排除 null 和 undefined
        typeof value !== "boolean" && // 排除布尔值
        value !== "" && // 排除空字符串
        (typeof value !== "string" || value.trim() !== ""), // 排除空格字符串
    );
  }

  //验证数据可否转化为数字
  static validateIsNumber(values) {
    if (!Array.isArray(values) || values.length === 0) {
      return false;
    }

    return values.every((value) => {
      // 不允许 undefined、null
      if (value == null) {
        return false;
      }

      // 排除布尔值
      if (typeof value === "boolean") {
        return false;
      }

      // 检查是否为有效数字
      const num = Number(value);
      return !isNaN(num) && isFinite(num);
    });
  }

  //获取货号筛选选项
  static getSelectOption() {
    let selectOption = {};

    if (UserForm1.CheckBox2.Value) {
      selectOption.mainSalesSeason = ["春秋"];
    }
    if (UserForm1.CheckBox3.Value) {
      selectOption.mainSalesSeason
        ? selectOption.mainSalesSeason.push("夏")
        : (selectOption.mainSalesSeason = ["夏"]);
    }
    if (UserForm1.CheckBox4.Value) {
      selectOption.mainSalesSeason
        ? selectOption.mainSalesSeason.push("冬")
        : (selectOption.mainSalesSeason = ["冬"]);
    }
    if (UserForm1.CheckBox5.Value) {
      selectOption.mainSalesSeason
        ? selectOption.mainSalesSeason.push("四季")
        : (selectOption.mainSalesSeason = ["四季"]);
    }

    if (UserForm1.CheckBox14.Value) {
      selectOption.applicableGender = ["男童"];
    }
    if (UserForm1.CheckBox15.Value) {
      selectOption.applicableGender
        ? selectOption.applicableGender.push("女童")
        : (selectOption.applicableGender = ["女童"]);
    }
    if (UserForm1.CheckBox16.Value) {
      selectOption.applicableGender
        ? selectOption.applicableGender.push("中性")
        : (selectOption.applicableGender = ["中性"]);
    }

    if (UserForm1.CheckBox17.Value) {
      selectOption.itemStatus = ["商品上线"];
    }
    if (UserForm1.CheckBox18.Value) {
      selectOption.itemStatus
        ? selectOption.itemStatus.push("部分上线")
        : (selectOption.itemStatus = ["部分上线"]);
    }
    if (UserForm1.CheckBox19.Value) {
      selectOption.itemStatus
        ? selectOption.itemStatus.push("商品下线")
        : (selectOption.itemStatus = ["商品下线"]);
    }

    if (UserForm1.CheckBox27.Value) {
      selectOption.offlineReason = ["新品下架"];
    }
    if (UserForm1.CheckBox28.Value) {
      selectOption.offlineReason
        ? selectOption.offlineReason.push("过季下架")
        : (selectOption.offlineReason = ["过季下架"]);
    }
    if (UserForm1.CheckBox29.Value) {
      selectOption.offlineReason
        ? selectOption.offlineReason.push("更换吊牌")
        : (selectOption.offlineReason = ["更换吊牌"]);
    }
    if (UserForm1.CheckBox35.Value) {
      selectOption.offlineReason
        ? selectOption.offlineReason.push("转移品牌")
        : (selectOption.offlineReason = ["转移品牌"]);
    }
    if (UserForm1.CheckBox36.Value) {
      selectOption.offlineReason
        ? selectOption.offlineReason.push("清仓淘汰")
        : (selectOption.offlineReason = ["清仓淘汰"]);
    }
    if (UserForm1.CheckBox37.Value) {
      selectOption.offlineReason
        ? selectOption.offlineReason.push("内网撞款")
        : (selectOption.offlineReason = ["内网撞款"]);
    }
    if (UserForm1.CheckBox39.Value) {
      selectOption.offlineReason
        ? selectOption.offlineReason.push("资质问题")
        : (selectOption.offlineReason = ["资质问题"]);
    }
    if (UserForm1.CheckBox40.Value) {
      selectOption.offlineReason
        ? selectOption.offlineReason.push("内在质检")
        : (selectOption.offlineReason = ["内在质检"]);
    }
    if (UserForm1.CheckBox41.Value) {
      selectOption.offlineReason
        ? selectOption.offlineReason.push(undefined)
        : (selectOption.offlineReason = [undefined]);
    }

    if (UserForm1.TextEdit1.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit1.Value)) {
        throw new Error("售龄输入必须是一个有效的数字");
      }
      selectOption.salesAge = [Number(UserForm1.TextEdit1.Value), undefined];
    }
    if (UserForm1.TextEdit11.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit11.Value)) {
        throw new Error("售龄输入必须是一个有效的数字");
      }
      if (selectOption.salesAge) {
        selectOption.salesAge[1] = Number(UserForm1.TextEdit11.Value);
      } else {
        selectOption.salesAge = [undefined, Number(UserForm1.TextEdit11.Value)];
      }
    }

    if (UserForm1.CheckBox20.Value) {
      selectOption.marketingPositioning = ["引流款"];
    }
    if (UserForm1.CheckBox21.Value) {
      selectOption.marketingPositioning
        ? selectOption.marketingPositioning.push("利润款")
        : (selectOption.marketingPositioning = ["利润款"]);
    }
    if (UserForm1.CheckBox22.Value) {
      selectOption.marketingPositioning
        ? selectOption.marketingPositioning.push("清仓款")
        : (selectOption.marketingPositioning = ["清仓款"]);
    }

    if (UserForm1.OptionButton23.Value) {
      selectOption.activityStatus = "活动中";
    }
    if (UserForm1.OptionButton24.Value) {
      selectOption.activityStatus = "未提报";
    }

    if (UserForm1.CheckBox8.Value) {
      selectOption.userOperations1 = true;
    }
    if (UserForm1.CheckBox9.Value) {
      selectOption.userOperations2 = true;
    }

    if (UserForm1.TextEdit3.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit3.Value)) {
        throw new Error("利润输入必须是一个有效的数字");
      }
      selectOption.profit = [Number(UserForm1.TextEdit3.Value), undefined];
    }
    if (UserForm1.TextEdit4.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit4.Value)) {
        throw new Error("利润输入必须是一个有效的数字");
      }
      if (selectOption.profit) {
        selectOption.profit[1] = Number(UserForm1.TextEdit4.Value);
      } else {
        selectOption.profit = [undefined, Number(UserForm1.TextEdit4.Value)];
      }
    }

    if (UserForm1.TextEdit5.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit5.Value)) {
        throw new Error("利润率输入必须是一个有效的数字");
      }
      selectOption.profitRate = [Number(UserForm1.TextEdit5.Value), undefined];
    }
    if (UserForm1.TextEdit6.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit6.Value)) {
        throw new Error("利润率输入必须是一个有效的数字");
      }
      if (selectOption.profitRate) {
        selectOption.profitRate[1] = Number(UserForm1.TextEdit6.Value);
      } else {
        selectOption.profitRate = [
          undefined,
          Number(UserForm1.TextEdit6.Value),
        ];
      }
    }

    if (UserForm1.CheckBox10.Value) {
      selectOption.isPriceBroken = true;
    }

    if (UserForm1.CheckBox23.Value) {
      selectOption.stockingMode = ["现货"];
    }
    if (UserForm1.CheckBox24.Value) {
      selectOption.stockingMode
        ? selectOption.stockingMode.push("通版通货")
        : (selectOption.stockingMode = ["通版通货"]);
    }
    if (UserForm1.CheckBox25.Value) {
      selectOption.stockingMode
        ? selectOption.stockingMode.push("专版通货")
        : (selectOption.stockingMode = ["专版通货"]);
    }

    if (UserForm1.TextEdit7.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit7.Value)) {
        throw new Error("可售库存输入必须是一个有效的数字");
      }
      selectOption.sellableInventory = [
        Number(UserForm1.TextEdit7.Value),
        undefined,
      ];
    }
    if (UserForm1.TextEdit8.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit8.Value)) {
        throw new Error("可售库存输入必须是一个有效的数字");
      }
      if (selectOption.sellableInventory) {
        selectOption.sellableInventory[1] = Number(UserForm1.TextEdit8.Value);
      } else {
        selectOption.sellableInventory = [
          undefined,
          Number(UserForm1.TextEdit8.Value),
        ];
      }
    }

    if (UserForm1.CheckBox12.Value) {
      selectOption.isOutOfStock = true;
    }

    if (UserForm1.TextEdit9.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit9.Value)) {
        throw new Error("可售天数输入必须是一个有效的数字");
      }
      selectOption.sellableDays = [
        Number(UserForm1.TextEdit9.Value),
        undefined,
      ];
    }
    if (UserForm1.TextEdit10.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit10.Value)) {
        throw new Error("可售天数输入必须是一个有效的数字");
      }
      if (selectOption.sellableDays) {
        selectOption.sellableDays[1] = Number(UserForm1.TextEdit10.Value);
      } else {
        selectOption.sellableDays = [
          undefined,
          Number(UserForm1.TextEdit10.Value),
        ];
      }
    }

    if (UserForm1.TextEdit17.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit17.Value)) {
        throw new Error("成品库存输入必须是一个有效的数字");
      }
      selectOption.finishedGoodsTotalInventory = [
        Number(UserForm1.TextEdit17.Value),
        undefined,
      ];
    }
    if (UserForm1.TextEdit18.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit18.Value)) {
        throw new Error("成品库存输入必须是一个有效的数字");
      }
      if (selectOption.finishedGoodsTotalInventory) {
        selectOption.finishedGoodsTotalInventory[1] = Number(
          UserForm1.TextEdit18.Value,
        );
      } else {
        selectOption.finishedGoodsTotalInventory = [
          undefined,
          Number(UserForm1.TextEdit18.Value),
        ];
      }
    }

    if (UserForm1.TextEdit19.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit19.Value)) {
        throw new Error("通货库存输入必须是一个有效的数字");
      }
      selectOption.generalGoodsTotalInventory = [
        Number(UserForm1.TextEdit19.Value),
        undefined,
      ];
    }
    if (UserForm1.TextEdit20.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit20.Value)) {
        throw new Error("通货库存输入必须是一个有效的数字");
      }
      if (selectOption.generalGoodsTotalInventory) {
        selectOption.generalGoodsTotalInventory[1] = Number(
          UserForm1.TextEdit20.Value,
        );
      } else {
        selectOption.generalGoodsTotalInventory = [
          undefined,
          Number(UserForm1.TextEdit20.Value),
        ];
      }
    }

    if (UserForm1.TextEdit21.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit21.Value)) {
        throw new Error("合计库存输入必须是一个有效的数字");
      }
      selectOption.totalInventory = [
        Number(UserForm1.TextEdit21.Value),
        undefined,
      ];
    }
    if (UserForm1.TextEdit22.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit22.Value)) {
        throw new Error("合计库存输入必须是一个有效的数字");
      }
      if (selectOption.totalInventory) {
        selectOption.totalInventory[1] = Number(UserForm1.TextEdit22.Value);
      } else {
        selectOption.totalInventory = [
          undefined,
          Number(UserForm1.TextEdit22.Value),
        ];
      }
    }

    if (UserForm1.TextEdit14.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit14.Value)) {
        throw new Error("近7天销量输入必须是一个有效的数字");
      }
      selectOption.salesQuantityOfLast7Days = [
        Number(UserForm1.TextEdit14.Value),
        undefined,
      ];
    }
    if (UserForm1.TextEdit15.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit15.Value)) {
        throw new Error("近7天销量输入必须是一个有效的数字");
      }
      if (selectOption.salesQuantityOfLast7Days) {
        selectOption.salesQuantityOfLast7Days[1] = Number(
          UserForm1.TextEdit15.Value,
        );
      } else {
        selectOption.salesQuantityOfLast7Days = [
          undefined,
          Number(UserForm1.TextEdit15.Value),
        ];
      }
    }

    if (UserForm1.TextEdit16.Value) {
      if (!/^-?\d+(\.\d+)?$/.test(UserForm1.TextEdit16.Value)) {
        throw new Error("近7天销量排名输入必须是一个有效的数字");
      }
      selectOption.topProductsBySales = Number(UserForm1.TextEdit16.Value);
    }
    return selectOption;
  }
}
