//配置管理器
class Config {
  constructor(brandName) {
    this.brandName = brandName;
    this.brandConfig = this.loadBrandConfig();
    this.entityConfigs = this.loadEntityConfigs();
  }

  loadBrandConfig() {
    const configSheet = Workbooks(this.brandName).Sheets("品牌配置");
    if (!configSheet)
      throw new Error(`找不到【${this.brandName}】的品牌配置表`);

    const data = configSheet.UsedRange.Value2;
    const headers = data[0];
    const configRow = data.find((row) => row[0] === this.brandName);

    if (!configRow) throw new Error(`品牌配置中找不到【${this.brandName}】`);

    const config = {};
    headers.forEach((header, index) => {
      config[this.toCamelCase(header)] = configRow[index];
    });

    return config;
  }

  loadEntityConfigs() {
    // 从配置文件加载实体类配置
    return {
      VipshopGoods: {
        wsName: "货号总表",
        brandSN: this.brandConfig.brandSN,
        dynamicColumns: this.generateDynamicColumnConfigs(),
        validations: this.loadValidations("VipshopGoods"),
        mappings: this.loadColumnMappings("VipshopGoods"),
      },
      ProductPrice: {
        wsName: "商品价格",
        validations: this.loadValidations("ProductPrice"),
        mappings: this.loadColumnMappings("ProductPrice"),
      },
      RegularProduct: {
        wsName: "常态商品",
        validations: this.loadValidations("RegularProduct"),
        mappings: this.loadColumnMappings("RegularProduct"),
      },
      ComboProduct: {
        wsName: "组合商品",
        validations: this.loadValidations("ComboProduct"),
        mappings: this.loadColumnMappings("ComboProduct"),
      },
      Inventory: {
        wsName: "商品库存",
        validations: this.loadValidations("Inventory"),
        mappings: this.loadColumnMappings("Inventory"),
      },
    };
  }

  generateDynamicColumnConfigs() {
    return {
      dateColumns: {
        pattern: /^\d{4}-\d{2}-\d{2}$/,
        generator: this.generateDateColumns.bind(this),
        storage: "dynamicDateColumns",
      },
      yearColumns: {
        pattern: /^\d{4}$/,
        generator: this.generateYearColumns.bind(this),
        storage: "dynamicYearColumns",
      },
      monthColumns: {
        pattern: /^\d{4}\d{2}$/,
        generator: this.generateMonthColumns.bind(this),
        storage: "dynamicMonthColumns",
      },
    };
  }

  toCamelCase(str) {
    return str.replace(/(?:^\w|[A-Z]|\b\w|\s+)/g, (match, index) => {
      if (+match === 0) return "";
      return index === 0 ? match.toLowerCase() : match.toUpperCase();
    });
  }
}
