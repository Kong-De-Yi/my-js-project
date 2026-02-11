//应用入口
class App {
  constructor() {
    this.brand = this.detectBrand();
    this.config = new Config(this.brand);
    this.dao = new DAO(this.brand, this.config);
    this.init();
  }

  detectBrand() {
    const workbookName = ActiveWorkbook.Name;
    const match = workbookName.match(/【(.+?)】/);
    if (!match) throw new Error("工作簿名称格式不正确，无法识别品牌");
    return match[1];
  }

  init() {
    // 初始化UI
    UI.init(this.config);
  }
}
