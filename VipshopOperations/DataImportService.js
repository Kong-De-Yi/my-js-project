/**
 * 数据导入服务 - 负责从"导入数据"工作表导入数据到各个业务实体
 *
 * @class DataImportService
 * @description 作为系统的数据导入核心，提供以下功能：
 * - 自动识别导入数据的实体类型（基于表头匹配必填字段）
 * - 支持两种导入模式：
 *   - overwrite（覆盖模式）：直接覆盖目标工作表的全部数据
 *   - append（追加模式）：基于主键进行新增或更新（存在则更新，不存在则新增）
 * - 数据验证：导入前对数据进行完整性校验
 * - 系统记录更新：导入成功后更新对应实体的导入日期
 * - 清空临时数据：导入完成后自动清空"导入数据"工作表
 *
 * 导入流程：
 * 1. 读取"导入数据"工作表的内容
 * 2. 提取表头并标准化（去除空格）
 * 3. 遍历可导入实体，匹配必填字段
 * 4. 根据实体配置的导入模式执行导入
 * 5. 更新系统记录
 * 6. 清空临时数据
 *
 * 该类采用单例模式，确保全局只有一个数据导入服务实例。
 *
 * @example
 * // 获取导入服务实例
 * const importService = DataImportService.getInstance(repository, excelDAO);
 *
 * // 执行导入
 * try {
 *   const result = importService.import();
 *   MsgBox(result.message);
 * } catch (e) {
 *   MsgBox("导入失败：" + e.message);
 * }
 */
class DataImportService {
  /** @type {DataImportService} 单例实例 */
  static _instance = null;

  /**
   * 创建数据导入服务实例
   * @param {Repository} [repository] - 数据仓库实例，若不提供则自动获取
   * @param {ExcelDAO} [excelDAO] - Excel数据访问对象实例，若不提供则自动获取
   */
  constructor(repository, excelDAO) {
    if (DataImportService._instance) {
      return DataImportService._instance;
    }

    this._repository = repository || Repository.getInstance();
    this._excelDAO = excelDAO || ExcelDAO.getInstance();
    this._config = DataConfig.getInstance();
    this._importableEntities = this._config.getImportableEntities();

    this._validationEngine = ValidationEngine.getInstance();

    DataImportService._instance = this;
  }

  /**
   * 获取数据导入服务的单例实例
   * @static
   * @param {Repository} [repository] - 数据仓库实例
   * @param {ExcelDAO} [excelDAO] - Excel数据访问对象实例
   * @returns {DataImportService} 数据导入服务实例
   */
  static getInstance(repository, excelDAO) {
    if (!DataImportService._instance) {
      DataImportService._instance = new DataImportService(repository, excelDAO);
    }
    return DataImportService._instance;
  }

  /**
   * 验证实体是否支持导入
   * @private
   * @param {string} entityName - 实体名称
   * @returns {boolean} true=可导入，false=不可导入
   */
  _canImport(entityName) {
    return this._importableEntities.includes(entityName);
  }

  /**
   * 获取实体的导入模式
   * @private
   * @param {string} entityName - 实体名称
   * @returns {string|null} 导入模式：'overwrite' 或 'append'，不可导入时返回null
   */
  _getImportMode(entityName) {
    if (!this._canImport(entityName)) return null;

    const entity = this._config.get(entityName);
    return entity?.importMode || null;
  }

  /**
   * 检查表头是否匹配实体的必填字段
   * @private
   * @param {string} entityName - 实体名称
   * @param {string[]} headers - 表头数组
   * @returns {boolean} true=匹配，false=不匹配
   * @description
   * 匹配规则：表头必须包含实体配置中 requiredTitles 定义的所有字段
   * 例如：
   * - 商品导入需要包含：货号、款号、颜色等必填字段
   * - 库存导入需要包含：条码、主库存、在途库存等必填字段
   */
  _matchesEntity(entityName, headers) {
    const entity = this._config.get(entityName);
    if (!entity) return false;

    const requiredTitles = entity.requiredTitles || [];
    if (requiredTitles.length === 0) return false;

    // 检查是否包含所有必填字段
    return requiredTitles.every((title) => headers.includes(title));
  }

  /**
   * 识别导入数据的实体类型
   * @private
   * @param {string[]} headers - 表头数组
   * @returns {string|null} 实体名称，无法识别时返回null
   * @description
   * 识别流程：
   * 1. 标准化表头：去除每个表头的前后空格
   * 2. 遍历所有可导入实体
   * 3. 对每个实体调用 _matchesEntity 检查是否匹配
   * 4. 返回第一个匹配的实体名称
   *
   * 注意：如果有多个实体匹配，返回第一个。因此实体的匹配规则应该互斥。
   */
  _identify(headers) {
    if (!headers || headers.length === 0) {
      return null;
    }

    // 标准化表头：去除前后空格
    const normalizedHeaders = headers.map((h) => String(h).trim());

    for (const entityName of this._importableEntities) {
      if (this._matchesEntity(entityName, normalizedHeaders)) {
        return entityName;
      }
    }

    return null;
  }

  /**
   * 覆盖模式导入数据
   * @private
   * @param {string} entityName - 实体名称
   * @param {Object[]} items - 要导入的数据数组
   * @returns {Object} 导入结果
   * @returns {boolean} return.success - 是否成功
   * @returns {string} return.entityName - 工作表名称
   * @returns {string} return.mode - 导入模式（'overwrite'）
   * @returns {number} return.total - 导入数据总数
   * @returns {string} return.message - 导入结果消息
   * @description
   * 覆盖模式特点：
   * - 直接调用 repository.save 覆盖目标工作表的全部数据
   * - 不保留历史数据
   * - 适用于需要完全替换的场景（如每日全量导入）
   */
  _overwriteData(entityName, items) {
    // 直接保存（覆盖）
    this._repository.save(entityName, items);

    const entityConfig = this._config.get(entityName);
    return {
      success: true,
      entityName: entityConfig.worksheet,
      mode: "overwrite",
      total: items.length,
      message: `成功导入 ${items.length} 条数据到【${entityConfig.worksheet}】`,
    };
  }

  /**
   * 追加模式导入数据
   * @private
   * @param {string} entityName - 实体名称
   * @param {Object[]} newItems - 要导入的新数据数组
   * @returns {Object} 导入结果
   * @returns {boolean} return.success - 是否成功
   * @returns {string} return.entityName - 工作表名称
   * @returns {string} return.mode - 导入模式（'append'）
   * @returns {number} return.total - 导入数据总数
   * @returns {number} return.new - 新增数据条数
   * @returns {number} return.updated - 更新数据条数
   * @returns {string} return.message - 导入结果消息
   * @throws {Error} 如果实体未配置主键，抛出错误
   *
   * @description
   * 追加模式流程：
   * 1. 验证实体必须配置主键（uniqueKey）
   * 2. 验证所有新数据的完整性
   * 3. 读取历史数据
   * 4. 基于主键建立映射：
   *    - 主键存在于历史数据中 -> 更新
   *    - 主键不存在于历史数据中 -> 新增
   * 5. 合并数据并保存
   *
   * 适用于需要保留历史数据、增量更新的场景。
   */
  _appendData(entityName, newItems) {
    // 检查业务实体是否配置主键
    const entityConfig = this._config.get(entityName);
    const uniqueKeyConfig = this._config.parseUniqueKey(entityConfig.uniqueKey);
    const fields = uniqueKeyConfig.fields;

    if (fields.length === 0) {
      throw new Error("追加模式导入的业务实体必须配置主键");
    }

    // 提前验证新数据
    const validationResult = this._validationEngine.validateAll(
      newItems,
      entityConfig,
    );

    if (!validationResult.valid) {
      const errorMsg = this._validationEngine.formatErrors(
        validationResult,
        entityConfig.worksheet,
      );
      throw new Error(errorMsg);
    }

    // 读取历史数据
    let existingItems = [];
    existingItems = this._repository.findAll(entityName);

    // 建立索引：货号+日期
    const existingMap = new Map();
    existingItems.forEach((item) => {
      const key = fields.map((field) => item[field]).join("|");
      existingMap.set(key, item);
    });

    // 合并数据
    const mergedItems = [...existingItems];
    let updatedCount = 0;
    let newCount = 0;

    newItems.forEach((newItem) => {
      const key = fields.map((field) => newItem[field]).join("|");

      if (existingMap.has(key)) {
        // 更新
        const index = mergedItems.findIndex((item) => {
          // 使用配置的主键字段进行比较
          return fields.every((field) => item[field] === newItem[field]);
        });
        if (index !== -1) {
          mergedItems[index] = newItem;
          updatedCount++;
        }
      } else {
        // 新增
        mergedItems.push(newItem);
        newCount++;
      }
    });

    // 保存
    this._repository.save(entityName, mergedItems);

    return {
      success: true,
      entityName: entityConfig.worksheet,
      mode: "append",
      total: newItems.length,
      new: newCount,
      updated: updatedCount,
      message: `【${entityConfig.worksheet}】导入完成：：新增${newCount}条，更新${updatedCount}条`,
    };
  }

  /**
   * 执行数据导入
   * @returns {Object} 导入结果
   * @returns {boolean} return.success - 是否成功
   * @returns {string} return.entityName - 工作表名称
   * @returns {string} return.mode - 导入模式
   * @returns {number} return.total - 导入数据总数
   * @returns {number} [return.new] - 新增数据条数（追加模式）
   * @returns {number} [return.updated] - 更新数据条数（追加模式）
   * @returns {string} return.message - 导入结果消息
   * @throws {Error} 当以下情况时抛出错误：
   * - 找不到"导入数据"工作表
   * - 工作表中没有数据
   * - 无法识别实体类型
   * - 实体不支持导入
   * - 数据验证失败
   * - 追加模式但实体未配置主键
   *
   * @description
   * 完整的导入流程：
   * 1. 获取"导入数据"工作表并验证
   * 2. 提取表头（第一行）
   * 3. 识别实体类型（基于表头匹配）
   * 4. 验证实体是否支持导入
   * 5. 获取导入模式（覆盖/追加）
   * 6. 读取数据（调用 ExcelDAO.read）
   * 7. 根据模式执行导入（覆盖或追加）
   * 8. 更新系统记录的导入日期
   * 9. 清空"导入数据"工作表
   *
   * @example
   * // 执行导入
   * const result = importService.import();
   * if (result.success) {
   *   MsgBox(result.message);
   * }
   */
  import() {
    // 1. 获取导入数据
    const wb = this._excelDAO.getWorkbook();
    const sheet = wb.Sheets("导入数据");

    if (!sheet) {
      throw new Error("未找到【导入数据】工作表");
    }

    const usedRange = sheet.UsedRange;
    if (!usedRange || usedRange.Value2 == null) {
      throw new Error("【导入数据】工作表中没有数据");
    }

    const data = usedRange.Value2;
    if (data.length < 2) {
      throw new Error("【导入数据】工作表中只有标题行，没有数据");
    }

    // 2. 提取表头
    const headers = data[0].map((h) => String(h).trim());

    // 3. 识别实体类型
    const entityName = this._identify(headers);

    if (!entityName) {
      const requiredTitles = this._importableEntities
        .map((entityName) => {
          const entityConfig = this._config.get(entityName);
          return `【${entityConfig.worksheet}】: ${entityConfig.requiredTitles.toString()}`;
        })
        .join("\n");
      throw new Error(
        "无法识别导入数据的类型，请确保表头包含以下必填字段之一：\n" +
          requiredTitles,
      );
    }

    // 4. 验证是否支持导入
    if (!this._canImport(entityName)) {
      throw new Error(`实体【${entityName}】不支持导入操作`);
    }

    // 5. 获取导入模式
    const mode = this._getImportMode(entityName);

    // 6. 读取数据
    const items = this._excelDAO.read(entityName, "导入数据");

    // 7. 根据模式处理数据
    let result;
    if (mode === "append") {
      result = this._appendData(entityName, items);
    } else {
      result = this._overwriteData(entityName, items);
    }

    // 8.更新系统记录
    this._repository.updateSystemRecord(entityName, "importDate");

    // 9. 清空导入数据表
    this._excelDAO.clear("ImportData");

    return result;
  }
}
