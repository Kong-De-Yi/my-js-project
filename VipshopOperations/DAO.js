//数据访问层
class DAO {
  constructor(brandName, config) {
    this.workbookName = brandName;
    this.config = config;
    this.cache = new CacheManager();
    this.columnIndexes = new Map(); // 列索引缓存
  }

  readWorksheet(entityName, entityClass) {
    const cacheKey = `${entityName}_${new Date().toDateString()}`;
    if (this.cache.has(cacheKey)) {
      return this.cache.get(cacheKey);
    }

    const entityConfig = this.config.entityConfigs[entityName];
    const ws = Workbooks(this.workbookName).Sheets(entityConfig.wsName);
    const data = ws.UsedRange.Value2;

    if (!data || data.length < 2) {
      throw new Error(
        `【${entityConfig.wsName}】中没有任何数据，请更新数据后重试！`,
      );
    }

    // 构建列索引（带缓存）
    const columnIndexes = this.buildColumnIndexes(
      data[0],
      entityConfig.mappings,
    );

    // 验证并转换数据
    const entities = this.parseAndValidateData(
      data.slice(1),
      entityClass,
      columnIndexes,
      entityConfig.validations,
    );

    this.cache.set(cacheKey, entities);
    return entities;
  }

  buildColumnIndexes(headers, mappings) {
    const cacheKey = headers.join("|");
    if (this.columnIndexes.has(cacheKey)) {
      return this.columnIndexes.get(cacheKey);
    }

    const indexes = {};
    headers.forEach((header, index) => {
      const propertyName = this.findPropertyName(header, mappings);
      if (propertyName) {
        indexes[propertyName] = index;
      }
    });

    this.columnIndexes.set(cacheKey, indexes);
    return indexes;
  }

  findPropertyName(header, mappings) {
    // 先在固定映射中查找
    for (const [prop, colName] of Object.entries(mappings.fixed || {})) {
      if (colName === header) return prop;
    }

    // 在动态映射中查找
    for (const [pattern, generator] of Object.entries(mappings.dynamic || {})) {
      if (new RegExp(pattern).test(header)) {
        return generator(header);
      }
    }

    return null;
  }

  parseAndValidateData(rows, EntityClass, columnIndexes, validations) {
    return rows
      .filter((row) => row.some((cell) => cell != null && cell !== ""))
      .map((row, rowIndex) => {
        const entity = new EntityClass();
        entity._rowNumber = rowIndex + 2; //Excel行号

        //设置属性并验证
        for (const [property, colIndex] of Object.entities(columnIndexes)) {
          const value = row[colIndex];

          //动态属性特殊处理
          if (property.startsWith("_dynamic_")) {
            entity.setDynamicPropery(property, value);
          } else {
            entity[property] = value;
          }

          //执行验证
          this.validateProperty(entity, property, value, validations[property]);
        }

        return entity;
      });
  }

  validateProperty(entity, property, value, validationRules) {
    if (!validationRules) return;

    const errors = [];

    for (const rule of validationRules) {
      const result = rule.validate(value, entity);
      if (!result.isValid) {
        errors.push(result.message);
      }
    }

    if (errors.length > 0) {
      entity.addValidationError(property, errors.join(";"));
    }
  }
}
