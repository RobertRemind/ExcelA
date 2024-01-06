﻿

/**
 * Add two numbers
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
function add(a, b) {
  
  return a + b;
}

function dim(dimension, filters) {
  
  return dimension
}

/**
 * Creates an SQL CREATE TABLE statement.
 * @customfunction MAKE_SQL
 * @param {string} tableName Name of the table.
 * @param {string[][]} columnNames Range of column names.
 * @param {string[][]} dataTypes Range of data types.
 * @param {string[][]} precision Range of precision for each data type.
 * @returns {string} The SQL CREATE TABLE statement.
 */
function makeSQL(tableName, columnNames, dataTypes, precision) {
  let sqlStatement = `CREATE TABLE ${tableName} (`;
  let columns = [];

  for (let i = 0; i < columnNames.length; i++) {
      let columnName = columnNames[i][0];
      let dataType = dataTypes[i][0];
      let precisionValue = precision[i][0];

      let columnDef = `${columnName} ${dataType}`;
      if (precisionValue) {
          columnDef += `(${precisionValue})`;
      }
      columns.push(columnDef);
  }

  sqlStatement += columns.join(', ') + ');';
  return sqlStatement;
}

CustomFunctions.associate("ADD",add);
CustomFunctions.associate("DIM",dim);
CustomFunctions.associate("MAKE_SQL",makeSQL);