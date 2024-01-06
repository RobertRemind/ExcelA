﻿﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

/**
 * @customfunction
 * @description Adds two numbers together. 
 * @param {number} first First number to be added.
 * @param {number} second Second number to be added.
 */
function add(first, second){
  return first + second;
}

/**
 * @customfunction
 * @description Stores a value in Office.storage.
 * @param {any} key Key in the key-value pair you will store. 
 * @param {any} value Value in the key-value pair you will store. 
 */
function StoreValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

/**
 * @customfunction
 * @description Gets value from Office.storage. 
 * @param {any} key Key of item you intend to get.
 */
function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
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
function MakeSQL(tableName, columnNames, dataTypes, precision) {
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



CustomFunctions.associate("ADD", add);
CustomFunctions.associate("STOREVALUE",StoreValue);
CustomFunctions.associate("GETVALUE",GetValue);
CustomFunctions.associate("MAKESQL",MakeSQL);