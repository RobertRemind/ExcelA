﻿

/**
 * @customfunction
 * @description Generates an SQL create table statement
 * @param {any} tableName The name of the table to be created.
 * @param {any} columnNames The name of the table columns.
 * @param {any} dataTypes The type of the table columns.
 * @param {any} precision The precision of the table columns.
 */
function makeSQL (tableName, columnNames, dataTypes, precision){
  debugger;
  let sqlStatement = `CREATE TABLE ${tableName} (\n`;
    let columns = [];

    for (let i = 0; i < columnNames.length; i++) {
        let columnName = columnNames[i][0];
        let dataType = dataTypes[i][0];
        let precisionValue = precision[i][0];

        let columnDef = `\t${columnName} ${dataType}`;
        if (precisionValue) {
            columnDef += `(${precisionValue})`;
        }
        columns.push(columnDef);
    }

    sqlStatement += columns.join(',\n') + '\n);';
    return sqlStatement;
}


/**
 * Creates a JSON string for SQL mappings.
 * @customfunction
 * @description Generates an JSON object for Spotify mapping
 * @param {any} tableName Name of the SQL table for every element.
 * @param {any} columnNames Range of cells for the "sqlColumn" attribute.
 * @param {any} paths Range of cells for the "path" attribute.
 * @param {any} dataTypes Range of cells for the "type" attribute.
 * @param {any} precision Range of cells for the "precision" attribute.
 */
function generateJsonMap(tableName, columnNames, paths, dataTypes, precision) {
  const sqlMappings = {
    SQLMappings: [
      {
        table: tableName,
        columnsMap: []
      }
    ]
  };

  // Find the length of the longest array
  const maxLength = Math.max(columnNames.length, paths.length, dataTypes.length, precision.length);

  for (let i = 0; i < maxLength; i++) {
    const columnMap = {
      sqlColumn: columnNames[i] && columnNames[i][0], // Check for undefined
      path: paths[i] && paths[i][0], // Check for undefined
      type: dataTypes[i] && dataTypes[i][0], // Check for undefined
      precision: precision[i] && precision[i][0], // Check for undefined
      nullable: false // Assuming nullable is always false as per the example
    };

    // If an attribute is undefined or empty, delete it from the columnMap object
    Object.keys(columnMap).forEach(key => {
      if (columnMap[key] === undefined || columnMap[key] === '') {
        delete columnMap[key];
      }
    });

    sqlMappings.SQLMappings[0].columnsMap.push(columnMap);
  }

  return JSON.stringify(sqlMappings, null, 2); // Pretty print the JSON
}



/**
 * @customfunction
 * @description Adds two numbers together. 
 * @param {number} first First number to be added.
 * @param {number} second Second number to be added.
 */
function add(first, second){
  debugger;
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




CustomFunctions.associate("MAKESQL", makeSQL);
CustomFunctions.associate("ADD", add);
CustomFunctions.associate("STOREVALUE",StoreValue);
CustomFunctions.associate("GETVALUE",GetValue);