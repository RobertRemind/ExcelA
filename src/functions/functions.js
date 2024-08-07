﻿﻿

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
  
  let sqlStatement = 'DECLARE @queryId int;\n\n'
  sqlStatement += `IF object_id('${tableName}') is null BEGIN \n`
  sqlStatement += `CREATE TABLE ${tableName} (\n`;

  
    let columns = [];

    for (let i = 0; i < columnNames.length; i++) {
      
      if (columnNames[i][0] != 0 && columnNames[i][0] != "0" && columnNames[i][0] != "") {
      
        let columnName = `[${columnNames[i][0]}]`;
        let dataType = dataTypes[i][0];
        let precisionValue = precision[i][0];

        let columnDef = `\t${columnName} ${dataType}`;
        if (precisionValue) {
            columnDef += `(${precisionValue})`;
        }
        
        columns.push(columnDef);
      }
        
    }

    sqlStatement += columns.join(',\n') + '\n);';

    sqlStatement += '\n\nSELECT @queryId = SCOPE_IDENTITY();\n\n';

    sqlStatement += 'END\n\n\nGO\n\n';

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

    if (columnMap.sqlColumn != 0 && columnMap.sqlColumn != "0" && columnMap.sqlColumn != "") {
      sqlMappings.SQLMappings[0].columnsMap.push(columnMap);
    }
    
  }

  return JSON.stringify(sqlMappings, null, 2); // Pretty print the JSON
}



/**
 * Creates an SQL INSERT statement for each mapping entry.
 * @customfunction
 * @description Generates an SQL INSERT statement for Spotify mapping.
 * @param {any} sourceFileName The name of the source file.
 * @param {any} tableName Name of the SQL table for every element.
 * @param {any} columnNames Range of cells for the "sqlColumn" attribute.
 * @param {any} paths Range of cells for the "path" attribute.
 * @param {any} dataTypes Range of cells for the "type" attribute.
 * @param {any} precision Range of cells for the "precision" attribute. 
 */
function generateSQLInsertMap(sourceFileName, tableName, columnNames, paths, dataTypes, precision) {
  // Begin the SQL INSERT statement for a temp table
  let insertStatements = [];
  let maxLength = Math.max(columnNames.length, paths.length, dataTypes.length, precision.length);


  
  insertStatements.push(`INSERT INTO [connect].[SourceQueryMapping]([sourceQueryId],[pathInSource],[targetTable],[targetColumn],[typeDataType],[precision],[nullable],[createdBy]) VALUES `);
  let first = true;

  for (let i = 0; i < maxLength; i++) {
    let sqlColumn = (columnNames[i] && columnNames[i][0]) || null;
    let path = (paths[i] && paths[i][0]) || null;
    let type = (dataTypes[i] && dataTypes[i][0]) || null;
    let precisionValue = (precision[i] && precision[i][0]) || null;
    
    if (sqlColumn && sqlColumn != 0 && sqlColumn != "0" && sqlColumn != "") {
      
      // Construct the INSERT statement
      let insertStatement = `${first ? '' : ','}(@queryId, `;
      insertStatement +=  `${path ? `'${path}'` : "NULL"},`;
      insertStatement +=  `'${tableName}',`;
      insertStatement +=  `${sqlColumn ? `'${sqlColumn}'` : "NULL"}, `;
      insertStatement +=  ` ${type ? `'${type}'` : "NULL"}, `;
      insertStatement +=  `${precisionValue ? `'${precisionValue}'` : "NULL"},`;
      insertStatement +=  '0,';
      insertStatement +=  "'Mapping' )";
        
      insertStatements.push(insertStatement);
      first = false;

    } else { 
      debugger;
    }
  }

  return insertStatements.join('\n');
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



/**
 * Function to handle Table button click.
 * Opens the task pane if not already open and navigates to the Table page.
 */
function onTableButtonClick() {
  Office.context.ui.displayDialogAsync('https://localhost:3000', { width: 50, height: 50 }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      asyncResult.value.messageChild(JSON.stringify({ action: 'navigate', page: '/Table' }));
    } else {
      console.error("Failed to open task pane: ", asyncResult.error.message);
    }
  });
}




CustomFunctions.associate("MAKESQL", makeSQL);
CustomFunctions.associate("JSONMAP", generateJsonMap);
CustomFunctions.associate("INSERTMAP", generateSQLInsertMap);
