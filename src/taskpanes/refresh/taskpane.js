/* 
###########################################################################################
	Init
########################################################################################### 
*/


/**
 * List of Azure functions called for refresh. A function may request one or more Dimensions.
 */
const azureFunctions = [
    { 
        id: 0, 
        name: 'Shopify Products', 
        url: "http://localhost:7071/api/SyncShopify",
        data: {
            "importRequest": [
                {
                    "sourceSystem": "Shopify", 
                    "obj": ["Product", "Sales"]
                }
            ]
        }
    }    
    
];


let completedFunctionsCount = 0;
let errorOccurred = false;


/**
 * Bind the update button to the function call events.
 */
Office.onReady((info) => {
    
    if (info.host === Office.HostType.Excel) {        

        // Bind Refresh of source data
        document.getElementById('startFunctionsBtn').addEventListener('click', function() {
            completedFunctionsCount = 0;
            errorOccurred = false;

            const finalStatus = document.getElementById('finalStatusIndicator')
            if(finalStatus) {
                finalStatus.textContent = "";
            }

            azureFunctions.forEach(functionDetails => {
                startAzureFunction(functionDetails.id);
            });
        });

        // Bind make table
        document.getElementById('btnCreateDimensionTable').addEventListener('click', function() {
          setupProducts();
        });

        document.getElementById('btnAddEntity').addEventListener('click', function() {
            addEntitiesToTable();
        });


        document.getElementById('btnGradient').addEventListener('click', async function() {
          await applyGradient();
      });


    }
});


/* 
###########################################################################################
	API Calls
########################################################################################### 
*/



/**
 * Start an Azure function
 * @param {number} functionId index of the azureFunctions array
 * @returns promise
 */
async function callAzureFunction(functionId) {    
    updateStatus(functionId, 'Running...', 'running');
    return await fetch(azureFunctions[functionId].url, {        
        method: 'POST',
        mode: 'cors',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(azureFunctions[functionId].data)
    });

}


/* 
###########################################################################################
	General Excel Actions
########################################################################################### 
*/

/**
 * Create a new worksheet in workbook. 
 * @param {Excel.RequestContext} context Excel request context
 * @param {string} name New worksheet name
 * @param {boolean} deleteFirst Delete before create
 * @param {boolean} activate Focus after create
 * @returns new worksheet
 */
async function createWorksheet(context, name, deleteFirst = false, activate = true) {
    if (deleteFirst) {
        let worksheet = context.workbook.worksheets.getItemOrNullObject(name);
        worksheet.load('name');
        await context.sync();
        if (worksheet.name) {
            worksheet.delete();
        }
    }

    const sheet = context.workbook.worksheets.add(name);
    if (activate) {
        sheet.activate();
    }
    await context.sync();
    return sheet;
}



/* 
###########################################################################################
	Styles - Generic
########################################################################################### 
*/




async function addNewStyle(styleName, removeFirst) {
	await Excel.run(async (context) => {		
				
		if (removeFirst) {
			// Remove the style with this name if it exists.		
			await removeStyle(styleName); 
		} else if(isStyleName(context, styleName)) {
			// If the style already exists return.
			return context.sync();			
		}
		
		
		// Add a new style to the style collection.
	  	// Styles is in the Home tab ribbon.		
		context.workbook.styles.add(styleName);  
		let newStyle = context.workbook.styles.getItem(styleName);

		// Set Formatting		
		newStyle = removeStyleBorders(newStyle);
		newStyle.includeBorder = true; 		// Set the style as including border information.
		
		if(styleName == "Remind Table Body") {
			newStyle.fill.color = "#900000";
		} else {
			newStyle.fill.color = "#000099";
		}

		newStyle.formulaHidden = false;
		newStyle.locked = false;
		newStyle.shrinkToFit = false;	
		newStyle.textOrientation = 0;		
		newStyle.autoIndent = true;
		newStyle.includeProtection = false;
		newStyle.wrapText = true;

		
	
		console.log("Successfully added a new style with diagonal orientation to the Home tab ribbon.");		
		return context.sync();	

	  });
  	  
}
  

async function removeStyle(styleName) {
    await Excel.run(async (context) => {
        let styles = context.workbook.styles;

        // Get the style if it exists and delete it.
        let style = styles.getItemOrNullObject(styleName);
        await context.sync();

        // Check if the style exists before trying to delete it.
        if (!style.isNullObject) {
            style.delete();            
        }

        await context.sync();
    });
}



async function isStyleName(context, styleName) {    
        let styles = context.workbook.styles;

        // Get the style if it exists and delete it.
        let style = styles.getItemOrNullObject(styleName);
        await context.sync();

        return !style.isNullObject
}


async function removeStyleBorders(style) {
    
	// Check if the style exists before trying to modify it
	if (!style.isNullObject) {
		// Removing all borders from the style
		const borderProperties = {
			style: "None",
			color: "none"
		};

		style.borderTop = borderProperties;
		style.borderLeft = borderProperties;
		style.borderRight = borderProperties;
		style.borderBottom = borderProperties;
		style.borderDiagonal = borderProperties;
		style.borderHorizontal = borderProperties;
		style.borderVertical = borderProperties;

	}
	return style
}


/* 
###########################################################################################
	Styles
########################################################################################### 
*/



/**
 * Apply formatting to the table
 * @param {Excel.Table} table The table to format
 */
async function formatGradientTable(context, table) {
    await cleartableFormat(table.name);
	
	// Format the header row
    const headerRange = table.getHeaderRowRange();
    headerRange.format.fill.color = 'white';  // I dont like this but I can't seem to get clear() to work
    headerRange.format.font.bold = true;      // Example header font style

    /*
	// Format the data rows
    const dataRange = table.getDataBodyRange();
    dataRange.format.fill.color = 'white';	
    dataRange.format.font.name = 'Arial';       
    dataRange.format.font.size = 10;
	*/
	
	headerRange.load(["width", "columnCount"]);	
	
	await context.sync();  

	
	const cells = []
	for (let i=0; i < headerRange.columnCount; i++) {
		cells.push(headerRange.getCell(0,i));				
		cells[i].load('width');
	}

	await context.sync();  


	var startColorRgb = hexToRgb("#FFD700");
	var endColorRgb = hexToRgb("#008080");
				
	let runningTotal = 0;
	// Set a new bottom border style for each column in the header	
	for (let i = 0; i < cells.length; i++) {			
		runningTotal += cells[i].width;	

		let interpolatedColorRgb = interpolateColor(startColorRgb, endColorRgb, runningTotal / headerRange.width);

		cells[i].format.borders.getItem(Excel.BorderIndex.edgeBottom).style = 'Continuous';
		cells[i].format.borders.getItem(Excel.BorderIndex.edgeBottom).color = rgbToHex(interpolatedColorRgb);
		cells[i].format.borders.getItem(Excel.BorderIndex.edgeBottom).weight = 'Medium';
		cells[i].format.fill.color = 'white';	
	}
	
    
}


async function demo_addNewStyle() {
	await Excel.run(async (context) => {
	  let styles = context.workbook.styles;
  
	  // Add a new style to the style collection.
	  // Styles is in the Home tab ribbon.
	  styles.add("RM 3 Style");
  
	  let newStyle = styles.getItem("RM 3 Style");
  
	  // The "Diagonal Orientation Style" properties.
	  newStyle.textOrientation = 90;
	  newStyle.autoIndent = true;
	  newStyle.includeProtection = true;
	  newStyle.shrinkToFit = true;
	  newStyle.locked = false;
  
	  await context.sync();
  
	  console.log("Successfully added a new style with diagonal orientation to the Home tab ribbon.");
	});
  }
  

/**
 * Apply a predefined Style to a table.
 * @param {string} sheetName Name of Worksheet holding the table
 * @param {string} tableName Name of Data Table to be formatted
 * @param {string} headerStyleName Name of Style for Data Table header row.
 * @param {string} bodyStyleName Name of Style for the Data Table body rows.
 * @param {string} totalStyleName Name of Style for the Data Table total row.
 */
async function applyTableStyle(sheetName, tableName, headerStyleName, bodyStyleName, totalStyleName) {

	await Excel.run(async (context) => {
		let sheet = context.workbook.worksheets.getItem(sheetName);
		let table = sheet.tables.getItem(tableName);
		
		table.load(["showTotals"]);	
		await context.sync();


		if(headerStyleName) {
			table.getHeaderRowRange().style = headerStyleName;		
		}
		
		if(bodyStyleName) {
			table.getDataBodyRange().style = bodyStyleName;		
		}
		
		if (table.showTotals && totalStyleName) {
            table.getTotalRowRange().style = totalStyleName;
        }

		await context.sync();
	  });
}



async function tableStyle() {
	Excel.run(function (context) {
		var sheet = context.workbook.worksheets.getItem("Products"); // Replace with your sheet name
		var table = sheet.tables.getItem("ProductsTable"); // Replace with your table name
	
		// Reset table formatting to defaults
		table.style = "TableStyleLigt1";		
	
		return context.sync();
	}).catch(function (error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
	});
	
}

async function cleartableFormat(tableName) {
	await Excel.run(async (context) => {
		let sheet = context.workbook.worksheets.getItem("Products");
		let expensesTable = sheet.tables.getItem("ProductsTable");
	
		expensesTable.getHeaderRowRange().format.fill.color = "#FFFFFF";
		expensesTable.getHeaderRowRange().format.font.color = "#000000";
		expensesTable.getDataBodyRange().format.fill.color = "#FFFFFF";
		
		await context.sync();
	});
	
}



/**
 * Apply a colour gradient to a selected range based on the columns width.
 */
async function applyGradient() { 
  
	Excel.run(function (context) {
		// Get the currently selected range
		const range = context.workbook.getSelectedRange();

		// Load the required range attributes.
		range.load(['rowCount', 'columnCount', 'width']);		
  

		return context.sync().then(function () {
			      						
			// Loop through each column in the range and get the width.
      	const columns = []
		for (let i=0; i < range.columnCount; i++) {
			columns.push(range.getColumn(i));				
			columns[i].load('width');
		}


		return context.sync().then(function () {
				
        let runningTotal = 0;
				
        // Colors in Hex
				var startColorRgb = hexToRgb("#FFD700");
				var endColorRgb = hexToRgb("#008080");
				
        // Loop through the rows and columns in the range.
				for (let row=0; row<range.rowCount; row++ ) {
					for (let col=0; col < range.columnCount; col++ ) {

						const cell = range.getCell(row,col)
						runningTotal += columns[col].width;

						// Find the colour at this cells percentage of the total width.
						let interpolatedColorRgb = interpolateColor(startColorRgb, endColorRgb, runningTotal / range.width);

            			// Set the cell border colour.
						cell.format.borders.getItem('EdgeBottom').style = 'Continuous';
						cell.format.borders.getItem('EdgeBottom').color = rgbToHex(interpolatedColorRgb);
						cell.format.borders.getItem('EdgeBottom').weight = 'Thick';

            
						/*
						//Testing - Set the value of the cell to be it's width
						const value = columns[col].width / range.width * 100;
									cell.values = [[value]]; 		
						*/

					}
				}				
				
				return context.sync();
			});
		});
	});

}

/**
 * Convert a Hex colour value to rgb format.
 * @param {string} hex Hex colour value
 * @returns {string}
 */
function hexToRgb(hex) {
    // Remove the hash at the start if it's there
    hex = hex.replace(/^\s*#|\s*$/g, '');

    // Parse the hex color
    var bigint = parseInt(hex, 16);
    var r = (bigint >> 16) & 255;
    var g = (bigint >> 8) & 255;
    var b = bigint & 255;

    return [r, g, b];
}


/**
 * The rgb value of a point on a gradient between two colours.
 * @param {integer[]} color1 rgb colour array
 * @param {integer[]} color2 rgb colour array
 * @param {number} factor percentage of gradient completion
 * @returns 
 */
function interpolateColor(color1, color2, factor) {
    // Linear interpolation between the color components
    var result = color1.slice();
    for (var i = 0; i < 3; i++) {
        result[i] = Math.round(result[i] + factor * (color2[i] - color1[i]));
    }
    return result;
}

/**
 * Converts an rgb array to a Hex colour string
 * @param {number[]} rgb rgb colour array
 * @returns 
 */
function rgbToHex(rgb) {
    return "#" + rgb.map(function (value) {
        return ("0" + value.toString(16)).slice(-2);
    }).join('');
}




/* 
###########################################################################################
	Refresh
########################################################################################### 
*/




/**
 * Start and Azure function from the array of function calls.
 * @param {number} functionId 
 */
function startAzureFunction(functionId) {
    if (!document.getElementById('statusIndicator' + functionId)) {
        createStatusIndicator(functionId);
    }
    updateStatus(functionId, 'Starting...', 'running');
    callAzureFunction(functionId)
        .then(() => {            
            updateStatus(functionId, 'Completed', 'completed');
        })
        .catch((error) => {            
            updateStatus(functionId, 'Error: ' + error.message, 'error');
            errorOccurred = true;
        })
        .finally(() => {            
            checkAllFunctionsCompleted();
        });
}


/**
 * Create a label on the taskpane to show the status of the function call.
 * @param {number} functionId 
 */
function createStatusIndicator(functionId) {
    const statusIndicators = document.getElementById('statusIndicators');
    const indicator = document.createElement('div');
    indicator.id = 'statusIndicator' + functionId;
    indicator.textContent = `Import ${azureFunctions[functionId].name} Status: Idle`;
    statusIndicators.appendChild(indicator);
}

/**
 * Update the Azure Function status label.
 * @param {number} functionId 
 * @param {string} message New text for the label
 * @param {string} status Status text of icon
 */
function updateStatus(functionId, message, status) {
    const statusIndicator = document.getElementById('statusIndicator' + functionId);
    statusIndicator.innerHTML = `${azureFunctions[functionId].name} Status: <span class="${status}">${message}</span>`;
}



/**
 * We each Azure function promise returns check to see that overall status.
 */
function checkAllFunctionsCompleted() {
    completedFunctionsCount++;
    if (completedFunctionsCount === azureFunctions.length) {
        if (errorOccurred) {
            showFinalStatus('Completed with Errors');
        } else {
            showFinalStatus('All Functions Completed Successfully');
        }
    }
}

/**
 * Update the taskpane with a final status message success/fail.
 * @param {string} message Text or icon html
 */
function showFinalStatus(message) {    
    const finalStatus = document.getElementById('finalStatusIndicator')
    if(finalStatus) {
        finalStatus.textContent = message    
    }
}


/* 
###########################################################################################
	Data Tables
########################################################################################### 
*/

async function addEntitiesToTable() {
  // This function retrieves data for each of the existing products in the table,
  // creates entity values for each of those products, and adds the entities
  // to the table.
  await Excel.run(async (context) => {
    const productsTable = context.workbook.tables.getItem("ProductsTable");

    // Add a new column to the table for the entity values.
    productsTable.columns.getItemOrNullObject("Product").delete();
    const productColumn = productsTable.columns.add(0, null, "Product");

    // Get product data from the table.
    const dataRange = productsTable.getDataBodyRange();
    dataRange.load("values");

    await context.sync();

    // Loop through the rows of the table
    const entities = dataRange.values.map((rowValues) => {
      // Get products and product properties.
      
      const product = getProduct(rowValues[1]);

      // Get product categories and category properties.
      /*const category = product ? getCategory(product.categoryID) : null;

      // Get product suppliers and supplier properties.
      const supplier = product ? getSupplier(product.supplierID) : null;
        */
      // Create entities by combining product, category, and supplier properties.
      return [makeProductEntity(rowValues[1], rowValues[2], product)];
    });
    
    // Add the complete entities to the Products Table.
    productColumn.getDataBodyRange().valuesAsJson = entities;

    productColumn.getRange().format.autofitColumns();
    await context.sync();
  });
}


// Create entities from product properties.
function makeProductEntity(productID, productName, product) {
  const entity = {
    type: Excel.CellValueType.entity,
    text: productName,


    properties: {
      "Code": {
        type: Excel.CellValueType.string,
        basicValue: productID.toString() || ""
      },
      "Product Name": {
        type: Excel.CellValueType.string,
        basicValue: productName || ""
      },
      "Description": {
        type: Excel.CellValueType.string,
        basicValue: product.description || ""
      },
      "Handle": {
        type: Excel.CellValueType.string,
        basicValue: product.handle || ""
      },      
      "Created By": {
        type: Excel.CellValueType.string,
        basicValue: product.createdBy || ""
      },      
      "Created": {
        type: Excel.CellValueType.string,
        basicValue: product.created || ""
      }
      
    },
    layouts: {
      compact: {
        icon: Excel.EntityCompactLayoutIcons.shoppingBag
      },
      card: {
        title: { property: "Product Name" },
        sections: [
          {
            layout: "List",
            properties: ["Code"]
          },
          {
            layout: "List",
            title: "Details",
            collapsible: true,
            collapsed: false,
            properties: ["Description", "Handle"]
          },
          {
            layout: "List",
            title: "Additional information",
            collapsed: true,
            properties: ["Created By", "Created"]
          }
        ]
      }
    }
  };

  // Add image property to the entity and then add it to the card layout.
  /*if (product.productImage) {
    entity.properties["Image"] = {
      type: Excel.CellValueType.webImage,
      address: product.productImage || ""
    };
    entity.layouts.card.mainImage = { property: "Image" };
  }
*/
  // Add a nested entity for the product category.
  /*if (category) {
    entity.properties["Category"] = {
      type: Excel.CellValueType.entity,
      text: category.categoryName,
      properties: {
        "Category ID": {
          type: Excel.CellValueType.double,
          basicValue: category.categoryID,
          propertyMetadata: {
            // Exclude the category ID property from the card view and auto complete.
            excludeFrom: {
              cardView: true,
              autoComplete: true
            }
          }
        },
        "Category Name": {
          type: Excel.CellValueType.string,
          basicValue: category.categoryName || ""
        },
        Description: {
          type: Excel.CellValueType.string,
          basicValue: category.description || ""
        }
      },
      layouts: {
        compact: {
          icon: Excel.EntityCompactLayoutIcons.branch
        }
      }
    };

    // Add nested product category to the card layout.
    entity.layouts.card.sections[0].properties.push("Category");
  }
*/
  // Add a nested entity for the supplier.
 /* if (supplier) {
    entity.properties["Supplier"] = {
      type: Excel.CellValueType.entity,
      text: supplier.companyName,
      properties: {
        "Supplier ID": {
          type: Excel.CellValueType.double,
          basicValue: supplier.supplierID
        },
        "Company Name": {
          type: Excel.CellValueType.string,
          basicValue: supplier.companyName || ""
        },
        "Contact Name": {
          type: Excel.CellValueType.string,
          basicValue: supplier.contactName || ""
        },
        "Contact Title": {
          type: Excel.CellValueType.string,
          basicValue: supplier.contactTitle || ""
        }
      },
      layouts: {
        compact: {
          icon: Excel.EntityCompactLayoutIcons.boxMultiple
        },
        card: {
          title: { property: "Company Name" },
          sections: [
            {
              layout: "List",
              properties: ["Supplier ID", "Company Name", "Contact Name", "Contact Title"]
            }
          ]
        }
      }
    };

    // Add nested product supplier to the card layout.
    entity.layouts.card.sections[2].properties.push("Supplier");
  }
  */
  return entity;
}


// Get products and product properties.
function getProduct(id) {
  return shopifyProducts.find((p) => p.primarySystemCode == id);
}


// Get product categories and category properties.
function getCategory(categoryID) {
  return categories.find((c) => c.categoryID == categoryID);
}


// Get product suppliers and supplier properties.
function getSupplier(supplierID) {
  return suppliers.find((s) => s.supplierID == supplierID);
}



/**
 * Create and populate a new data table
 * @param {Excel.RequestContext} context Context of the Excel request
 * @param {Excel.Worksheet} worksheet Target worksheet object
 * @param {string} range Cell range "A1:C3"
 * @param {string} name New table name
 * @param {string[]} columns Array of column names
 * @param {any[]} rows Array of objects representing rows
 * @returns table
 */
async function createDataTable(context, worksheet, range, name, columns, rows) {
	const tbl = worksheet.tables.add(range, true /*hasHeaders*/);
	tbl.name = name;

	tbl.getHeaderRowRange().values = [columns];

	rows.forEach((r) => {
		let rowData = columns.map((c) => r[c]);
		tbl.rows.add(null /*add at the end*/, [rowData]);
	});

	worksheet.getUsedRange().format.autofitColumns();
	worksheet.getUsedRange().format.autofitRows();
	
	await context.sync(); 	
	
	return tbl;
}





/** Set up Sample worksheet. */
async function setupProducts() {

	await getShopifyProducts();

	await Excel.run(async (context) => {
		
		const sheet = await createWorksheet(context, "Products", true, true);
		const productsTable = await createDataTable(context, sheet, "A1:C1", "ProductsTable", ["Product", "primarySystemCode", "memberCaption"], shopifyProducts);
		

		sheet.getUsedRange().format.autofitColumns();
		sheet.getUsedRange().format.autofitRows();

		sheet.activate();

		await context.sync();

	});


	//await cleartableFormat("ProductsTable");	
	debugger;
	await addNewStyle("Remind Table Header", true);
	await addNewStyle("Remind Table Body", true);	
	applyTableStyle("Products", "ProductsTable", "Remind Table Header", "Remind Table Body")
}




/**
 * Start an Azure function for Dimension Query
 * @returns promise
 */
async function getShopifyProducts() {  
    
    const response = await fetch("http://localhost:7071/api/DimensionQuery", {        
        method: 'POST',
        mode: 'cors',
        headers: {
            'Content-Type': 'application/json',
        },
        body: `{
            "query": [
                {
                    "dimension": "Product"
                    "filters": []
                }
            ]
        }
        `
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();

    if (data && data[0] && data[0].result) {
      const j = JSON.parse(data[0].result)
          
      shopifyProducts.splice(0, shopifyProducts.length); // Remove all elements from the array
      shopifyProducts.push(...j); // Merge arrays
    }
    

}



/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

const shopifyProducts = []




/** Sample JSON product data. */
const products = [
  {
    productID: 1,
    productName: "Chai",
    supplierID: 1,
    categoryID: 1,
    quantityPerUnit: "10 boxes x 20 bags",
    unitPrice: 18,
    discontinued: false,
    productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/0/04/Masala_Chai.JPG/320px-Masala_Chai.JPG"
  },
  {
    productID: 2,
    productName: "Chang",
    supplierID: 1,
    categoryID: 1,
    quantityPerUnit: "24 - 12 oz bottles",
    unitPrice: 19,
    discontinued: false,
    productImage: ""
  },
  {
    productID: 3,
    productName: "Aniseed Syrup",
    supplierID: 1,
    categoryID: 2,
    quantityPerUnit: "12 - 550 ml bottles",
    unitPrice: 10,
    discontinued: false,
    productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/8/81/Maltose_syrup.jpg/185px-Maltose_syrup.jpg"
  },
  {
    productID: 4,
    productName: "Chef Anton's Cajun Seasoning",
    supplierID: 2,
    categoryID: 2,
    quantityPerUnit: "48 - 6 oz jars",
    unitPrice: 22,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/8/82/Kruidenmengeling-spice.jpg/193px-Kruidenmengeling-spice.jpg"
  },
  {
    productID: 5,
    productName: "Chef Anton's Gumbo Mix",
    supplierID: 2,
    categoryID: 2,
    quantityPerUnit: "36 boxes",
    unitPrice: 21.35,
    discontinued: true,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/Okra_in_a_Bowl_%28Unsplash%29.jpg/180px-Okra_in_a_Bowl_%28Unsplash%29.jpg"
  },
  {
    productID: 6,
    productName: "Grandma's Boysenberry Spread",
    supplierID: 3,
    categoryID: 2,
    quantityPerUnit: "12 - 8 oz jars",
    unitPrice: 25,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/1/10/Making_cranberry_sauce_-_in_the_jar.jpg/90px-Making_cranberry_sauce_-_in_the_jar.jpg"
  },
  {
    productID: 7,
    productName: "Uncle Bob's Organic Dried Pears",
    supplierID: 3,
    categoryID: 7,
    quantityPerUnit: "12 - 1 lb pkgs.",
    unitPrice: 30,
    discontinued: false,
    productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/f/fd/DriedPears.JPG/120px-DriedPears.JPG"
  },
  {
    productID: 8,
    productName: "Northwoods Cranberry Sauce",
    supplierID: 3,
    categoryID: 2,
    quantityPerUnit: "12 - 12 oz jars",
    unitPrice: 40,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/0/07/Making_cranberry_sauce_-_stovetop.jpg/90px-Making_cranberry_sauce_-_stovetop.jpg"
  },
  {
    productID: 9,
    productName: "Mishi Kobe Niku",
    supplierID: 4,
    categoryID: 6,
    quantityPerUnit: "18 - 500 g pkgs.",
    unitPrice: 97,
    discontinued: true,
    productImage: ""
  },
  {
    productID: 10,
    productName: "Ikura",
    supplierID: 4,
    categoryID: 8,
    quantityPerUnit: "12 - 200 ml jars",
    unitPrice: 31,
    discontinued: false,
    productImage: ""
  },
  {
    productID: 11,
    productName: "Queso Cabrales",
    supplierID: 5,
    categoryID: 4,
    quantityPerUnit: "1 kg pkg.",
    unitPrice: 21,
    discontinued: false,
    productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/9/96/Tilsit_cheese.jpg/190px-Tilsit_cheese.jpg"
  },
  {
    productID: 12,
    productName: "Queso Manchego La Pastora",
    supplierID: 5,
    categoryID: 4,
    quantityPerUnit: "10 - 500 g pkgs.",
    unitPrice: 38,
    discontinued: false,
    productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/5/59/Manchego.jpg/177px-Manchego.jpg"
  },
  {
    productID: 13,
    productName: "Konbu",
    supplierID: 6,
    categoryID: 8,
    quantityPerUnit: "2 kg box",
    unitPrice: 6,
    discontinued: false,
    productImage: ""
  },
  {
    productID: 14,
    productName: "Tofu",
    supplierID: 6,
    categoryID: 7,
    quantityPerUnit: "40 - 100 g pkgs.",
    unitPrice: 23.25,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e5/Korean.food-Dubu.gui-01.jpg/120px-Korean.food-Dubu.gui-01.jpg"
  },
  {
    productID: 15,
    productName: "Genen Shouyu",
    supplierID: 6,
    categoryID: 2,
    quantityPerUnit: "24 - 250 ml bottles",
    unitPrice: 15.5,
    discontinued: false,
    productImage: ""
  },
  {
    productID: 16,
    productName: "Pavlova",
    supplierID: 7,
    categoryID: 3,
    quantityPerUnit: "32 - 500 g boxes",
    unitPrice: 17.45,
    discontinued: false,
    productImage: ""
  },
  {
    productID: 17,
    productName: "Alice Mutton",
    supplierID: 7,
    categoryID: 6,
    quantityPerUnit: "20 - 1 kg tins",
    unitPrice: 39,
    discontinued: true,
    productImage: ""
  },
  {
    productID: 18,
    productName: "Carnarvon Tigers",
    supplierID: 7,
    categoryID: 8,
    quantityPerUnit: "16 kg pkg.",
    unitPrice: 62.5,
    discontinued: false,
    productImage: ""
  },
  {
    productID: 19,
    productName: "Teatime Chocolate Biscuits",
    supplierID: 8,
    categoryID: 3,
    quantityPerUnit: "10 boxes x 12 pieces",
    unitPrice: 9.2,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/d/df/Macau_Koi_Kei_Bakery_Almond_Biscuits_2.JPG/120px-Macau_Koi_Kei_Bakery_Almond_Biscuits_2.JPG"
  },
  {
    productID: 20,
    productName: "Sir Rodney's Marmalade",
    supplierID: 8,
    categoryID: 3,
    quantityPerUnit: "30 gift boxes",
    unitPrice: 81,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/3/30/Homemade_marmalade%2C_England.jpg/135px-Homemade_marmalade%2C_England.jpg"
  }
];

const categories = [
  {
    categoryID: 1,
    categoryName: "Beverages",
    description: "Soft drinks, coffees, teas, beers, and ales"
  },
  {
    categoryID: 2,
    categoryName: "Condiments",
    description: "Sweet and savory sauces, relishes, spreads, and seasonings"
  },
  {
    categoryID: 3,
    categoryName: "Confections",
    description: "Desserts, candies, and sweet breads"
  },
  {
    categoryID: 4,
    categoryName: "Dairy Products",
    description: "Cheeses"
  },
  {
    categoryID: 5,
    categoryName: "Grains/Cereals",
    description: "Breads, crackers, pasta, and cereal"
  },
  {
    categoryID: 6,
    categoryName: "Meat/Poultry",
    description: "Prepared meats"
  },
  {
    categoryID: 7,
    categoryName: "Produce",
    description: "Dried fruit and bean curd"
  },
  {
    categoryID: 8,
    categoryName: "Seafood",
    description: "Seaweed and fish"
  }
];

const suppliers = [
  {
    supplierID: 1,
    companyName: "Exotic Liquids",
    contactName: "Charlotte Cooper",
    contactTitle: "Purchasing Manager"
  },
  {
    supplierID: 2,
    companyName: "New Orleans Cajun Delights",
    contactName: "Shelley Burke",
    contactTitle: "Order Administrator"
  },
  {
    supplierID: 3,
    companyName: "Grandma Kelly's Homestead",
    contactName: "Regina Murphy",
    contactTitle: "Sales Representative"
  },
  {
    supplierID: 4,
    companyName: "Tokyo Traders",
    contactName: "Yoshi Nagase",
    contactTitle: "Marketing Manager",
    address: "9-8 Sekimai Musashino-shi"
  },
  {
    supplierID: 5,
    companyName: "Cooperativa de Quesos 'Las Cabras'",
    contactName: "Antonio del Valle Saavedra",
    contactTitle: "Export Administrator"
  },
  {
    supplierID: 6,
    companyName: "Mayumi's",
    contactName: "Mayumi Ohno",
    contactTitle: "Marketing Representative"
  },
  {
    supplierID: 7,
    companyName: "Pavlova, Ltd.",
    contactName: "Ian Devling",
    contactTitle: "Marketing Manager"
  },
  {
    supplierID: 8,
    companyName: "Specialty Biscuits, Ltd.",
    contactName: "Peter Wilson",
    contactTitle: "Sales Representative"
  }
];
