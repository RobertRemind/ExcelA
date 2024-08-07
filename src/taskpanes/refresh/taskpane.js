
/* 
###########################################################################################
	Global Objects - Essentailly our types.
########################################################################################### 
*/


/**
 * List of Azure functions called for refresh. A function may request one or more Dimensions.
 */
const AzureFunctions = [
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


/**
 * Object to hold the user's visual identity settings.
 */
const VisualStyle = {
	colors: {
		startColor: "#FFD700",
		endColor: "#008080",
		startRGB: function(){ return hexToRgb(this.startColor) },
		endRGB: function(){  return hexToRgb(this.endColor) }
	}
	
		
}

/**
 * Status of refreshing data async process.
 */
const RefreshStatus = {
	completedFunctionsCount: 0,
	errorOccurred: false, 
	reset: function () {
		this.completedFunctionsCount = 0;
		this.errorOccurred = false;
	}
}


/**
 * The list of styles created by the addin and how they should look.
 * 
 * To be imported from an Azure function, except for defaults which will be in this code.
 */
const TrackedStyles = { 
	styles: {			
		defaultTableHeader: {
			name: "Remind Table Header",
			format: { // Set the style as excluding border and fill information.
				includeBorder: true, 						
				formulaHidden: false,
				locked:  false,
				shrinkToFit:  false,
				textOrientation:  0,
				autoIndent: true,
				includeProtection: false,
				wrapText: true,
			}, 
			fill: {
				color: "", 
				isClear: true
			}, 
			borders: {}
		}, 
		defaultTableBody: {
			name: "Remind Table Body",
			format: { // Set the style as excluding border and fill information.
				includeBorder: true, 						
				formulaHidden: false,
				locked:  false,
				shrinkToFit:  false,
				textOrientation:  0,
				autoIndent: true,
				includeProtection: false,
				wrapText: true,
			}, 
			fill: {
				color: "", 
				isClear: true
			}, 
			borders: {}
		}
	}, 
	defaults: {
		format: { // Set the style as excluding border and fill information.
			includeBorder: true, 		
			includePatterns: true,
			formulaHidden: false,
			locked:  false,
			shrinkToFit:  false,
			textOrientation:  0,
			autoIndent: true,
			includeProtection: false,
			wrapText: true,
		},
		borders: {
			positions: [
				"borderTop",
				"borderLeft",
				"borderRight",
				"borderBottom",
				"borderDiagonal",
				"borderHorizontal",
				"borderVertical"
			], 
			format: {
				style: "None",
				color: "none"
			}
		}
	}
}


/**
 * A list to tables that have been made by the add-in.
 */
const TrackedTables = {
	tables: [ ]
}


const TablesLibrary = {
	tables: [
		{
			name: "ProductsTable",
			worksheet: "Product", //!!! Testing
			dimension: "products",
			range: "A1:C1",
			styles: {
				header: "defaultTableHeader", 
				body: "defaultTableBody"
			}, 
			columns: ["Product", "Primary System Code", "Member Caption"], // All the columns in the table
			trackedColumns: [
				{
					name: "Product", // What is the name of the column in the data table. 
					source: null, // What is the Data source column name?
					isDirty: false	 // Used to tag that the column has been changed and prompt the user to disable tracking. 
				}, 
				{
					name: "Primary System Code",
					source: "primarySystemCode", 
					isDirty: false	
				}, 
				{
					name: "Member Caption",
					source: "memberCaption", 
					isDirty: false	
				}
			], 
			query: {
				url: "http://localhost:7071/api/DimensionQuery", 
				ajax: {
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
				}				
			},
			rows: []
		}
	]
}


const UIState = {
	infoPanel: {
		trackedTableName: ""
	}
}


/* 
###########################################################################################
	Init
########################################################################################### 
*/


/**
 * Bind the update button to the function call events.
 */
Office.onReady((info) => {
    	
    if (info.host === Office.HostType.Excel) {        		

		initTrackedStyles();	

        // Bind Refresh of source data
        document.getElementById('startFunctionsBtn').addEventListener('click', function() {
			      RefreshStatus.reset();
            RefreshStatus.completedFunctionsCount = 0;
            RefreshStatus.errorOccurred = false;

            const finalStatus = document.getElementById('finalStatusIndicator')
            if(finalStatus) {
                finalStatus.textContent = "";
            }

            AzureFunctions.forEach(functionDetails => {
                startAzureFunction(functionDetails.id);
            });
        });

        
        document.getElementById('btnAddEntity').addEventListener('click', function() {
            addEntitiesToTable();
        });


        document.getElementById('btnGradient').addEventListener('click', async function() {
          await applyGradientToRange();
      	});

		document.getElementById('btnSaveState').addEventListener('click', async function() {
			await saveState(States.TrackedTables);
		});

		document.getElementById('btnGetState').addEventListener('click', async function() {
			await getState(States.TrackedTables);
		});

		document.getElementById('btnListState').addEventListener('click', async function() {
			await listStates();
		});

		document.getElementById('btnClearState').addEventListener('click', async function() {
			await clearState("all");
		});
		
		document.getElementById('btnSyncTables').addEventListener('click', async function() {
			syncTrackedTableInfo();
		});


		document.getElementById('btnDisplayMappingObject').addEventListener('click', async function() {
			
			// Example JSON object
			/*const jsonObject = {
				"id": "gid://shopify/Order/6002828705854",
				"name": "#1002",  
				"subtotalPriceSet": {
					"shopMoney": {
						"amount": "2620.0",
						"currencyCode": "AUD"
					},
					"presentmentMoney": {
						"amount": "2620.0",
						"currencyCode": "AUD"
					}
				}, 
				"tags": [ "a", "b","c" ], 
				"altTag": [ 
					{"label": {"text": "A", "colour": "blue"}, "shape": "rounded"}, 
					{"label": {"text": "B", "colour": "red"}, "shape": "square"}
				]
			
			}
			*/
			const txtArea = document.getElementById('txtMappingObject');

			  displayMappingObject(JSON.parse(txtArea.value));
		});
		

		

		// Get the states from the workbook Custom XML Parts
		getAllStates().then(() => {
			// Then bind events to the tracked tables
			bindAllTrackedTableEvents();					
		});

		// Make the Library controls and bind events table
		populateLibraryDropDown();
        document.getElementById('btnLibTableWorksheet').addEventListener('click', function() {						
			createTrackedTable(document.getElementById('tableLibrarySelect').value, true);
			
		});

		document.getElementById('btnLibTableSelectedCell').addEventListener('click', function() {
			createTrackedTable(document.getElementById('tableLibrarySelect').value, false);			
		});

    }
});


/* 
###########################################################################################
	UI 
########################################################################################### 
*/

function populateLibraryDropDown() {
	// Try to get an existing select element with the ID 'slTableList'
	let dropdown = document.querySelector('#tableLibrarySelect');

	// Adding options to the dropdown
	TablesLibrary.tables.forEach(table => {
		const option = document.createElement('option');
		option.value = table.name;
		option.text = table.name;
		dropdown.appendChild(option);
	});
}


/* 
###########################################################################################
	API Calls
########################################################################################### 
*/



/**
 * Start an Azure function
 * @param {number} functionId index of the AzureFunctions array
 * @returns promise
 */
async function callAzureFunction(functionId) {    
    updateStatus(functionId, 'Running...', 'running');
    return await fetch(AzureFunctions[functionId].url, {        
        method: 'POST',
        mode: 'cors',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(AzureFunctions[functionId].data)
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


/**
 * Get a list of standard types styles. Apply style with setTableStyle()
 */
async function listAllTableStyles() {
    await Excel.run(async (context) => {
        const workbook = context.workbook;
        const styles = workbook.tableStyles;
        styles.load('items/name');

        await context.sync();

        const styleNames = styles.items.map(style => style.name);
        console.log("Available Table Styles:", styleNames);
    }).catch(error => {
        console.error(error);
    });
}

/**
 * Apply a style to a data table.
 * @param {string} tableName Data Table Name
 * @param {string} styleName Style name see listAllTableStyles()
 */
async function setTableStyle(tableName, styleName) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.getItem(tableName);

        // Apply the custom style to the table
        table.style = styleName;

        await context.sync();
        console.log(`Custom style '${styleName}' applied to table '${tableName}'.`);
    }).catch(error => {
        console.error(error);
    });
}


/**
 * Check if two Excel range strings intersect.
 * @param {string} range1 - First range string (e.g., "'Sheet1'!A1:B2")
 * @param {string} range2 - Second range string (e.g., "'Sheet1'!C3:D4")
 * @returns {boolean} - True if the ranges intersect, false otherwise.
 */
function doRangesIntersect(range1, range2) {
    
    const range1Parsed = parseRange(range1);
    const range2Parsed = parseRange(range2);

    // Convert column letters to numbers
    range1Parsed.startColumn = columnToNumber(range1Parsed.startColumn);
    range1Parsed.endColumn = columnToNumber(range1Parsed.endColumn);
    range2Parsed.startColumn = columnToNumber(range2Parsed.startColumn);
    range2Parsed.endColumn = columnToNumber(range2Parsed.endColumn);

    // Check for intersection
    const rowsIntersect = range1Parsed.startRow <= range2Parsed.endRow && range1Parsed.endRow >= range2Parsed.startRow;
    const colsIntersect = range1Parsed.startColumn <= range2Parsed.endColumn && range1Parsed.endColumn >= range2Parsed.startColumn;

    return rowsIntersect && colsIntersect;
}


/**
 * Helper function to parse a range string into row and column bounds
 * @param {string} rangeString Excel Range as string
 * @returns 
 */
function parseRange(rangeString) {
	const sheetRegex = /^'?(.+?)'?!(.+)$/;
	const match = rangeString.match(sheetRegex);
	const rangePart = match[2];

	const cellRegex = /([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/;
	const cells = rangePart.match(cellRegex);

	let startColumn = cells[1];
	let startRow = parseInt(cells[2], 10);
	let endColumn = cells[3] ? cells[3] : startColumn;
	let endRow = cells[4] ? parseInt(cells[4], 10) : startRow;

	return {
		startRow: startRow,
		endRow: endRow,
		startColumn: startColumn,
		endColumn: endColumn
	};
}



/**
 * Convert a column letter (e.g., "AA") to a number (e.g., 27)
 * @param {string} column Excel column name to number
 * @returns int
 */
function columnToNumber(column) {
	let sum = 0;
	for (let i = 0; i < column.length; i++) {
		sum *= 26;
		sum += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
	}
	return sum;
}


/**
 * Splits an Excel address into its worksheet and range parts.
 * If the worksheet name is not included in the address, the function returns only the range part.
 * @param {string} address - The full Excel address (e.g., "Product!D22:F23" or "D22:F23").
 * @returns {Object} An object with 'worksheet' and 'range' properties. 'worksheet' may be undefined if not included.
 */
function splitExcelAddress(address) {
    if (address.includes('!')) {
        const [worksheet, range] = address.split('!');
        return { worksheet, range };
    } else {
        return { worksheet: undefined, range: address };
    }
}



async function doesTableExist(context, tableName) {
	const worksheets = context.workbook.worksheets;
	worksheets.load("items/name"); // Load the collection of worksheets
	await context.sync();

	for (let i = 0; i < worksheets.items.length; i++) {
		const worksheet = worksheets.items[i];
		const tables = worksheet.tables;
		tables.load("items/name"); // Load the names of tables
		await context.sync();

		const tableExists = tables.items.some(table => table.name === tableName);
		if (tableExists) {
			return true; // Table found
		}
	}

	return false; // Table not found in any worksheet
}



/* 
###########################################################################################
	Styles
########################################################################################### 
*/


async function initTrackedStyles() {

    const stylePromises = Object.keys(TrackedStyles.styles).map(key => {
        return syncTrackedStyle(TrackedStyles.styles[key], true);
    });

    await Promise.all(stylePromises);
}


/**
 * Add a Tracked Style to the Excel workbook.
 * @param {TrackedStyle} trackedStyle Tracked Style to be added.
 * @param {boolean} sync sync the Excel settings with the tracked setttings. Only required where the tracked setting may have changed. 
 */
async function syncTrackedStyle(trackedStyle, sync) {
	
	await Excel.run(async (context) => {		
				
		if (sync) {						
			await removeStyle(trackedStyle.name); 	// Remove the style with this name if it exists.		
		} else if(isStyleName(context, trackedStyle.name)) {			
			return context.sync();	// If the style already exists return.
		}
		
		
		// Add a new style to the style collection.
	  	// Styles is in the Home tab ribbon.		
		context.workbook.styles.add(trackedStyle.name);  
		let newStyle = context.workbook.styles.getItem(trackedStyle.name);

		// Set borders		
		newStyle = syncStyleBorders(newStyle, trackedStyle);
		newStyle = syncStyleFill(newStyle, trackedStyle);
		newStyle = syncStyleFormat(newStyle, trackedStyle);
		
	
		return context.sync();	

	});
  	  
}
  

/**
 * Remove and Excel Style from the current context.
 * @param {string} styleName Name of a Style
 */
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


/**
 * Check if a Style name exists in this Context.
 * @param {Excel.context} context Current Excel Context
 * @param {string} styleName The name of an Excel Style to text
 * @returns boolean
 */
async function isStyleName(context, styleName) {    
        let styles = context.workbook.styles;

        // Get the style if it exists and delete it.
        let style = styles.getItemOrNullObject(styleName);
        await context.sync();

        return !style.isNullObject
}


/**
 * Apply border settings to an Excel Style
 * @param {Excel.Style} style Excel Style to be sync'd
 * @param {*} trackedStyle Tracked Style settings
 * @returns Excel.Style
 */
function syncStyleBorders(style, trackedStyle) {
    
	// Apply Borders 
	if (style && trackedStyle && trackedStyle.borders) {
		style.includeBorder = trackedStyle.borders.includeBorder ? trackedStyle.borders.includeBorder : true; 		// Set the style as including border information.

		//Apply defaults
		TrackedStyles.defaults.borders.positions.map((pos) => {
			style[pos] = TrackedStyles.defaults.borders.format;
		});

		// loop through border settings and apply
		Object.keys(trackedStyle.borders).forEach(key => {
			style[key] = trackedStyle.borders[key];
		});			
	}
	return style;

}


/**
 * Apply fill settings to an Excel Style
 * @param {Excel.Style} style Excel Style to be sync'd
 * @param {*} trackedStyle Tracked Style settings
 * @returns Excel.Style
 */
function syncStyleFill(style, trackedStyle) {
	
	// Apply fill
	if (style && trackedStyle && trackedStyle.fill) {
		style.includePatterns = true;		
		
		if (trackedStyle.fill.isClear && style.fill) {
			style.fill.clear();

		} else {
			if (style.fill) {
				if (trackedStyle.fill.color) {
					style.fill.color = trackedStyle.fill.color;		
				} else {
					style.fill.clear();			
				}				
			}			
		}	

	}
	return style;
}




/**
 * Apply general format settings to an Excel Style
 * @param {Excel.Style} style Excel Style to be sync'd
 * @param {*} trackedStyle Tracked Style settings
 * @returns Excel.Style
 */
function syncStyleFormat(style, trackedStyle) {
	
	if (style && trackedStyle && trackedStyle.format) {
	
		// Assign the style settings from tracked styles is it has been defined. Otherwise use the defaults.
		const defaultFormat = TrackedStyles.defaults.format;
		Object.keys(trackedStyle.format).forEach(key => {
			style[key] = trackedStyle[key] === undefined ? trackedStyle[key] : defaultFormat[key];
		});

	}			
	
	return style;
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
	
	// Format the header row
    const headerRange = table.getHeaderRowRange();    
	headerRange.load(["width", "columnCount"]);	
	
	await context.sync();  
	
	const cells = []
	for (let i=0; i < headerRange.columnCount; i++) {
		cells.push(headerRange.getCell(0,i));				
		cells[i].load('width');
	}

	await context.sync();  


	var startColorRgb = VisualStyle.colors.startRGB();
	var endColorRgb = VisualStyle.colors.endRGB();
				
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

/**
 * Apply a TableStyle to a DataTable.
 * @param {string} tableName Target Data Table name
 * @param {string} tableStyleName Target TableStyle name
 */
async function applyCustomStyleToTable(tableName, tableStyleName) {
	// TableStyles are different to a Style. TableStyles are found in Excels "Table Format" Ribbion. They control the table style default formatting. 
	// Currently the Excel API does not allow table styles to be created or managed. 
	
    await Excel.run(async (context) => {
        
		const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.getItem(tableName);
        
        table.style = styleName;
        await context.sync();        

    }).catch(error => {
        console.error(error);
    });
}


/**
 * Apply a Tracked Style to a Data Table.
 * @param {string} sheetName Name of Worksheet holding the table
 * @param {string} tableName Name of Data Table to be formatted
 * @param {string} headerStyleName Name of Style for Data Table header row.
 * @param {string} bodyStyleName Name of Style for the Data Table body rows.
 * @param {string} totalStyleName Name of Style for the Data Table total row.
 */
async function applyStyleToTable(trackedTable) {

	await Excel.run(async (context) => {

		let sheet = context.workbook.worksheets.getItem(trackedTable.worksheet);
		let table = sheet.tables.getItem(trackedTable.name);
		
		// Get table info.
		table.load(["showTotals", "style"]);			
		await context.sync();

		table.style = null; // Remove the default TableStyle ( not Tracked Style )
		table.showBandedRows = false;
		
		if(trackedTable.styles.body) {
			const bodyStyle = TrackedStyles.styles[trackedTable.styles.body].name
			table.getRange().style = bodyStyle; // format the whole table to match the body.
			table.getDataBodyRange().style =  bodyStyle;		
		}

		if(trackedTable.styles.header) {
			table.getHeaderRowRange().style = TrackedStyles.styles[trackedTable.styles.header].name;		
		}

		if (table.showTotals && trackedTable.styles.total) {
			table.getTotalRowRange().style = TrackedStyles.styles[trackedTable.styles.total].name;
		}
		
		await context.sync();

		await formatGradientTable(context,table);

	});
}


/**
 * Apply a colour gradient to a selected range based on the columns width.
 */
async function applyGradientToRange() { 
  
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
 * The rgb value of a point on a gradient between two colors.
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
            RefreshStatus.errorOccurred = true;
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
    indicator.textContent = `Import ${AzureFunctions[functionId].name} Status: Idle`;
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
    statusIndicator.innerHTML = `${AzureFunctions[functionId].name} Status: <span class="${status}">${message}</span>`;
}



/**
 * We each Azure function promise returns check to see that overall status.
 */
function checkAllFunctionsCompleted() {
    RefreshStatus.completedFunctionsCount++;
    if (RefreshStatus.completedFunctionsCount === AzureFunctions.length) {
        if (RefreshStatus.errorOccurred) {
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
  return TrackedTables.tables[0].rows.find((p) => p.primarySystemCode == id);
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
 * Create and populate a new Tracked Table
 * @param {string} libraryTableName Name of the new library table to create.
 * @param {boolean} newWorksheet Create on a new worksheet
 */
async function createTrackedTable(libraryTableName, newWorksheet) {
		
	const tableSettings = TablesLibrary.tables.find((lib) => {return lib.name === libraryTableName});
	
	// Deep copy the Library object to theTtracked Tables. 
	const clone = JSON.parse(JSON.stringify(tableSettings));
	const i = TrackedTables.tables.push(clone);
	let indexOfNewElement = i - 1;

  //!!!!!! Dont await here track these two awaits as promises.
	await getTrackedTableData(TrackedTables.tables[indexOfNewElement]);

	// Make the new Excel table.
	await Excel.run(async (context) => {
		
		if(newWorksheet) {
			const sheet = await createWorksheet(context, TrackedTables.tables[indexOfNewElement].worksheet, true, true);		
			sheet.activate();
		} else {			
			// If inserting at the selected cell, calc the range that at the table header will require.
			TrackedTables.tables[indexOfNewElement] = await calculateTableTargetRange(context, TrackedTables.tables[indexOfNewElement])			
		}				

		// Generate a new table name if the default table name is in use.
		TrackedTables.tables[indexOfNewElement].name = await generateTableName(context, TrackedTables.tables[indexOfNewElement].name);					

		// Create the table on worksheet
		const wsTable = await createWorksheetTable(context, TrackedTables.tables[indexOfNewElement]);
		TrackedTables.tables[indexOfNewElement].id = wsTable.id;
		saveState(States.TrackedTables);

		await context.sync();		

		// Apply format.	
		applyStyleToTable(TrackedTables.tables[indexOfNewElement]);

		return TrackedTables.tables[indexOfNewElement];
	});	
	
}

/**
 * Get the selected cells and calculate a range wide enough to hold the new table.
 * @param {*} context 
 * @param {TrackedTable} TrackedTable Table to update
 * @returns 
 */
async function calculateTableTargetRange(context, TrackedTable) {	
	
	// From the top left selected cell, define a range wide enough for the column array.
	const selectedCells = context.workbook.getSelectedRange();
	const topLeftCell = selectedCells.getCell(0, 0);				
	const overrideRange = topLeftCell.getResizedRange(0, TrackedTable.columns.length - 1);		
	const sheet = context.workbook.worksheets.getActiveWorksheet();
	
	overrideRange.load(["address"]);
	sheet.load(["name"]);

	await context.sync();

	// Update the Tracked Table definition with the new range. 
	TrackedTable.range = overrideRange.address;				
	TrackedTable.worksheet = sheet.name

	return TrackedTable;		
}




/**
 * Creates a new table name generated from a string. 
 * If a table existed with the proposed name N+1 is added as a suffix.  
 * @param {*} context 
 * @param {string} tableName Proposed table name
 * @returns 
 */
async function generateTableName(context, tableName) {
	let suffix = 1;
	let newTableName = tableName;

	while (await doesTableExist(context, newTableName)) {
		newTableName = `${tableName}${suffix}`
		suffix = suffix + 1;  
	}
	return newTableName;
}



/**
 * Create and populate a new data table on the passed worksheet
 * @param {Excel.RequestContext} context Context of the Excel request
 * @param {TrackedTable} trackedTable Settings of the Tracked Table.
 * @returns 
 */
async function createWorksheetTable(context, trackedTable) {
	
	const worksheet = context.workbook.worksheets.getItemOrNullObject(trackedTable.worksheet);	
	const table = worksheet.tables.add(trackedTable.range, true /*hasHeaders*/);	
	table.name = trackedTable.name;

	// Populate table from Tracked Settings.
	setTrackedTableHeader(table, trackedTable);
	setTrackedTableRows(table, trackedTable);

	// Bind Table Change Event
	handleBindTrackedTableEvents(context, trackedTable);

	// Auto fit new data. This is used by the gradient to determine colors.
	table.getRange().format.autofitColumns();
	table.getRange().format.autofitRows();

	await context.sync(); 	
	
	return table;
}

/**
 * Update an Excel table header with the settings of a Tracked Table
 * @param {Excel.Table} table Excel Table to be updated
 * @param {TrackedTable} trackedTable Tracked Table definition to match
 */
function setTrackedTableHeader(table, trackedTable) {
	const headerValues = [];
	trackedTable.trackedColumns.map((c) => {
		headerValues.push(c.name);
	});
	table.getHeaderRowRange().values = [headerValues];
}


/**
 * Update an Excel table body with the settings of a Tracked Table
 * @param {Excel.Table} table Excel Table to be updated
 * @param {TrackedTable} trackedTable Tracked Table definition to match
 */
function setTrackedTableRows(table, trackedTable) {
	trackedTable.rows.forEach((r) => {
		let rowData = trackedTable.trackedColumns.map((c) => {
			if(c.source) {
				return r[c.source]
			} 
			return null;
		});
		table.rows.add(null /*add at the end*/, [rowData]);
	});
}


/**
 * Loop through Tracked Tables and bind the events.
 */
async function bindAllTrackedTableEvents() {
	await Excel.run(async (context) => {
		TrackedTables.tables.forEach((tbl) => {
			handleBindTrackedTableEvents(context, tbl);
		});
		await context.sync();
	});
}

/**
 * Bind change events to a Tracked Table.
 * @param {TrackedTable} trackedTable Table to bind events
 */
async function bindTrackedTableEvents(trackedTable) {
	await Excel.run(async (context) => {
		handleBindTrackedTableEvents(context, trackedTable);
		await context.sync()
	});
}

/**
 * Implements Tracked Table bindings 
 * @param {*} context 
 * @param {TrackedTable} trackedTable 
 */
async function handleBindTrackedTableEvents(context, trackedTable) {
	const worksheet = context.workbook.worksheets.getItemOrNullObject(trackedTable.worksheet);	
	const table = worksheet.tables.getItemOrNullObject(trackedTable.name);
	
	if (table) {
		// Bind Table Change Event
		table.onChanged.add((eventArgs) => {
			onTrackedTableChange(worksheet, table, eventArgs);
		});
		
		
		// seems to be causing the onChange event to not fire.
		table.onSelectionChanged.add((eventArgs) => {						
    		updateInfoPanel(trackedTable.name, eventArgs); // Update on initial load
		});
		
	}
}



/**
 * Show an info panel describing the selected Data Table.
 * @param {string} tableName The name of the Tracked Table to display
 * @param {*} eventArgs Excel Event information
 */
function updateInfoPanel(tableName, eventArgs) {

	const infoPanel = document.getElementById('infoPanel');

	if (eventArgs && eventArgs.isInsideTable && UIState.infoPanel.trackedTableName != tableName)  {
		const selectedTable = TrackedTables.tables.find(table => table.name === tableName);
    	
		// Clear the current info
		infoPanel.innerHTML = '';

		// Add new info
		if (selectedTable) {
			// Add table name as a header
			let header = document.createElement('h3');
			header.textContent = selectedTable.name;
			infoPanel.appendChild(header);

			// Add columns as list items
			let list = document.createElement('ul');
			selectedTable.columns.forEach(column => {
				let listItem = document.createElement('li');
				listItem.textContent = column;
				// Check if column is in trackedColumns and add input if necessary
				let trackedColumn = selectedTable.trackedColumns.find(tc => tc.name === column);
				if (trackedColumn) {
					// For simplicity, using text input; this can be extended for different data types
					let input = document.createElement('input');
					input.type = 'text';
					input.value = trackedColumn.source || '';
					listItem.appendChild(input);
				}
				list.appendChild(listItem);
			});
			infoPanel.appendChild(list);
		}
		UIState.infoPanel.trackedTableName = tableName
	} else {
		infoPanel.innerHTML = '';
		UIState.infoPanel.trackedTableName = '';
	} 
}


/**
 * Event to watch user updates to Tracked Tables.
 * @param {Excel.TableChangedEventArgs} eventArg 
 */
async function onTrackedTableChange(worksheet, table, eventArg) {

	let hasChange = false;

	const headerRange = table.getHeaderRowRange();
    headerRange.load(["address", "values"]); // Load the address property of the header range
	worksheet.load(["name"]);
	table.load(["name"]);
    await table.context.sync();

	// Find the member of Tracked tables that has the same table name. 
	const tableConfig = TrackedTables.tables.find(tt => {
		return tt.name === table.name;
	});

	// Does this change intersect the header?
	const intersectsHeader = doRangesIntersect(headerRange.address, `${worksheet.name}!${eventArg.address}`);

	console.log(eventArg);

	// Actions based on the change type.
	switch (eventArg.changeType) {
		case "RangeEdited":
			if (intersectsHeader) {
				hasChange = onHeaderChange(tableConfig, headerRange) || hasChange;
				
			}
			break;

		case "RowInserted": 
			break;

		case "RowDeleted":
			break;

		case "ColumnInserted":
			onHeaderChange(tableConfig, headerRange);
			hasChange = true;
			break;

		case "ColumnDeleted":
			onRemoveColumns(tableConfig, headerRange)
			hasChange = true;
			break;

		case "CellInserted":
			if (intersectsHeader) {
				hasChange = onHeaderChange(tableConfig, headerRange)  || hasChange;
			}
			break;
			
		case "CellDeleted": 
			// Deleteing a header rows just renames the column to "Column x" 
			// This should not be called when the user deletes a value in the table header.  
			if (intersectsHeader) {
				hasChange = onHeaderChange(tableConfig, headerRange) || hasChange;
			}
			break;
	}

	await table.context.sync();

	if (hasChange) {
		saveState(States.TrackedTables);
		applyStyleToTable(tableConfig);
	}


}

/**
 * On header rename, update the Tracked Table Config with the current Table state.  
 * @param {TrackedTable} tableConfig Tracked Table defintion of the table being changed.
 * @param {Excel.range} headerRange The range of cells for the table header.
 */
function onHeaderChange(tableConfig, headerRange) {

	const beforeColumnNames = tableConfig.trackedColumns.map(tc => tc.name);
	const afterColumnNames = headerRange.values[0];
	const changes = findColumnChanges(beforeColumnNames, afterColumnNames);
	
	renameTrackedColumns(tableConfig, changes.renamedColumns);	
	tableConfig.columns = afterColumnNames; 
	return changes.hasChanges
} 


/**
 * On column remove, update the Tracked Table Config with the current Table state.  
 * @param {TrackedTable} tableConfig Tracked Table defintion of the table being changed.
 * @param {Excel.range} headerRange The range of cells for the table header.
 */
function onRemoveColumns(tableConfig, headerRange){
	const beforeColumnNames = tableConfig.trackedColumns.map(tc => tc.name);
	const afterColumnNames = headerRange.values[0];
	const deleted = findColumnRemoved(beforeColumnNames, afterColumnNames);

	removeTrackedColumns(tableConfig, deleted);
	tableConfig.columns = afterColumnNames; 
}


/**
 * Rename Tracked Columns to match the users label.
 * @param {TrackedTable} tableConfig The Tracked Table being changed
 * @param {BeforeAfter[]} renamedArray From/to change pairs {before: "aaa", after: "bbb"} 
 */
function renameTrackedColumns(tableConfig, renamedArray) {

	// Loop through the renamed columns
	renamedArray.forEach((r) =>{
		// Find the renamed columns from the before name.
		let trackedCol = tableConfig.trackedColumns.find((c) => {
			return c.name == r.before;
		});
		
		if(trackedCol) {		
			// Rename the Tracked Column
			logEvent(LogEvents.Table.RenamedColumn, trackedCol, trackedCol.name);			
			trackedCol.name = r.after; 
		}
	});
}

/**
 * Remove a column from the tracked array.
 * @param {TrackedTable} tableConfig The Tracked Table being changed
 * @param {string[]} deletedArray Array of column name strings to be removed from the table.
 */
function removeTrackedColumns(tableConfig, deletedArray) {
	
	// Loop through the deleted columns.
	deletedArray.forEach((d) =>{
		let index = tableConfig.trackedColumns.findIndex(obj => obj.name == d);

		if (index !== -1) {
			
			logEvent(LogEvents.Table.DeleteColumn, tableConfig, tableConfig.trackedColumns[index]);	
			// remove the Tracked Column
			tableConfig.trackedColumns.splice(index, 1);
		}
	});
}


/**
 * Find the list of columns renamed or reordered in the table.
 * @param {[string]} before Array of column names before the change.
 * @param {[string]} after Array of column names after the change.
 * @returns 
 */
function findColumnChanges(before, after) {
    const renamedColumns = [];
    const reorderedColumns = [];

    // Check for renamed columns
    const beforeSet = new Set(before);
    const afterSet = new Set(after);

    // Identify removed and added columns
    const removed = before.filter(col => !afterSet.has(col));
    const added = after.filter(col => !beforeSet.has(col));

    // Assuming each removed column corresponds to an added column
    if (removed.length === added.length) {
        for (let i = 0; i < removed.length; i++) {
            renamedColumns.push({ before: removed[i], after: added[i] });
        }
    }

    // Check for reordered columns (if no renamed columns are found)
    if (renamedColumns.length === 0) {
        before.forEach((col, index) => {
            if (col !== after[index] && afterSet.has(col)) {
                reorderedColumns.push(col);
            }
        });
    }

    return {
        renamedColumns,
        reorderedColumns,
        hasChanges: renamedColumns.length > 0 || reorderedColumns.length > 0
    };
}





/* 
###########################################################################################
	Sync Table settings
########################################################################################### 
*/


/**
 * Find the list of columns removed from the table.
 * @param {[string]} before Array of column names before the change.
 * @param {[string]} after Array of column names after the change.
 * @returns 
 */
function findColumnRemoved (before, after) {
	const deletedColumns = [];

	// Detect inserted and deleted columns
    const afterSet = new Set(after);
	
	// Identify columns absent from after state.
	deletedColumns.push(...before.filter(col => !afterSet.has(col)));
	return deletedColumns;
}


/**
 * Loads all necessary properties of worksheets and their tables.
 * @param {Excel.RequestContext} context - The Excel request context.
 * @returns {Excel.WorksheetCollection} The loaded collection of worksheets.
 */
async function loadWorksheetAndTableProperties(context) {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name"); // Load the name property of worksheets
    await context.sync();

    // Now that the worksheets are loaded, load the properties of their tables
    worksheets.items.forEach(sheet => {        
		sheet.tables.load("items/name"); // Load name and address properties of tables
    });

    await context.sync(); // Perform another sync to load the tables
    return worksheets;
}


/**
 * Updates tracked table information based on the loaded data.
 * @param {Excel.WorksheetCollection} worksheets - The collection of worksheets.
 * @returns {string[]} Array of found table IDs.
 */
async function onSyncUpdateTrackedTables(context, worksheets) {
    const foundTables = [];
    for (const sheet of worksheets.items) {
        for (const wsTable of sheet.tables.items) {
            
			let table = TrackedTables.tables.find((t) => t.id === wsTable.id);
			if (table) {
				// Update the Tracked Table settings.
				if (table.name !== wsTable.name) {
					logEvent(LogEvents.Table.RenamedTable, table, table.name);
					table.name = wsTable.name;
				}

				if (table.worksheet !== sheet.name) {
					logEvent(LogEvents.Table.MovedRange, table, table.worksheet);
					table.worksheet = sheet.name;
				}

				const address = await getTableRange(context, wsTable);								

				if (table.range !== address.range) {
					logEvent(LogEvents.Table.MovedRange, table, table.range);
					table.range = address.range;
				}
				
			}

            foundTables.push(wsTable.id);
        }
    }
    return foundTables;
}

/**
 * Get the range an Excel Table occupies 
 * @param {Excel.Table} table 
 * @returns 
 */
async function getTableRange (context, table) {
	const tableRange = table.getRange(); 
	tableRange.load("address")
	await context.sync();
	return splitExcelAddress(tableRange.address);
}

/**
 * Removes tables from the tracked list if they are no longer present.
 * @param {string[]} foundTables - Array of found table IDs.
 */
function removeAbsentTables(foundTables) {
    TrackedTables.tables = TrackedTables.tables.filter(table => {
        let exists = foundTables.some(ft => ft === table.id);
        if (!exists) {
            logEvent(LogEvents.Table.DeleteTable, TrackedTables, table);
            return false;
        }
        return true;
    });
}

/**
 * Main function to synchronize tracked table information.
 */
async function syncTrackedTableInfo() {
    await Excel.run(async (context) => {
        const worksheets = await loadWorksheetAndTableProperties(context);
        const foundTables = await onSyncUpdateTrackedTables(context, worksheets);
        removeAbsentTables(foundTables);
        saveState(States.TrackedTables);
    });
}



/* 
###########################################################################################
	Log Table Change
########################################################################################### 
*/

/**
 * Enum for Logging Events
 */
const LogEvents = {
    Table: {
		RenamedColumn: 'column_rename',
		DeleteColumn: 'column_delete',
		RenamedTable: 'table_rename',
		MovedRange: 'table_address',		
		DeleteTable: 'table_delete'
	}, 
	StylingChangeEvents: {  
		SetPrimaryColor: 'style_setPrimary',
		SetSecondaryColor: 'style_setSecondary',
	}
};


/**
 * Log a change event. HistoryItem is stored in for logging purposes and possibly rollback in the future. 
 * @param {LogEvents} event Event type to log
 * @param {*} trackedItem Object against which the history is stored.
 * @param {*} historyItem Item stored in history
 */
function logEvent(event, trackedItem, historyItem ) {
	
	if (Object.values(LogEvents.Table).includes(event)) {
		handleLogTableChangeEvent(event, trackedItem, historyItem);
	
	} else if (Object.values(LogEvents.StylingChangeEvents).includes(event)) {
		handleLogStylingChangeEvent(event, trackedItem, historyItem);
	
	} else {
		console.log("logEvent(): Unknown event type");
	}

}

/**
 * Log events related to Tracked Tables.
 * @param {LogEvents} event Event type to log
 * @param {*} trackedItem Object against which the history is stored.
 * @param {*} historyItem Item stored in history
 */
function handleLogTableChangeEvent(event, trackedItem, historyItem) {

	switch (event) {
        case LogEvents.Table.RenamedColumn:

			if(trackedItem) {		
				ensurePathExists(trackedItem, "history");
				if(!trackedItem.history.name) {
					trackedItem.history.name = [];
				}				
				trackedItem.history.name.push(historyItem);				
			}
			break;
	
		case LogEvents.Table.DeleteColumn:

			ensurePathExists(trackedItem, "history.columns")
			if(!trackedItem.history.columns.removed) {
				trackedItem.history.columns.removed = [];
			}
			
			trackedItem.history.trackedColumns.removed.push(historyItem);
			
            break;

		case LogEvents.Table.RenamedTable:

			ensurePathExists(trackedItem, "history")
			if(!trackedItem.history.name) {
				trackedItem.history.name = [];
			}
			
			trackedItem.history.name.push(historyItem);
			
            break;

		case LogEvents.Table.MovedRange:

			ensurePathExists(trackedItem, "history")
			if(!trackedItem.history.range) {
				trackedItem.history.range = [];
			}
			
			trackedItem.history.range.push(historyItem);
			
            break;

		case LogEvents.Table.DeleteTable:

			ensurePathExists(trackedItem, "history")
			if(!trackedItem.history.deleted) {
				trackedItem.history.deleted = [];
			}
			
			trackedItem.history.deleted.push(historyItem);
			
            break;
					
        
		default:
            console.log("handleLogTableChangeEvent(): Unknown event type");
    }

}


/**
 * Log events related to Styles.
 * @param {LogEvents} event Event type to log
 * @param {*} trackedItem Object against which the history is stored.
 * @param {*} historyItem Item stored in history
 */
function handleLogStylingChangeEvent(event, trackedItem, historyItem) {
	switch (event) {
        case LogEvents.StylingChangeEvents.SetPrimaryColor:
            // Handle Set Primary Color
            //handleSetPrimaryColor(trackedTable, historyItem);
            break;
        case LogEvents.StylingChangeEvents.SetSecondaryColor:
            // Handle Set Secondary Color
            //handleSetSecondaryColor(trackedTable, historyItem);
            break;
        default:
            console.log("handleLogTableChangeEvent(): Unknown event type");
    }
}


/**
 * Check that an object path exists, if not, create it.
 * @param {*} obj target Object
 * @param {string} path Object path to create/check
 */
function ensurePathExists(obj, path) {
    const parts = path.split('.');

    let currentPart = obj;
    for (let part of parts) {
        if (!currentPart[part]) {
            currentPart[part] = {}; // Create a new object if the part doesn't exist
        }
        currentPart = currentPart[part]; // Move to the next level
    }
}


/* 
###########################################################################################
	Custom XML Part
########################################################################################### 
*/

/**
 * A list of session states stored in the workbook.
 */
const States = {
    TrackedTables: "Remind_TrackedTables"
};

/**
 * Store a session state object to the workbook custom XML part.
 * @param {States} stateType Target state type 
 */
async function saveState(stateType) {
	await Excel.run(async (context) => {
		switch (stateType) {
			case States.TrackedTables:
				handleSaveState(context, stateType, TrackedTables)
			break;
		}
	});
}


/**
 * Load a state object from the workbook's XML custom parts.
 * @param {States} stateType State type to load from the XML Custom Part
 */
async function getState(stateType) {
	await Excel.run(async (context) => {
		
		if (stateType == States.TrackedTables || stateType == "all" ) {
			const tt = await handleGetState(context, States.TrackedTables);
			if(tt && tt.tables)  {
				TrackedTables.tables = tt.tables;
			}
		}
	});	
}


async function clearState(stateType) {
	await Excel.run(async (context) => {
		if(stateType == "all") {
			return promises = Object.values(States).map(stateType => handleClearState(context, stateType));
		} else {
			await handleClearState(context, stateType);
		}
	});	
}

/**
 * Load all state types from the workbook's XML custom parts
 * @returns 
 */
async function getAllStates() {
	await getState("all");		
}


/**
 * Handler for saveState.
 * @param {States} stateType Target state type
 * @param {*} stateObject Object to be stored
 */
async function handleSaveState(context, stateType, stateObject) {
	
	const existingId = await getStateId(context, stateType);
	const xmlContent = createStateXml(stateObject)
	let customXmlPart;
	
	if(!existingId.value) {
		customXmlPart = context.workbook.customXmlParts.add(xmlContent);
		customXmlPart.load("id");

	} else {
		customXmlPart = context.workbook.customXmlParts.getItem(existingId.value);
		customXmlPart.setXml(xmlContent);
		customXmlPart.load("id");
	}
	await context.sync();

	// Store the XML part's ID in a setting.
	const settings = context.workbook.settings;
	settings.add(stateType, customXmlPart.id);
	await context.sync();
}



/**
 * Load a state object from the workbooks XML custom parts.
 * @param {States} stateType State type to load from the XML Custom Part
 */
async function handleGetState(context, stateType) {
	
	const settings = context.workbook.settings;
	const xmlPartIDSetting = settings.getItemOrNullObject(stateType).load("value");
	await context.sync();

	if (xmlPartIDSetting.value) {
		const customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
		const xml = customXmlPart.getXml()
		await context.sync();

		const obj =  parseStateXml(xml.value);
		return obj
		
	}

}


async function handleClearState(context, stateType) {
	  const settings = context.workbook.settings;
	  const xmlPartIDSetting = settings.getItemOrNullObject(stateType).load("value");
	  await context.sync();
  
	  if (xmlPartIDSetting.value) {
		let customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
		const xmlBlob = customXmlPart.getXml();
		customXmlPart.delete();
		customXmlPart = context.workbook.customXmlParts.getItemOrNullObject(xmlPartIDSetting.value);
  
		await context.sync();
  
		if (customXmlPart.isNullObject) {
		  // Delete the unneeded setting too.
		  xmlPartIDSetting.delete();
		} 
		await context.sync();
	  }
  }


/**
 * Get the Id of the custom XML part from the settings.
 * @param {Excel.context} context 
 * @param {States} stateType Target stateful object
 * @returns 
 */
async function getStateId(context, stateType) {
	const settings = context.workbook.settings;
	const xmlPartIDSetting = settings.getItemOrNullObject(stateType).load("value");
	await context.sync();

	return xmlPartIDSetting;
}


async function listStates() {
	await Excel.run(async (context) => {
		const customXmlParts = context.workbook.customXmlParts;
		customXmlParts.getCount();
		customXmlParts.load("items")
		await context.sync();
	});
}
  
/**
 * Convert the JSON object to a string
 * @param {*} stateObject 
 * @returns 
 */
function createStateXml(stateObject) {
	const jsonString = JSON.stringify(stateObject);
	return `<RemindState xmlns='http://schemas.remindbi.com/review/1.0'>${jsonString}</RemindState>`;
}

/**
 * Convert XML from state to an object. Only works for XML created by createStateXml()
 * @param {string} xml XML created by createStateXml()
 * @returns 
 */
function parseStateXml(xml) {
	// May need to be changed to a DOM parser. But as createStateXml() is so simple this is more efficient.
	const start = "<RemindState xmlns='http://schemas.remindbi.com/review/1.0'>".length;
	const end = xml.length - "</RemindState>".length;
	const jsonString = xml.substring(start, end);
	return JSON.parse(jsonString);	

}


/* 
###########################################################################################
	Products testing
########################################################################################### 
*/





/**
 * Start an Azure function for Dimension Query
 * @returns promise
 */
async function getTrackedTableData(trackedTable) {  
    /*
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
	*/

	if (trackedTable && trackedTable.query) {
		// Make an Ajax call to the Azure functions for the table query data
		const response = await fetch(trackedTable.query.url, trackedTable.query.ajax);

		if (!response.ok) {
			throw new Error(`HTTP error! status: ${response.status}`);
		}

		const data = await response.json();

		// If the response is valid add the data to the table rows.
		if (data && data[0] && data[0].result) {
			const j = JSON.parse(data[0].result)
			trackedTable.rows.splice(0, TrackedTables.tables[0].rows.length); // Remove all elements from the array
			trackedTable.rows.push(...j); // Merge arrays
		}
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


/* 
###########################################################################################
	JSON - SQL Mapping
########################################################################################### 
*/

async function  displayMappingObject (jsonObject) {
	
	Excel.run(async (context) => {
		const sheet = context.workbook.worksheets.getActiveWorksheet();

		// Generate the data to be written to Excel
		const data = traverseObject(jsonObject);

		// Write the data to Excel
		const range = sheet.getRange(`A1:F${data.length}`);
		range.values = data;
		range.format.autofitColumns();

		await context.sync();
	});			
}

function removeQuotes(str) {
    if (typeof str === 'string' && str.startsWith('"') && str.endsWith('"')) {
        return str.substring(1, str.length - 1);
    }
    return str;
}

// Function to recursively traverse the JSON object
function traverseObject(obj, path = '') {
    let result = [];
    for (const key in obj) {
        let newPath = path ? `${path}.${key}` : key;
        let value = obj[key];
        let isLeaf = !(value instanceof Object);
        let isArray = Array.isArray(value);

        if (isArray) {
            if (value.length > 0 && (typeof value[0] !== 'object' || value[0] === null)) {
                // Array of primitives
                let row = [
                    'L',                // Leaf column
                    newPath + '[]',     // Path with array indicator
                    key,                // Attribute name
                    value[0] ? value[0] : '',   // Value (blank for array of primitives)                    
                    value[0] ? guessSqlType(value[0]) : '', // SQL type 
					''
                ];
                result.push(row);
            } else {
                // Array of objects
                result = result.concat(traverseObject(value[0], `${newPath}[]`));
            }
        } else if (!isLeaf) {
            // Non-leaf object
            result = result.concat(traverseObject(value, newPath));
        } else {
            // Leaf node
            let formattedValue = typeof value === 'string' ? value : JSON.stringify(value);
            let row = [
                'L',                     // Leaf column
                newPath,                // Path
                key,                    // Attribute name
                removeQuotes(formattedValue), // Value with quotes removed
                guessSqlType(value),     // SQL type
				''						// Blank column	
            ];
            result.push(row);
        }
    }
    return result;
}


// Function to guess SQL type
function guessSqlType(value) {
	const valueType = typeof value;
	switch (valueType) {
		case 'number': return 'INT';
		case 'boolean': return 'BOOLEAN';
		case 'string': default: return 'VARCHAR';
	}
}
