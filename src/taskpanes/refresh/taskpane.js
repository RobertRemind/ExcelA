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
            azureFunctions.forEach(functionDetails => {
                startAzureFunction(functionDetails.id);
            });
        });

        // Bind make table
        document.getElementById('btnCreateDimensionTable').addEventListener('click', function() {
            setup();
        });


    }
});





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
    const finalStatus = document.createElement('div');
    finalStatus.textContent = '<p>' + message + '</p>';
    document.body.appendChild(finalStatus);
}


/* ######################################################################################## */ 

/** Set up Sample worksheet. */
async function setup() {
    debugger;
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItemOrNullObject("Products").delete();
      const sheet = context.workbook.worksheets.add("Products");
  
      const productsTable = sheet.tables.add("A1:C1", true /*hasHeaders*/);
      productsTable.name = "ProductsTable";
  
      productsTable.getHeaderRowRange().values = [["Product", "ProductID", "ProductName"]];
  
      productsTable.rows.add(
        null /*add at the end*/,
        products.map((p) => [null, p.productID, p.productName])
      );
  
      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
  
      sheet.activate();
  
      await context.sync();
    });
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
  