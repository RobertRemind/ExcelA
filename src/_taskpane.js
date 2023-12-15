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
        document.getElementById('startFunctionsBtn').addEventListener('click', function() {
            completedFunctionsCount = 0;
            errorOccurred = false;
            azureFunctions.forEach(functionDetails => {
                startAzureFunction(functionDetails.id);
            });
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
