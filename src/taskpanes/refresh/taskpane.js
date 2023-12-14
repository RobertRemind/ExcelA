const azureFunctions = [
    { 
        id: 1, 
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

document.getElementById('startFunctionsBtn').addEventListener('click', function() {
    completedFunctionsCount = 0;
    errorOccurred = false;
    azureFunctions.forEach(functionDetails => {
        startAzureFunction(functionDetails.id);
    });
});



function startAzureFunction(functionId) {
    if (!document.getElementById('statusIndicator' + functionId)) {
        createStatusIndicator(functionId);
    }
    updateStatus(functionId, 'Starting...', 'running');
    callAzureFunction(functionId)
        .then(() => {
            debugger
            updateStatus(functionId, 'Completed', 'completed');
        })
        .catch((error) => {
            debugger
            updateStatus(functionId, 'Error: ' + error.message, 'error');
            errorOccurred = true;
        })
        .finally(() => {
            debugger
            checkAllFunctionsCompleted();
        });
}



function createStatusIndicator(functionId) {
    const statusIndicators = document.getElementById('statusIndicators');
    const indicator = document.createElement('div');
    indicator.id = 'statusIndicator' + functionId;
    indicator.textContent = `Function ${functionId} Status: Idle`;
    statusIndicators.appendChild(indicator);
}

function updateStatus(functionId, message, status) {
    const statusIndicator = document.getElementById('statusIndicator' + functionId);
    statusIndicator.innerHTML = `Function ${functionId} Status: <span class="${status}">${message}</span>`;
}

async function callAzureFunction(functionId) {
    debugger
    return await fetch(azureFunctions[functionId].url, {        
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: azureFunctions[functionId].data
    });
}

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

function showFinalStatus(message) {
    const finalStatus = document.createElement('div');
    finalStatus.textContent = message;
    document.body.appendChild(finalStatus);
}
