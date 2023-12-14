const azureFunctions = [
    { id: 1, name: 'Function 1' },
    { id: 2, name: 'Function 2' },
    // ... other functions
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
    // Replace with actual Azure function call logic
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            const simulatedError = false; // Set to true to simulate an error
            if (simulatedError) {
                reject(new Error('Simulated error'));
            } else {
                resolve();
            }
        }, 3000 * functionId);
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
