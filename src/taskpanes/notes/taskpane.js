Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
        // Assign event handlers and interact with Excel
        if (info.platform === Office.PlatformType.PC) {
            // Office on Windows
            // Configure your add-in for Windows here
        } else if (info.platform === Office.PlatformType.OfficeOnline) {
            // Office Online
            // Configure your add-in for Office Online here
        }
    }

    document.getElementById("run").onclick = run;
    
});


async function run() {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();
  
        // Read the range address
        range.load("address");
  
        // Update the fill color
        range.format.fill.color = "yellow";
  
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
  