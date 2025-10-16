Office.onReady(() => {
    pollSidebarMessage();
});

async function pollSidebarMessage() {
    try {
        await Excel.run(async (context) => {
            //console.log("Entered Excel.run");
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            // queue both worksheet and workbook name lookups
            const sheetNameObj = sheet.names.getItemOrNullObject("SidebarMessage");
            const wbNameObj = context.workbook.names.getItemOrNullObject("SidebarMessage");

            await context.sync();
            //console.log(sheet.name);

            let namedRangeSource = null;
            if (!sheetNameObj.isNullObject) {
                namedRangeSource = sheetNameObj;
                //console.log("Found SidebarMessage as worksheet-scoped name");
            } else if (!wbNameObj.isNullObject) {
                namedRangeSource = wbNameObj;
                //console.log("Found SidebarMessage as workbook-scoped name");
            } else {
                //console.log("SidebarMessage named range not found (worksheet- or workbook-scoped).");
                return;
            }

            const range = namedRangeSource.getRange();
            range.load(["address", "values"]);
            await context.sync();
            //console.log(range.address);

            const msg = range.values && range.values[0] && range.values[0][0];
            if (msg) {
                document.getElementById("content").innerText = msg;
            }
        });
    } catch (error) {
        // log the real error so you can inspect it
        console.error("Error in pollSidebarMessage:", error);
    } finally {
        setTimeout(pollSidebarMessage, 1000); // poll every 1 second
    }
}
