Office.onReady(() => {
    pollSidebarMessage();
});

async function pollSidebarMessage() {
    try {
        await Excel.run(async (context) => {
            console.log("Entered Excel.run")
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            console.log(sheet.name)
            const range = sheet.names.getItem("SidebarMessage").getRange();
            console.log(range.address)
            range.load("values");
            await context.sync();

            const msg = range.values[0][0];
            if (msg) {
                document.getElementById("content").innerText = msg;
            }
        });
    } catch (error) {
         //console.log(error)
        console.log("error thrown")
        
    } finally {
        setTimeout(pollSidebarMessage, 1000); // poll every 1 second
    }
}