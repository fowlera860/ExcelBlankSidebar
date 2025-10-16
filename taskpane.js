Office.onReady(() => {
    pollSidebarMessage();
});

async function pollSidebarMessage() {
    try {
        await Excel.run(async (context) => {
            console.log("Entered Excel.run")
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            const range = sheet.names.getItem("SidebarMessage").getRange();
            range.load("address");
            range.load("values");
            await context.sync();
            
            console.log(sheet.name)
            console.log(range.address)

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