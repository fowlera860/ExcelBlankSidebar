Office.onReady(() => {
    pollSidebarMessage();
});

async function pollSidebarMessage() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.names.getItem("SidebarMessage").getRange();
            range.load("values");
            await context.sync();

            const msg = range.values[0][0];
            if (msg) {
                document.getElementById("content").innerText = msg;
            }
        });
    } catch (error) {
        console.error(error);
    } finally {
        setTimeout(pollSidebarMessage, 1000); // poll every 1 second
    }
}