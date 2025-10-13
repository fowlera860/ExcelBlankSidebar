Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        pollForUpdates();
    }
});


async function pollForUpdates() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("SidebarMessage");
            range.load("values");
            await context.sync();

            const msg = range.values[0][0];
            if (msg) {
                // Update the task pane content
                document.getElementById("content").innerText = msg;
            }
        });
    } catch (error) {
        console.error(error);
    } finally {
        // poll every 1 second
        setTimeout(pollForUpdates, 1000);
    }
}
