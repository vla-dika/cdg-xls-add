// Ждем, пока Office.js загрузится
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("run-button").onclick = run;
    }
});

async function run() {
    try {
        await Excel.run(async (context) => {
            // Получаем активный лист
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            // Выбираем ячейку A1
            const range = sheet.getRange("A1");
            // Записываем текст и делаем его жирным
            range.values = [["Привет из GitHub Pages, Андр12331223123123123ей!"]];
            range.format.font.bold = true;
            
            // Синхронизируем изменения с Excel
            await context.sync();
        });
    } catch (error) {
        console.error("Ошибка: " + error);
    }
}