Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("fill-table").onclick = fillCableData;
    }
});

async function fillCableData() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Заголовки таблицы
            const headers = [["№ п/п", "Материал", "Ед. изм.", "Кол-во", "СТ, тыс. руб."]];
            
            // Данные из вашего скриншота
            const data = [
                ["1", "Кабель ППГнг(А) HF 1х120", "км", 0.08, 99]
            ];

            // 1. Очищаем диапазон и вставляем заголовки
            const headerRange = sheet.getRange("A1:E1");
            headerRange.values = headers;
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = "#D3D3D3"; // Серый цвет заголовка

            // 2. Вставляем данные под заголовками
            const dataRange = sheet.getRange("A2:E2");
            dataRange.values = data;

            // 3. Автоподбор ширины колонок
            headerRange.getUsedRange().format.autofitColumns();

            await context.sync();
            console.log("Данные успешно добавлены!");
        });
    } catch (error) {
        console.error(error);
    }
}