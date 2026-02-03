Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("fill-table").onclick = fillCableData;
        document.getElementById("new-table").onclick = createVORSheet;       
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

async function createVORSheet() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // 1. Аналог ws.Cells.Clear (полная очистка)
            const allCells = sheet.getUsedRange(true); 
            allCells.clear();

            // 2. Настройка ширины столбцов (коэффициент пересчета ~7.1)
            const columnWidths = [15.29, 41.14, 10.57, 18.29, 32.14, 41.14, 22, 28.14, 41.14];
            columnWidths.forEach((width, index) => {
                sheet.getRangeByIndexes(0, index, 100, 1).format.columnWidth = width * 7; 
            });

            // 3. Заполнение метаданных (A1:D13)
            sheet.getRange("A1:A13").values = [
                ["Документ"], ["Версия"], [""], ["Наименование стройки"], 
                ["Наименование объекта"], ["ВОР №"], ["Основание"], ["Дата составления"],
                [""], ["Составил ФИО"], ["Должность"], ["Проверил ФИО"], ["Должность"]
            ];
            sheet.getRange("A1:A13").format.font.color = "gray";

            sheet.getRange("D1:D8").values = [
                ["Ведомость объемов работ"], ["3_01"], [""], 
                ["Капитальный ремонт конструкций"], ["Объект"], ["ВОР-01-01-01"], 
                ["Техническая документация"], [new Date().toLocaleDateString()]
            ];

            // 4. Шапка таблицы (A15:I16)
            const headerRange = sheet.getRange("A15:I15");
            headerRange.values = [[
                "№ п.п.", 
                "Наименование работ, ресурсов, затрат по проекту", 
                "Ед. изм.", 
                "Объем работ / Количество", 
                "Формула расчета объемов работ и расхода материалов, потребности ресурсов", 
                "Ссылка на чертежи, спецификации в проектной документации", 
                "Наименование файла", 
                "Номер страниц (через пробел)", 
                "Дополнительная информация (комментарий)"
            ]];
            headerRange.format.fill.color = "#E5E4E2";
            headerRange.format.font.bold = true;
            headerRange.format.wrapText = true;
            headerRange.format.horizontalAlignment = "Center";

            sheet.getRange("A16:I16").values = [["1", "2", "3", "4", "5", "6", "6.1", "6.2", "7"]];
            sheet.getRange("A16:I16").format.horizontalAlignment = "Center";

            // 5. Объединение ячеек (Раздел 1. XXX)
            const sectionRange = sheet.getRange("A17:I17");
            sectionRange.merge();
            sectionRange.values = [["Раздел: 1. XXX"]];
            sectionRange.format.font.bold = true;
            sectionRange.format.fill.color = "#E5E4E2";

            // 6. Границы для всей таблицы
            const tableRange = sheet.getRange("A15:I18");
            const borders = tableRange.format.borders;
            borders.getItem('EdgeTop').style = 'Continuous';
            borders.getItem('EdgeBottom').style = 'Continuous';
            borders.getItem('EdgeLeft').style = 'Continuous';
            borders.getItem('EdgeRight').style = 'Continuous';
            borders.getItem('InsideVertical').style = 'Continuous';
            borders.getItem('InsideHorizontal').style = 'Continuous';

            await context.sync();
        });
    } catch (error) {
        console.error("Ошибка при создании шаблона: " + error);
    }
}