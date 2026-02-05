Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("new-table").onclick = createNewTable;
        document.getElementById("formating").onclick = formatVORTable;
    }
});

async function createNewTable() {
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


async function formatVORTable() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.getActiveWorksheet();
            
            // 1. Проверка структуры (как в блоке wsVOR)
            const checkRange = sheet.getRange("A1:B5");
            checkRange.load(["values"]);
            await context.sync();

            if (checkRange.values[0][0] !== "Документ") {
                if (checkRange.values[4][0] === "№" || checkRange.values[4][1].startsWith("Наименование")) {
                    console.error("Макрос не настроен для формата МГЭ");
                    return;
                } else {
                    console.error("Таблица ГГЭ не обнаружена");
                    return;
                }
            }

            // 2. Оптимизация (ScreenUpdating в JS не нужен, но расчеты можно отключить)
            context.workbook.application.calculationMode = Excel.CalculationMode.manual;

            // 3. Замена текста "Раздел " на "Раздел: "
            const lastRowRange = sheet.getUsedRange().getLastRow();
            lastRowRange.load("rowIndex");
            await context.sync();

            const rangeA = sheet.getRange(`A17:A${lastRowRange.rowIndex + 1}`);
            rangeA.load("values");
            await context.sync();

            let valuesA = rangeA.values;
            for (let i = 0; i < valuesA.length; i++) {
                if (valuesA[i][0] && valuesA[i][0].toString().startsWith("Раздел ")) {
                    valuesA[i][0] = valuesA[i][0].replace("Раздел ", "Раздел: ");
                }
            }
            rangeA.values = valuesA;

            // 4. Групповое форматирование шапки (ячейки D1-D13)
            const headerData = sheet.getRange("D1:D13");
            headerData.load("values");
            await context.sync();

            // Подсвечиваем пустые ячейки розовым
            for (let i = 0; i < 13; i++) {
                // Пропускаем строки 3 и 9 (в VBA они пропущены: D1,2,4,5,6,7,8,10,11,12,13)
                if ([2, 8].includes(i)) continue; 
                
                if (!headerData.values[i][0]) {
                    sheet.getRange(`D${i + 1}`).format.fill.color = "#FF80FF";
                }
            }

            // 5. Форматирование шрифтов и выравнивания
            const topArea = sheet.getRange("A1:I14");
            topArea.format.font.size = 11;
            topArea.getRange("A1:A14").format.horizontalAlignment = "Left";
            topArea.getRange("A1:A13").format.font.color = "#808080";

            // 6. Вызов вспомогательных функций (форматирование таблицы)
            // В JS их нужно реализовать отдельно или вложить сюда
            await formatTableDesign(sheet);

            // 7. Финализация (аналог Ctrl+Home и Zoom)
            context.workbook.application.calculationMode = Excel.CalculationMode.automatic;
            sheet.getRange("A1").select();
            
            // Масштаб в Office JS API доступен через настройки представления (не во всех версиях)
            // sheet.activeView.zoom = 100; 

            await context.sync();
            console.log("Форматирование завершено");
        });
    } catch (error) {
        console.error(error);
    }
}
