Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("new-table").onclick = createNewTable;
        document.getElementById("renumb-positions").onclick = renumberPosition;
        
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

async function renumberPosition() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.getActiveWorksheet();
            
            // Определяем используемый диапазон для поиска последней строки
            const usedRange = sheet.getUsedRange();
            const lastRow = usedRange.getLastRow();
            lastRow.load("rowIndex");
            
            // Проверка формата таблицы (ячейки A1 и A5/B5)
            const checkRange = sheet.getRange("A1:B5");
            checkRange.load("values");
            
            await context.sync();

            let startRow;
            const isGGE = checkRange.values[0][0] === "Документ";
            const isMGE = checkRange.values[4][0] === "№" || (checkRange.values[4][1] && checkRange.values[4][1].toString().startsWith("Наименование"));

            if (isGGE) {
                startRow = 18; // В VBA i = 18
            } else if (isMGE) {
                startRow = 7;  // В VBA i = 7
            } else {
                // Аналог MsgBox в JS надстройках лучше делать через UI, здесь используем консоль или throw
                console.error("Формат таблицы не распознан");
                return;
            }

            // Получаем диапазон столбца А с учетом найденного формата
            const totalRows = lastRow.rowIndex + 1;
            const rangeA = sheet.getRange(`A${startRow}:A${totalRows}`);
            
            // Загружаем значения и информацию об объединении
            rangeA.load(["values", "address"]);
            
            // В Office JS проверка каждой ячейки на объединение в цикле медленная. 
            // Получаем все объединенные области в этом диапазоне одним запросом.
            const mergedAreas = rangeA.getMergedAreasOrNullObject();
            mergedAreas.load("address");

            await context.sync();

            let values = rangeA.values;
            let counter = 0;

            for (let i = 0; i < values.length; i++) {
                let currentRowIdx = startRow + i;
                let cell = sheet.getRange(`A${currentRowIdx}`);
                
                // Проверяем, является ли ячейка началом объединенной области
                // В JS API работа с mergedAreas сложнее, упростим логику:
                cell.load(["isMerged", "address"]);
                await context.sync(); // Примечание: для скорости лучше загружать всё сразу, но здесь для наглядности

                if (cell.isMerged) {
                    let mergeArea = cell.getMergeArea();
                    mergeArea.load("address");
                    await context.sync();

                    // Если адрес ячейки не совпадает с началом объединенной области - пропускаем
                    let firstCellAddress = mergeArea.address.split(":")[0];
                    if (cell.address !== firstCellAddress) {
                        continue; 
                    }
                }

                let cellValue = values[i][0] ? values[i][0].toString() : "";

                // Условие: ячейка пустая или начинается с цифры
                if (cellValue === "" || /^\d/.test(cellValue)) {
                    counter++;
                    values[i][0] = counter;
                }
            }

            // Записываем обновленные значения обратно одним махом
            rangeA.values = values;
            
            await context.sync();
            console.log("Нумерация позиций в столбце [A] выполнена.");
        });
    } catch (error) {
        console.error("Ошибка при перенумерации: " + error);
    }
}