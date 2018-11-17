const prsr = require("osmosis");
const xl = require('excel4node');

const MAX_PAGE_NUM = 377;   // Количество страниц, которое нужно спарсить, начиная с первой
const START = 1;           // Разрешенные значения: 0 <= START < COLS_NUMBER; Начальный столбец в таблице считая со второго столбца слева (первый с номерами)
const INTERVAL = 4;       // Разрешенные значения: 0 <= INTERVAL < COLS_NUMBER; Количество столбцов в таблице, которые нужно спарсить начиная со START

function setCols(START, INTERVAL, data, COLS_NUMBER = 11) {
    if (typeof START !== "number"       ||
        typeof INTERVAL !== "number"    ||
        typeof COLS_NUMBER !== "number" ||
        START < 0                       ||
        INTERVAL < 0                    ||
        COLS_NUMBER < 0                 ||
        START > COLS_NUMBER             ||
        INTERVAL > COLS_NUMBER          ||
        START + INTERVAL > COLS_NUMBER) return false;

    for (var i = 0; i < INTERVAL; i++) {
        for (var j = 0, counter = START; j < data.length; j += COLS_NUMBER, counter += COLS_NUMBER) {
            ws.cell(data[j], i + 1)
                .string(data[i + counter + 1])
                .style(style);
        }
    }

    return true;
}

var wb = new xl.Workbook();
var ws = wb.addWorksheet('data');

var style = wb.createStyle({
    font: {
        color: '#000000',
        size: 12,
    }
});

prsr.get('https://finviz.com/screener.ashx?v=111&r=1')
    .paginate('.body-table .tab-link:last-child:not(:nth-child(1))[href]', MAX_PAGE_NUM)
    .delay(100)
    .set(['.screener-body-table-nw a'])
    .data((data) => {
        let success = setCols(START, INTERVAL, data);
        success ? console.log("Loading...", "Line:", data[0]) : console.log("Error input data");
    })
    .done(() => { 
        wb.write('data.xlsx');
        console.log("Data saved");
    });