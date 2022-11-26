console.log("Spreadsheet data automation - Converter")

const ExceltoJSON = require('exceljs');
const fs = require('fs');


var format = {
    "name": "",
    "freeze": "",
    "styles": [],
    "merges": [],
    "rows": {},
    "cols": {},
    "validations": []
};

// Extract the current style inside the format and returns the index.
const getStyle = (c) => {
    const bgColor = c?.fill?.fgColor?.argb || null;

    let styleIndex = null;

    // For the cell with null value and has no background color, no need to add format.
    if ((c?.value === null || c?.text === "0") && bgColor !== null) {

        styleIndex = format.styles.findIndex(style => {
            const length = Object.keys(style).length;
            if (style.bgcolor === `#${bgColor.toLocaleLowerCase().slice(0, -2)}` && length === 1) {
                return style;
            }
        });

        if (styleIndex === -1) {
            format.styles.push({
                "bgcolor": `#${bgColor.toLocaleLowerCase().slice(0, -2)}`,
            });
            styleIndex = format.styles.length - 1;
        }
    }

    // Cell with number value and has a background color, add it to the format as an object.
    if (c?.text?.match(/^[0-9]+$/) && c?.text !== "0") {
        styleIndex = format.styles.findIndex(style => style.bgcolor === `#${bgColor.toLocaleLowerCase().slice(0, -2)}` && style.format === "numberNoDecimal");

        if (styleIndex === -1) {
            format.styles.push({
                "format": "numberNoDecimal",
                "bgcolor": `#${bgColor.toLocaleLowerCase().slice(0, -2)}`,
            });
            styleIndex = format.styles.length - 1;
        }
    }

    // Cell with percentage value and has a background color, add it to the format as an object.
    if (c.numFmt === '0.00%' && c?.value !== null) {
        styleIndex = format.styles.findIndex(style => style.bgcolor === `#${bgColor.toLocaleLowerCase().slice(0, -2)}` && style.format === "percentNoDecimal");

        if (styleIndex === -1) {
            format.styles.push({
                "format": "percentNoDecimal",
                "bgcolor": `#${bgColor.toLocaleLowerCase().slice(0, -2)}`,
            });
            styleIndex = format.styles.length - 1;

        }
    }

    // Cell is bold, add it to the format as an object.
    if (c.font?.bold) {
        styleIndex = format.styles.findIndex(style => style?.font?.bold === true);
        
        if (styleIndex === -1) {
            format.styles.push({
                "font": {
                    "bold": true,
                }
            });
            styleIndex = format.styles.length - 1;
        }
    }

    return styleIndex;
}

const converter = async (source, output) => {

    // Read the source file.
    const workbook = new ExceltoJSON.Workbook();
    await workbook.xlsx.readFile(source);

    const worksheet = workbook.getWorksheet(1);
    format.name = worksheet.name;
    format.freeze = worksheet.freeze != null ? worksheet.freeze : "A1";

    // Iterate over all rows (including empty rows) in a worksheet
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {

            format.rows[rowNumber - 1] = format.rows[rowNumber - 1] || { cells: {} };

            // if the cell has a formula, add it to the format as an object.
            if (cell.formula) {
                format.rows[rowNumber - 1].cells[colNumber - 1] = {
                    "text": `=${cell.formula.toLocaleLowerCase()}`
                }

                const index = getStyle(cell);
                if (index !== null) {
                    format.rows[rowNumber - 1].cells[colNumber - 1].style = index;
                }
            }

            // if the cell is a number, add it to the format as an object.
            if (typeof cell.value === 'number') {
                format.rows[rowNumber - 1].cells[colNumber - 1] = {
                    "text": String(cell.value)
                }

                const index = getStyle(cell);
                if (index !== null) {
                    format.rows[rowNumber - 1].cells[colNumber - 1].style = index;
                }
            }

            // if the cell is a string, add it to the format as an object.
            if (typeof cell.value === 'string') {
                format.rows[rowNumber - 1].cells[colNumber - 1] = {
                    "text": cell.value
                }

                const index = getStyle(cell);
                if (index !== null) {
                    format.rows[rowNumber - 1].cells[colNumber - 1].style = index;
                }
            }

            // if the cell is null, add it to the format as an object.
            if (cell.value === null) {

                format.rows[rowNumber - 1].cells[colNumber - 1] = {
                    "text": ""
                }

                const index = getStyle(cell);
                if (index !== null) {
                    format.rows[rowNumber - 1].cells[colNumber - 1].style = index;
                }
            }

            // fill the cols object with the column width of every column.
            const colWidth = worksheet.getColumn(colNumber)?.width;
            if (colWidth) {
                format.cols[colNumber - 1] = {
                    "width": colWidth
                }
            }

        });
    })

    console.log("Convertion Done successfully. Output JSON file is stored in '"+ output + "'");
    fs.writeFileSync(output, JSON.stringify(format, null, 2));
    return format;
}

if(process.argv.length != 4){
    console.log("Invalid number of arguments")
}

else{
    let source = process.argv[2]
    let output = process.argv[3]
    if(!source.includes(".xlsx") || !output.includes(".json")) {
        console.log("Invalid format(s) .xlsx or .json required");
    }
    else{
        converter(source, output);
    }

}


module.exports = { converter, getStyle };
