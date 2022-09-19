import xlsxStyle, { CellStyle, ColInfo, Range, RowInfo, WorkSheet } from "xlsx-js-style";
/**
 * 默认数据
 */
const defaultExcel: IExcel = {
    sheets: [],
    fileName: "excel",
    fileExtention: "xlsx"
};

export type INSArr = (string | number)[];

export interface ITable {
    [key: string | number]: string | number;
}

export interface ICellStyle {
    [key: string]: CellStyle;
}

export interface ISheet {
    title?: string;
    titleStyle?: CellStyle;
    tHeaders?: INSArr[];
    tHeaderStyle?: CellStyle;
    table: ITable[];
    cols?: ColInfo[];
    titleRow?: RowInfo;
    headerRows?: RowInfo[];
    row?: RowInfo;
    merges?: Range[];
    keys: INSArr;
    sheetName?: string;
    globalStyle?: CellStyle;
    cellStyle?: ICellStyle;
}

export interface IExcel {
    sheets: ISheet[];
    fileName?: string;
    fileExtention?: "xls" | "xlsx";
}

// 数字转字母
const convert = (num: number): string => {
    if (num < 26 && num > -1) {
        return String.fromCharCode(num + 65)
    } else if (num >= 26) {
        const left = num % 26;
        return convert(Math.floor(num / 26) - 1) + convert(left);
    }
    return "";
}

// 处理表数据
const dealExportTable = (sheet: ISheet, sheetData: INSArr[]) => {
    // 存在title
    if (sheet.title) {
        const titleArr = Array(sheet.keys.length).fill("");
        titleArr[0] = sheet.title;
        sheetData.push(titleArr);
    }

    // 存在tHeaders
    if (sheet.tHeaders) {
        sheet.tHeaders.forEach(tHeader => {
            const headerArr = Array(sheet.keys.length).fill("");
            tHeader.forEach((text, index) => {
                headerArr[index] = text;
            })
            sheetData.push(headerArr);
        });
    }

    if (sheet.table) {
        sheet.table.forEach((line) => {
            const lineData: INSArr = [];
            if (sheet.keys) {
                sheet.keys.forEach(key => {
                    lineData.push(line[key] || "");
                });
            }
            sheetData.push(lineData);
        });
    }
};

// 处理样式
const setCellStyle = (sheet: ISheet, worksheet: WorkSheet) => {
    // 设置表格样式
    Object.keys(worksheet).forEach((key) => {
        // 非!开头的属性都是单元格
        if (!key.startsWith("!")) {
            let cellStyle: CellStyle = {};
            if (sheet.globalStyle) {
                cellStyle = sheet.globalStyle;
            }

            if (sheet.cellStyle && sheet.cellStyle[key]) {
                cellStyle = { ...cellStyle, ...sheet.cellStyle };
            }

            worksheet[key].s = cellStyle;
        }
    });

    // 设置标题样式
    if (sheet.title && sheet.titleStyle) {
        worksheet["A1"].s = sheet.titleStyle;
    }

    // 设置header样式
    if (sheet.tHeaderStyle && sheet.tHeaders) {
        // header初始行数值
        const initialRow = sheet.title ? 2 : 1;
        sheet.tHeaders.forEach((tHeader, row) => {
            tHeader.forEach((text, col) => {
                const cellKey = convert(col) + (row + initialRow);
                worksheet[cellKey].s = sheet.tHeaderStyle;
            });
        });
    }
}

// 设置合并
const setMerges = (sheet: ISheet, worksheet: WorkSheet) => {
    let merges: Range[] = [];

    if (sheet.title) {
        merges.push({
            s: { c: 0, r: 0 },
            e: { c: sheet.keys.length - 1, r: 0 }
        });
    }

    if (sheet.merges) {
        merges = merges.concat(sheet.merges);
    }

    worksheet["!merges"] = merges;
}

// 设置行与列
const setRowAndCol = (sheet: ISheet, worksheet: WorkSheet) => {
    worksheet["!cols"] = sheet.cols || []; // 设置工作表列样式
    
    const rowNum = (sheet.table ? sheet.table.length : 0) + (sheet.tHeaders ? sheet.tHeaders.length : 0) + (sheet.title ? 1 : 0);
    const rows = Array(rowNum).fill(sheet.row || {});

    if (sheet.title && sheet.titleRow) {
        rows[0] = sheet.titleRow;
    }

    if (sheet.tHeaders && sheet.headerRows) {
        const initialRow = sheet.title ? 1 : 0;
        sheet.headerRows.forEach((headerRow, index) => {
            rows[initialRow + index] = headerRow;
        });
    }

    worksheet["!rows"] = rows;
};

export const exportExcel = (excelData: IExcel, success?: () => void, fail?: (err: unknown) => void) => {
    const newExcelData = {
        ...defaultExcel,
        ...excelData
    };
    // 创建工作簿
    const workbook = xlsxStyle.utils.book_new();

    try {
        excelData.sheets.forEach((sheet, index) => {
            const sheetData: INSArr[] = [];
    
            // 填充数据
            dealExportTable(sheet, sheetData);
    
            const worksheet = xlsxStyle.utils.aoa_to_sheet(sheetData);
    
            // 设置行与列
            setRowAndCol(sheet, worksheet);
    
            // 样式处理
            setCellStyle(sheet, worksheet);
    
            // 设置合并
            setMerges(sheet, worksheet);
    
            xlsxStyle.utils.book_append_sheet(
                workbook,
                worksheet,
                sheet.sheetName || `sheet${index + 1}`
            );
        });
    
        xlsxStyle.writeFile(
            workbook,
            `${newExcelData.fileName}.${newExcelData.fileExtention}`
        );

        success && success();
    } catch (err) {
        fail && fail(err);
    }
};
