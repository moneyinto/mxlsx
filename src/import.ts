import { read, utils } from "xlsx";

export interface IExcelResult {
    sheetName: string;
    data: unknown[];
}

export interface IExcelOpts {
    header: "A" | number | string[];
    range?: string;
}

export interface IExcelImportConfig {
    file: File | Blob;
    keys?: "A" | number | string[];
    keyRow?: number;
    dataRow?: number;
    customKey?: boolean;
    onProgress?: (event: ProgressEvent<FileReader>) => void;
}

export const importExcel = ({ file, onProgress, keys, keyRow, dataRow, customKey = true }: IExcelImportConfig): Promise<IExcelResult[]> => {
    return new Promise((resolve, reject) =>  {
        try {
            const reader = new FileReader();
            reader.onprogress = (event: ProgressEvent<FileReader>) => {
                onProgress && onProgress(event);
            };

            reader.onload = (event: ProgressEvent<FileReader>) => {
                const data = event.target?.result;
                const excel = read(data, {
                    type: "binary"
                });

                const json: IExcelResult[] = [];

                excel.SheetNames.forEach(item => {
                    const excelData = excel.Sheets[item];

                    const opts: IExcelOpts = {
                        header: []
                    };
        
                    if (keys) {
                        opts.header = keys;
                    } else if (typeof keyRow !== "undefined") {
                        const keys = [];
                        Object.keys(excelData).forEach((key) => {
                            // 非!开头的属性都是单元格
                            const keyMatch = key.match(/[0-9]+/);
                            if (!key.startsWith("!") && keyMatch && keyMatch[0] === keyRow.toString()) {
                                keys.push(excelData[key]);
                            }
                        });
                    } else {
                        opts.header = "A";
                    }
        
                    if (dataRow) {
                        opts.range = `A${dataRow + 1}:${excelData["!ref"]?.split(":")[1]}`;
                    }

                    const data = utils.sheet_to_json(
                        excelData,
                        customKey ? opts : {}
                    );
                    json.push({ sheetName: item, data });
                });

                resolve(json);
            }

            reader.readAsBinaryString(file);
        } catch(err) {
            reject(err);
        }
    })
}