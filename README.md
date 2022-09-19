## mxlsx插件使用文档

```
npm i mexcel
```

### excel导入

| 方法 | 描述 | 类型 |
| ----------- | ----------- | ----------- |
| importExcel | 导入 具体参数查看下方[IExcelImportConfig](#import-setting) | IExcelImportConfig |

<h5 id="import-setting">IExcelImportConfig</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| file | 文件 | File \| Blob | 是 | -- | -- |
| keys | 导入数据key值 | "A" \| number \| string[] | 否 | -- | -- |
| keyRow | 读取excel表格中某行数据作为key值， ***keys 存在的情况下 keyRow 不生效*** | number | 否 | -- | -- |
| dataRow | 读取excel表格中从某行开始 | number | 否 | -- | 0 |
| customKey | 是否采用自定义key ***keys、keyRow 和 dataRow 只有在 customKey 为 true 时才生效*** | boolean | 否 | true \| false | true |
| onProgress | 导入文件进度 | (event: ProgressEvent<FileReader>) => void | 否 | -- | -- |

### excel导出

| 方法 | 描述 | 类型 |
| ----------- | ----------- | ----------- |
| exportExcel | 导出 | (excelData: IExcel, success?: () => void, fail?: (err: unknown) => void) => void |

##### 参数说明

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| fileName | 文件名 | string | 否 | -- | excel |
| fileExtention | 文件格式 | string | 否 | xlsx \| xls | xlsx |
| sheets | 工作表数据[工作表参数](#sheet-setting) | ISheet | 是 | -- | -- |

- <h5 id="sheet-setting">工作表参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| title | 表格标题，自动合并单元格 | string | 否 | -- | -- |
| titleStyle | 表格标题样式 具体参数查看下方[单元格样式参数](#cell-setting) ***在title存在的情况才生效*** | CellStyle | 否 | -- | -- |
| tHeaders | 表格表头 ***可以配置多表头，自动合并单元格*** | INSArr[] | 否 | -- | -- |
| tHeaderStyle | 表格表头样式 具体参数查看下方[单元格样式参数](#cell-setting) ***在tHeaders存在的情况才生效*** | CellStyle | 否 | -- | -- |
| table | 表格数据 | ITable[] | 是 | -- | -- |
| cols | 列样式配置 具体参数查看下方[列样式参数](#col-setting) | ColInfo[] | 否 | -- | -- |
| titleRow | 表格标题行样式配置 [行样式参数](#row-setting) ***在title存在的情况才生效*** | RowInfo | 否 | -- | -- |
| headerRows | 列样式配置 具体参数查看下方[行样式参数](#row-setting) ***在tHeaders存在的情况才生效*** | RowInfo[] | 否 | -- | -- |
| row | 行样式配置 具体参数查看下方[行样式参数](#row-setting) | RowInfo | 否 | -- | -- |
| merges | 单元格合并配置 具体参数查看下方[合并参数](#merge-setting) | Range[] | 否 | -- | -- |
| keys | 表格数据key值描述 ***参数数据影响标题和表头的自动合并*** | INSArr | 是 | -- | -- |
| sheetName | 工作表名字 | string | 否 | -- | sheet + 索引值 |
| globalStyle | 单元格全局样式 具体参数查看下方[单元格样式参数](#cell-setting) | CellStyle | 否 | -- | -- |
| cellStyle | 具体单元格自定义样式 具体参数查看下方[单元格样式参数](#cell-setting) | ICellStyle | 否 | -- | -- |

- <h5 id="cell-setting">单元格样式参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| font | 字体样式 具体参数查看下方[字体样式参数](#font-setting) | CellStyle.font | 否 | -- | -- |
| alignment | 对齐方式 具体参数查看下方[对齐方式参数](#alignment-setting) | CellStyle.alignment | 否 | -- | -- |
| border | 边框样式 具体参数查看下方[边框样式参数](#border-setting) | CellStyle.border | 否 | -- | -- |
| fill | 背景样式 具体参数查看下方[背景样式参数](#fill-setting) | CellStyle.fill | 否 | -- | -- |
| numFmt | 数据格式 | string | 否 | 0 \| 0.00% \| 0.0% \| 0.00%;\\(0.00%\\);\\-;@ \| m/dd/yy | 0 |

- <h5 id="font-setting">字体样式参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| bold | 粗细 | boolean | 否 | true \| false | false |
| color | 字体颜色 具体参数查看下方[颜色参数](#color-setting) | CellStyleColor | 否 | -- | -- |
| italic | 斜体 | boolean | 否 | true \| false | false |
| name | 字体 | string | 否 | -- | Calibri |
| sz | 字体大小 | number | 否 | -- | -- |
| strike | 删除线 | boolean | 否 | true \| false | false |
| underline | 下划线 | boolean | 否 | true \| false | false |
| vertAlign | 上下标 | string | 否 | "superscript" \| "subscript" | null |

- <h5 id="alignment-setting">对齐方式参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| horizontal | 横向对齐 | string | 否 | left \| center \| right | left |
| vertical | 纵向对齐 | string | 否 | top \| center \| bottom | bottom |
| textRotation | 文字旋转 | number | 否 | 0 - 180 \| 255 | 0 |
| wrapText | 是否换行 | boolean | 否 | true \| false | false |

- <h5 id="border-setting">边框样式参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| top | 上边 具体参数查看下方[颜色参数](#color-setting) [边框属性参数](#border-style-setting) | { color: CellStyleColor; style?: BorderType } | 否 | -- | -- |
| bottom | 下边 | { color: CellStyleColor; style?: BorderType } | 否 | -- | -- |
| left | 左边 | { color: CellStyleColor; style?: BorderType } | 否 | -- | -- |
| right | 右边 | { color: CellStyleColor; style?: BorderType } | 否 | -- | -- |
| diagonal | 对角线 | { color: CellStyleColor; style?: BorderType; diagonalUp?: boolean; diagonalDown?: boolean } | 否 | -- | -- |


- <h5 id="fill-setting">背景样式参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| bgColor | 背景色 具体参数查看下方[颜色参数](#color-setting) | CellStyleColor | 否 | -- | -- |
| fgColor | 前景色 具体参数查看下方[颜色参数](#color-setting) | CellStyleColor | 否 | -- | -- |
| patternType | 模式 | string | 否 | solid \| none | solid |

- <h5 id="merge-setting">合并参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| s | 开始单元格 具体参数查看下方[表格位置参数](#address-setting) | CellAddress | 否 | -- | -- |
| e | 结束单元格 具体参数查看下方[表格位置参数](#address-setting) | CellAddress | 否 | -- | -- |

- <h5 id="address-setting">表格位置参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| c | 列数 | number | 是 | 0 - max | -- |
| r | 行数 | number | 是 | 0 - max | -- |

- <h5 id="row-setting">行样式参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| hidden | 是否隐藏行 | boolean | 否 | true \| false | false |
| hpx | 行高 ***屏幕像素高度*** | number | 否 | 0 - max | -- |
| hpt | 行高 ***以点为单位的高度*** | number | 否 | 0 - max | -- |
| level | 分组折叠 | number | 否 | -- | -- |

- <h5 id="col-setting">列样式参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| hidden | 是否隐藏列 | boolean | 否 | true \| false | false |
| width | 列宽 ***最大数字宽度中的宽度*** | number | 否 | 0 - max | -- |
| wpx | 列宽 ***Excel的“最大数字宽度”中的宽度，width*256是整数*** | number | 否 | 0 - max | -- |
| wch | 列宽 ***字符宽度*** | number | 否 | 0 - max | -- |
| level | 分组折叠 | number | 否 | -- | -- |
| MDW | 列宽 ***Excel 的“最大数字宽度”单位，始终为整数*** | number | 否 | 0 - max | -- |

- <h5 id="color-setting">颜色参数</h5>

| 参数 | 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- | ----------- |
| rgb | 颜色值  ***hex值不要带#号*** | string | 否 | -- | -- |
| theme | 主题色 ***theme与rgb不同时存在，theme覆盖rgb*** | number | 否 | -- | -- |
| tint | 透明度  ***在theme存在是生效*** | -1.0 - 1.0 | 否 | -- | -- |

- <h5 id="border-style-setting">边框属性参数</h5>

| 描述 | 类型 | 必填 | 可选值 | 默认值 |
| ----------- | ----------- | ----------- | ----------- | ----------- |
| 边框样式 | BorderType | 否 | dashDot \| dashDotDot \| dashed \| dotted \| hair \| medium \| mediumDashDot \| mediumDashDotDot \| mediumDashed \| slantDashDot \| thick \| thin | -- |

```js
import { exportExcel, IExcel } from "mexcel";

const excelData: IExcel = {
    sheets: [
        {
            title: "学生列表",
            tHeaders: [["学号", "姓名", "班级", "考试成绩"], ["", "", "", "语文", "数学", "英语"]],
            merges: [
                {
                    s: { c: 0, r: 1 },
                    e: { c: 0, r: 2 }
                },
                {
                    s: { c: 1, r: 1 },
                    e: { c: 1, r: 2 }
                },
                {
                    s: { c: 2, r: 1 },
                    e: { c: 2, r: 2 }
                },
                {
                    s: { c: 3, r: 1 },
                    e: { c: 5, r: 1 }
                }
            ],
            table: [
                {
                    no: "1",
                    name: "李浩",
                    class: "二年级2班",
                    yuwen: 93,
                    shuxue: 95,
                    yingyu: 88
                },
                {
                    no: "2",
                    name: "王明",
                    class: "二年级1班",
                    yuwen: 89,
                    shuxue: 99,
                    yingyu: 90
                }
            ],
            cols: [
                {
                    wpx: 50
                },
                {
                    wpx: 100
                },
                {
                    wpx: 200
                },
                {
                    wpx: 50
                },
                {
                    wpx: 50
                },
                {
                    wpx: 50
                }
            ],
            titleRow: {
                hpx: 60
            },
            headerRows: [
                {
                    hpx: 40
                },
                {
                    hpx: 40
                }
            ],
            row: {
                hpx: 30
            },
            keys: ["no", "name", "class", "yuwen", "shuxue", "yingyu"],
            sheetName: "学生列表",
            globalStyle: {
                font: {
                    sz: 18
                },
                alignment: {
                    horizontal: "center",
                    vertical: "center",
                    wrapText: true
                },
                border: {
                    top: { style: "thin", color: {} },
                    right: { style: "thin", color: {} },
                    bottom: { style: "thin", color: {} },
                    left: { style: "thin", color: {} }
                }
            },
            titleStyle: {
                font: {
                    sz: 22,
                    color: {
                        rgb: "f60000"
                    }
                },
                alignment: {
                    horizontal: "center",
                    vertical: "center",
                    wrapText: true
                },
                border: {
                    top: { style: "thin", color: {} },
                    right: { style: "thin", color: {} },
                    bottom: { style: "thin", color: {} },
                    left: { style: "thin", color: {} }
                }
            }
        }
    ],
    fileName: "学生信息"
};

exportExcel(excelData);
```