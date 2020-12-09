import XLSX from 'xlsx-style';


export default function downLoadExcel(data, type, config, ele) {
    console.log(config, 'excel配置参数');
    var arr = data;
    if (ele) {
        var ws = XLSX.utils.table_to_sheet(ele);
    } else {
        var ws = XLSX.sheetFromAoa(arr);
    }
    if (config.merge.length != 0) {
        ws['!merges'] = config.merge;
    }
    ws['!cols'] = config.size.cols;

    if (config.myStyle.all) { //作用在所有单元格的样式，必须在最顶层，然后某些特殊样式在后面的操作中覆盖基本样式
        Object.keys(ws).forEach((item, index) => {
            if (ws[item].t) {
                ws[item].s = config.myStyle.all;
            }
        });
    }
    if (config.myStyle.headerColor) {
        if (config.myStyle.headerLine) {
            let line = config.myStyle.headerLine;
            let p = /^[A-Z]{1}[A-Z]$/;
            Object.keys(ws).forEach((item, index) => {
                for (let i = 1; i <= line; i++) {
                    if (item.replace(i, '').length == 1 || (p.test(item.replace(i, '')))) {
                        let myStyle = getDefaultStyle();
                        myStyle.fill.fgColor.rgb = config.myStyle.headerColor;
                        myStyle.font.color.rgb = config.myStyle.headerFontColor;
                        ws[item].s = myStyle;
                    }
                }
            });

        }
    }
    if (config.myStyle.specialCol) {
        config.myStyle.specialCol.forEach((item, index) => {
            item.col.forEach((item1, index1) => {
                Object.keys(ws).forEach((item2, index2) => {
                    if (item.expect && item.s) {
                        if (item2.includes(item1) && !item.expect.includes(item2)) {
                            ws[item2].s = item.s;
                        }
                    }
                    if (item.t) {
                        if (item2.includes(item1) && item2.t) {
                            ws[item2].t = item.t;
                        }
                    }
                });
            });

        });
    }
    if (config.myStyle.bottomColor) {
        if (config.myStyle.rowCount) {
            Object.keys(ws).forEach((item, index) => {
                if (item.indexOf(config.myStyle.rowCount) != -1) {
                    let myStyle1 = getDefaultStyle();
                    myStyle1.fill.fgColor.rgb = config.myStyle.bottomColor;
                    ws[item].s = myStyle1;
                }
            })
        }
    }
    if (config.myStyle.colCells) {
        Object.keys(ws).forEach((item, index) => {
            if (item.split('')[0] === config.myStyle.colCells.col && item !== 'C1' && item !== 'C2') {
                ws[item].s = config.myStyle.colCells.s;
            }
        })
    }
    if (config.myStyle.mergeBorder) { //对导出合并单元格无边框的处理
        let arr = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
        let range = config.myStyle.mergeBorder;
        range.forEach((item, index) => {
            if (item.s.c == item.e.c) { //行相等,横向合并
                let star = item.s.r;
                let end = item.e.r;
                for (let i = star + 1; i <= end; i++) {
                    ws[arr[i] + (Number(item.s.c) + 1)] = {
                        s: ws[arr[star] + (Number(item.s.c) + 1)].s
                    }
                }
            } else { //列相等，纵向合并
                let star = item.s.c;
                let end = item.e.c;
                for (let i = star + 1; i <= end; i++) {
                    ws[arr[item.s.r] + (i + 1)] = {
                        s: ws[arr[item.s.r] + (star + 1)].s
                    }
                }
            }
        });
    }
    if (config.myStyle.specialHeader) {
        config.myStyle.specialHeader.forEach((item, index) => {
            Object.keys(ws).forEach((item1, index1) => {
                if (item.cells.includes(item1)) {
                    ws[item1].s.fill = {
                        fgColor: {
                            rgb: item.rgb
                        }
                    };
                    if (item.color) {
                        ws[item1].s.font.color = {
                            rgb: item.color
                        };
                    }
                }
            });
        });
    }
    if (config.myStyle.heightLightColor) {
        Object.keys(ws).forEach((item, index) => {
            if (ws[item].t === 's' && ws[item].v && ws[item].v.includes('%') && !item.includes(config.myStyle.rowCount)) {
                if (Number(ws[item].v.replace('%', '')) < 100) {
                    ws[item].s = {
                        fill: {
                            fgColor: {
                                rgb: config.myStyle.heightLightColor
                            }
                        },
                        font: {
                            name: "Meiryo UI",
                            sz: 11,
                            color: {
                                auto: 1
                            }
                        },
                        border: {
                            top: {
                                style: 'thin',
                                color: {
                                    auto: 1
                                }
                            },
                            left: {
                                style: 'thin',
                                color: {
                                    auto: 1
                                }
                            },
                            right: {
                                style: 'thin',
                                color: {
                                    auto: 1
                                }
                            },
                            bottom: {
                                style: 'thin',
                                color: {
                                    auto: 1
                                }
                            }
                        },
                        alignment: {
                            /// 自动换行
                            wrapText: 1,
                            // 居中
                            horizontal: "center",
                            vertical: "center",
                            indent: 0
                        }
                    }
                }

            }
        });
    }
    if (config.myStyle.rowCells) {
        config.myStyle.rowCells.row.forEach((item, index) => {
            Object.keys(ws).forEach((item1, index1) => {
                let num = Number(dislodgeLetter(item1));
                if (num === item) {
                    ws[item1].s = config.myStyle.rowCells.s;
                }
            });
        });

    }

    Object.keys(ws).forEach((item, index) => {
        if (ws[item].t === 's' && !ws[item].v) {
            ws[item].v = '-';
        }
    });

    // let newStyle = ws['A1'].s;
    // newStyle.bold = true;
    var blob = IEsheet2blob(ws);
    if (IEVersion() !== 11) {
        openDownloadXLSXDialog(blob, `${type}.xlsx`);
    } else {
        window.navigator.msSaveOrOpenBlob(blob, `${type}.xlsx`);
    }


}

function dislodgeLetter(str) {
    var result;
    var reg = /[a-zA-Z]+/; //[a-zA-Z]表示bai匹配字母，dug表示全局匹配
    while (result = str.match(reg)) { //判断str.match(reg)是否没有字母了
        str = str.replace(result[0], ''); //替换掉字母  result[0] 是 str.match(reg)匹配到的字母
    }
    return str;
}



function IEsheet2blob(sheet, sheetName) {
    try {
        new Uint8Array([1, 2]).slice(0, 2);
    } catch (e) {
        //IE或有些浏览器不支持Uint8Array.slice()方法。改成使用Array.slice()方法
        Uint8Array.prototype.slice = Array.prototype.slice;
    }
    sheetName = sheetName || 'sheet1';
    var workbook = {
        SheetNames: [sheetName],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;
    // 生成excel的配置项
    var wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    var blob = new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    });
    // 字符串转ArrayBuffer
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}



function getDefaultStyle() {
    let defaultStyle = {
        fill: {
            fgColor: {
                rgb: ''
            }
        },
        font: {
            name: "Meiryo UI",
            sz: 11,
            color: {
                rgb: ''
            },
            bold: true
        },
        border: {
            top: {
                style: 'thin',
                color: {
                    auto: 1
                }
            },
            left: {
                style: 'thin',
                color: {
                    auto: 1
                }
            },
            right: {
                style: 'thin',
                color: {
                    auto: 1
                }
            },
            bottom: {
                style: 'thin',
                color: {
                    auto: 1
                }
            }
        },
        alignment: {
            /// 自动换行
            wrapText: 1,
            // 居中
            horizontal: "center",
            vertical: "center",
            indent: 0
        }
    };
    return defaultStyle;
}

function IEVersion() {
    var userAgent = navigator.userAgent; //取得浏览器的userAgent字符串  
    var isIE = userAgent.indexOf("compatible") > -1 && userAgent.indexOf("MSIE") > -1; //判断是否IE<11浏览器  
    var isEdge = userAgent.indexOf("Edge") > -1 && !isIE; //判断是否IE的Edge浏览器  
    var isIE11 = userAgent.indexOf('Trident') > -1 && userAgent.indexOf("rv:11.0") > -1;
    if (isIE) {
        var reIE = new RegExp("MSIE (\\d+\\.\\d+);");
        reIE.test(userAgent);
        var fIEVersion = parseFloat(RegExp["$1"]);
        if (fIEVersion == 7) {
            return 7;
        } else if (fIEVersion == 8) {
            return 8;
        } else if (fIEVersion == 9) {
            return 9;
        } else if (fIEVersion == 10) {
            return 10;
        } else {
            return 6; //IE版本<=7
        }
    } else if (isEdge) {
        return 'edge'; //edge
    } else if (isIE11) {
        return 11; //IE11  
    } else {
        return -1; //不是ie浏览器
    }
}



function openDownloadXLSXDialog(url, saveName) {
    if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url); // 创建blob地址
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
    var event;
    if (window.MouseEvent) event = new MouseEvent('click');
    else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
}