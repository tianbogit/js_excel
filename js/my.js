function change() {
    var body = $("#body");
    body.append(body.children().first());
}

$(document).ready(function () {

    readFile();
});

// js读取解析Excel
// 定义一个carData,用来保存读取到的数据

var readFile = function () {
    var wb;//读取完成的数据
    var rABS = false; //是否将文件读取为二进制字符串

    function fixdata(data) { //文件流转BinaryString
        var o = "",
            l = 0,
            w = 10240;
        for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
        o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
        return o;
    }

    $("#file").change(function () {
        if (!this.files) {
            return;
        }
        var f = this.files[0];
        var reader = new FileReader();
        reader.onload = function (e) {
            var data = e.target.result;
            if (rABS) {
                wb = XLSX.read(btoa(fixdata(data)), {
                    type: 'base64'
                });
            } else {
                wb = XLSX.read(data, {
                    type: 'binary'
                });
            }
            readWorkbook(wb)
            // // carData就是我们需要的JSON数据
            // carData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[3]]);
            // create(carData);
        };
        if (rABS) {
            reader.readAsArrayBuffer(f);
        } else {
            reader.readAsBinaryString(f);
        }
    })

};

function readWorkbook(workbook) {
    var sheetNames = workbook.SheetNames; // 工作表名称集合
    var worksheet = workbook.Sheets[sheetNames[3]]; // 这里我们只读取第一张sheet
    var csv = XLSX.utils.sheet_to_csv(worksheet);
    document.getElementById('demo').innerHTML = csv2table(csv);
    var MyMar1 = setInterval(change, 1000);
    $("#demo").mouseenter(function () {
        clearInterval(MyMar1)
    });
    $("#demo").mouseleave(function () {
        MyMar1 = setInterval(change, 1000);
    });
}

// 将csv转换成简单的表格，会忽略单元格合并，在第一行和第一列追加类似excel的索引
function csv2table(csv) {
    var html = '<table class="table table-bordered" style="text-align: center">';
    var rows = csv.split('\n');
    delete rows[0]
    rows.pop(); // 最后一行没用的
    rows.forEach(function (row, idx) {
        var columns = row.split(',');
        // columns.unshift(idx+1); // 添加行索引
      
        if (idx == 1) {
            html += '<thead><tr>';
            columns.forEach(function (column) {
                html += '<td>' + column + '</td>';
            });
            html += '</tr></thead>';
        }
        else if (idx == 2) {
            html += '<tbody id="body">';
            html += '<tr>';
            columns.forEach(function (column, cli) {
                if (cli > 0) {
                    html += '<td contenteditable="true">' + column + '</td>';
                } else {
                    html += '<td>' + column + '</td>';
                }
            });
            html += '</tr>';
        }
        else {
            html += '<tr>';
            columns.forEach(function (column, cli) {
                if (cli > 0) {
                    html += '<td contenteditable="true">' + column + '</td>';
                } else {
                    html += '<td>' + column + '</td>';
                }
            });
            html += '</tr>';
        }

    });
    html += '</tbody></table>';
    return html;
}
