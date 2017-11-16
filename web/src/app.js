
window.onload = function () {

    var $ = function (select, ctx) {
        ctx = ctx || document;
        return ctx.querySelector(select);
    };

    var file = $('#file');
    var drop = $('#drop');
    var edit = $('#edit');
    var view = $('#view');
    var derive = $('#derive');

    var excelView = $('#excel-view');


    /**
     * 读取
     */
    function readExcelFile(filedata, cb) {
        // https://developer.mozilla.org/zh-CN/docs/Web/API/FileReader
        var reader = new FileReader();

        var types = [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ];

        if (types.indexOf(filedata.type) === -1) {
            return alert('文件类型不是 Excel 格式');
        }

        reader.readAsBinaryString(filedata);
        reader.onload = function (e) {
            // 解析数据
            var bstr = e.target.result;
            var wb = XLSX.read(bstr, { type: 'binary', cellStyles: true });
            console.log(wb);

            // 只取第一个 sheet
            // var wsname = wb.SheetNames[0];
            // var ws = wb.Sheets[wsname];
            // 渲染
            typeof cb === 'function' && cb(wb);
        };
    }

    /**
     * 渲染数据
     */
    function render(element, filedata) {
        readExcelFile(filedata, function (wb) {
            // 只取第一个 sheet
            var wsname = wb.SheetNames[0];
            var ws = wb.Sheets[wsname];
            // 渲染
            element.innerHTML = XLSX.utils.sheet_to_html(ws);
        })
    }

    /**
     * 上传的文件
     */
    file.onchange = function (event) {
        var files = event.target.files;

        if (files && files[0]) {
            render(excelView, files[0]);
        }
    }

    /**
     * 拖拽上传
     * https://developer.mozilla.org/zh-CN/docs/Web/Events/drop
     */
    drop.addEventListener('drop', function (event) {
        event.stopPropagation();
        event.preventDefault();
        var files = event.dataTransfer.files;

        if (files && files[0]) {
            render(excelView, files[0]);
            drop.classList.remove('active');
            drop.innerText = '把 Excel 文件拖动到这个区域！';
        }
    }, false);

    function dragover(event) {
        event.stopPropagation();
        event.preventDefault();
        event.dataTransfer.dropEffect = 'copy';

        if (drop.classList.contains('active') === false) {
            drop.classList.add('active');
            drop.innerText = '松开吧！';
        }
    }

    drop.addEventListener('dragenter', dragover, false);
    drop.addEventListener('dragover', dragover, false);
    drop.addEventListener('dragleave', function (event) {
        drop.classList.remove('active');
        drop.innerText = '把 Excel 文件拖拽到这个区域里！';
    }, false);


    // --------- export ---------//
    edit.value =
        "根据表格内容生成 Excel 文件" + "\n\n" +

        "学校 | 姓名 | 学号" + "\n" +
        "--- | --- | --- " + "\n" +
        "电子神技大学 | 张三 | A-201701010211" + "\n" +
        "电子神技大学 | 李四 | A-201701010212" + "\n" +
        "电子神技大学 | 王五 | A-201701010213" + "\n" +
        "";
    view.innerHTML = marked(edit.value);
    edit.onkeyup = function (event) {
        view.innerHTML = marked(edit.value);
    }

    // 把 string 转为 ArrayBuffer
    function s2ab(str) {
        var buf = new ArrayBuffer(str.length);
        var _view = new Uint8Array(buf);
        for (var i = 0, len = str.length; i < len; i++) {
            _view[i] = str.charCodeAt(i) & 0xFF;
        }
        return buf;
    }

    // 根据表格内容，生成 Excel 文件
    derive.onclick = function (event) {
        var table = $('table', view);
        var sheet = XLSX.utils.table_to_sheet(table);

        sheet['A1'] = Object.assign(sheet['A1'], {
            // 样式？
            s: {
                fill: {
                    fgColor: { rgb: "FFFF0000" }
                }
            },
        });

        var wb = XLSX.utils.book_new({ cellStyles: true });
        XLSX.utils.book_append_sheet(wb, sheet, "SheetJS");
        // 渲染
        var wbout = XLSX.write(wb, { type: "binary", bookType: "xlsx" });
        // 保存 - https://github.com/eligrey/FileSaver.js
        saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), Date.now() + ".xlsx");
    }
}

