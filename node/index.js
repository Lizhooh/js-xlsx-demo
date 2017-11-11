const XLSX = require('xlsx');
const util = require('util');

// const defaultWorkBook = require('./default-workbook.json');
const defaultWorkBook = XLSX.utils.book_new();

async function runtime(name = 'run', cb) {
    console.time(name);
    try {
        typeof cb === 'function' && await cb();
    }
    catch (err) {
        console.log(err);
    }
    console.timeEnd(name);
}

runtime('run', async () => {
    // 默认配置项
    const workbook = defaultWorkBook;

    // 自定义内容
    const ws = XLSX.utils.aoa_to_sheet([
        ['学校', '姓名', '学号'],

        ['电子神技大学', '张三', 'A-201701010211'],
        ['电子神技大学', '李四', 'A-201701010212'],
        ['电子神技大学', '王五', 'A-201701010213'],
    ]);

    workbook.SheetNames.unshift('test');
    workbook.Sheets['test'] = ws;

    // 生成 xlsx 文件
    XLSX.writeFile(workbook, './output.xlsx');
});
