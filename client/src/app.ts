import ReadExcel = require("./ReadExcel");
import fs = require("fs");

let args = process.argv;
let xlspath = args[2] || "E:/H5/excelDecoder/test.xlsx";
start();
function start() {
    let str = ReadExcel(xlspath);
    if (str) {
        let path = xlspath.replace("xlsx", "json");
        fs.writeFileSync(path, str);
    }
}