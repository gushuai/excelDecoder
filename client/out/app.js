"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var ReadExcel = require("./ReadExcel");
var fs = require("fs");
var args = process.argv;
var xlspath = args[2] || "E:/H5/excelDecoder/test.xlsx";
start();
function start() {
    var str = ReadExcel(xlspath);
    if (str) {
        var path = xlspath.replace("xlsx", "json");
        fs.writeFileSync(path, str);
    }
}
