import fs = require("fs");

import xlsx = require("xlsx");

function readExcel(path: string) {
    var data = xlsx.read(path, { type: "file" });
    var sheetData = data["Sheets"]["export"];
    let reg = /\d/;
    let tmpdic: any = {};
    let b2 = sheetData["B2"];
    let mainkey = b2 ? b2["v"] : "";
    for (let key in sheetData) {
        reg.lastIndex = 0;
        let val = reg.exec(key);
        if (val) {
            let inx = val.index;
            let rowKey = key.substr(0, inx);
            let rowinx = +key.substr(inx);
            if (rowinx > 4 && rowKey != "A") {
                let tmp = tmpdic[rowinx];
                if (!tmp) {
                    tmp = tmpdic[rowinx] = {};
                }
                let tmpkey = sheetData[rowKey + "4"]["v"];
                let tv = sheetData[key]["v"];
                let type = sheetData[rowKey + "1"]["v"];
                if (type == "int") {
                    tv = +tv;
                } else {
                    tv = tv + "";
                }
                tmp[tmpkey] = tv;
            }
        }
    }
    let outstr = "";
    if (mainkey) {
        let dic: any = {};
        for (let key in tmpdic) {
            let tmp = tmpdic[key];
            dic[tmp[mainkey]] = tmp;
        }
        outstr = JSON.stringify(dic);
    } else {
        let arr = [];
        for (let key in tmpdic) {
            arr.push(tmpdic[key]);
        }
        outstr = JSON.stringify(arr);
    }
    return outstr;
}

export = readExcel;