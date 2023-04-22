'use strict';
const xlsx = require('xlsx');
const utils = xlsx.utils;
const fs = require('fs');


function readConfig(file) {
    // 設定ファイルを読み込む
    try {
        return JSON.parse(fs.readFileSync(file));
    } catch(err) {
        console.error(`cannot find ${file}`);
        process.exit(1);
    }
}


function write(file, string) {
    // 文字列をファイルに書き込む

    fs.appendFile(file, string, (err) => {
        if (err) throw err;
    });
}


function createEmptyFile(filePath) {
    // 空のファイルを作成する
    // 既に存在する場合は、削除して再作成する

    if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
    }
    fs.writeFileSync(filePath, '');
  }


function alphabetIndex(letter) {
    // エクセルのカラムとしてAやBと受け取ったものをインデックスに変換して返す
    // js-xlsxの仕様に合わせて、A→0、B→1、....とする

    const alphabets = 'abcdefghijklmnopqrstuvwxyz';
    const lowerCaseLetter = letter.toLowerCase();
    const index = alphabets.indexOf(lowerCaseLetter);
    return index === -1 ? '' : index;
}


function readExcel(excelFile, sheetName) {
    // Excelファイルの指定シートの内容をjs-excelで読み込む

    const workbook = xlsx.readFile(excelFile);
    const worksheet = workbook.Sheets[sheetName];

    return worksheet;
}


function inputRange(worksheet) {
    // 値が入力されているセルの範囲を取得する
    // 以下のような形式で取得される
    // {
    //    s: { c: 0, r: 0 },  //開始セルアドレス
    //    e: { c: 5, r: 1 }   //終了セルアドレス
    // }

    const rangeAddress = worksheet['!ref'];
    const rangeDict = utils.decode_range(rangeAddress);

    return rangeDict;
}


function getValue(worksheet, rowIdx, colIdx) {
    // シートの指定のセルの値を取得する
    // ここで、Idxは0始まりの数値を指す。たとえばA1のrowIdx, ColIdxは共に0である。

    let adress = utils.encode_cell({c: colIdx, r: rowIdx});
    let cell = worksheet[adress];
    return cell.w;
}


module.exports = {
    readConfig: readConfig,
    write: write,
    createEmptyFile: createEmptyFile,
    alphabetIndex: alphabetIndex,
    readExcel: readExcel,
    inputRange: inputRange,
    getValue: getValue
};