'use strict';
const utils = require('./utils.js');
const os = require('os');


// TOMLに出力する結果を保存する変数
let output = '';


function writeToml(tomlFile) {
    // 最終結果をファイルに出力する
    utils.write(tomlFile, output);
}


function addTableTag(tableName) {
    // TOMLファイルにテーブル名を出力する
    // 例：writeTableTag('user') -> [user] と出力する
    output += `[${tableName}]${os.EOL}`;
}


function addTableArrayTag(tableName) {
    // TOMLファイルにテーブル配列名を出力する
    // 例：writeTableTag('user.name') -> [[user.name]] と出力する
    output += `[[${tableName}]]${os.EOL}`;
}


function addDataWithTableArray(worksheet, targetColumns, tableName) {
    // 指定したカラムのデータを出力する
    //   worksheet: js-xlsxで取得したExcelシート
    //   targetColumns: 出力するカラムの配列　(例) targetColumns = ['A', 'B']
    //   tableName: 出力するTomlのテーブル配列名

    // ヘッダー行の値を保持しておく変数
    let headerNames = [];

    // Excelの行ごとに処理を行う
    const rangeDict = utils.inputRange(worksheet);
    for (let rowIdx = rangeDict.s.r; rowIdx <= rangeDict.e.r; rowIdx++) {

        // ヘッダー行(1行目)の場合は、ヘッダー行に記載されたカラム名を取得する
        if (!headerNames.length) {
            for (let column of targetColumns) {
                let colIdx = utils.alphabetIndex(column);
                headerNames.push(utils.getValue(worksheet, rowIdx, colIdx));
            }
            continue;
        }

        // テーブル配列名を出力する
        addTableArrayTag(tableName);

        // 指定したカラムの値を出力する
        for (let i = 0; i < targetColumns.length; i++) {
            let colIdx = utils.alphabetIndex(targetColumns[i]);
            let key = headerNames[i];
            let value = utils.getValue(worksheet, rowIdx, colIdx);

            output += `${key} = ${value}${os.EOL}`;
        }
    }
}


module.exports = {
    writeToml: writeToml,
    addTableTag: addTableTag,
    addTableArrayTag: addTableArrayTag,
    addDataWithTableArray: addDataWithTableArray
};