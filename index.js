'use strict';
const utils = require('./utils.js');
const logic = require('./logic.js');


// パラメータを取得
const args = process.argv.slice(2);
if (args.length !== 1) {
    console.error('Usage: node index.js [config file]');
    process.exit(1);
}
const configFile = args[0];

// 設定ファイルの読込
const config = utils.readConfig(configFile);

// EXCELファイルの読込
const worksheet = utils.readExcel(config.EXCEL_FILE, config.SHEET_NAME);

// TOMLファイルへの変換
utils.createEmptyFile(config.TOML_FILE);
logic.addTableTag(config.TABLE);
logic.addDataWithTableArray(worksheet, [config.NAME_COLUMN, config.AGE_COLUMN], config.TABLE_ARRAY);
logic.writeToml(config.TOML_FILE);

