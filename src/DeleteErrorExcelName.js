import 'babel-polyfill';
import {Utility} from './modules/Utility.js';
import {Logger} from './modules/Logger.js';
import {Args} from './modules/Args.js';
import {ExcelAdapter} from './modules/ExcelAdapter.js';

Args.logger = Logger;
ExcelAdapter.logger = Logger;

let excel = new ExcelAdapter();
excel.read_only = false;
excel.save = true;

Logger.setting.output_level = Logger.level.ALL;

Utility.echo('Wait!');

excel.executeExcel(Args.getArgs(), (ws_book)=>{
  excel.excelErrorNameDelete(ws_book);
});

Logger.print();
Utility.echo('Done!');
