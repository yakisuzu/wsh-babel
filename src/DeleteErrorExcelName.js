import 'babel-polyfill';
import {Utility} from './modules/Utility.js';
import {Logger} from './modules/Logger.js';
import {Args} from './modules/Args.js';
import {ExcelAdapter} from './modules/ExcelAdapter.js';

let excel = new ExcelAdapter();
excel.read_only = false;
excel.save = true;

let msg = {};
msg.start = 'Wait!';
msg.end =  'Done!';

Logger.setting.output_level = Logger.level.ALL;

Logger.info(msg.start);
Logger.print();

excel.executeExcel(Args.getArgs(), (ws_book)=>{
  excel.excelErrorNameDelete(ws_book);
});

Logger.info(msg.end);
Logger.print();
