import 'babel-polyfill';
import {Utility} from './modules/Utility.js';
import {Logger} from './modules/Logger.js';
import {Args} from './modules/Args.js';
import {ExcelAdapter} from './modules/ExcelAdapter.js';

let excel = new ExcelAdapter(Logger);

excel.config.read_only = false;
excel.config.save = true;

Logger.getConfig().output_level = Logger.getLevel().ALL;

Utility.echo('Wait!');

excel.executeExcel(Args.getArgs(), (ws_book)=>{
  excel.excelErrorNameDelete(ws_book);
});

Logger.print();
Utility.echo('Done!');
