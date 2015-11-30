import 'babel-polyfill';
import {Utility} from './modules/Utility.js';
import {Logger} from './modules/Logger.js';
import {Args} from './modules/Args.js';
import {ExcelAdapter} from './modules/ExcelAdapter.js';

let args = new Args(Logger);
let excel = new ExcelAdapter(Logger);

excel.read_only = false;
excel.save = true;

Logger.getConfig().output_level = Logger.getLevel().ALL;

Utility.echo('Wait!');

excel.executeExcel(args.getArgs(), (ws_book)=>{
  excel.excelErrorFormatDelete(ws_book);
});

Logger.print();
Utility.echo('Done!');
