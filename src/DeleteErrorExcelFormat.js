import 'babel-polyfill';
import {Utility} from './modules/Utility.js';
import {Logger, LoggerStaticConfig, LevelListAll} from './modules/Logger.js';
import {Args} from './modules/Args.js';
import {ExcelAdapter} from './modules/ExcelAdapter.js';

const excel = new ExcelAdapter(Logger);

excel.config.read_only = false;
excel.config.save = true;

LoggerStaticConfig.output_level = LevelListAll.ALL;

Utility.echo('Wait!');

excel.executeExcel(Args.getArgs(), (ws_book)=>{
  excel.excelErrorFormatDelete(ws_book);
});

Logger.print();
Utility.echo('Done!');
