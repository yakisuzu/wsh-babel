import 'babel-polyfill';
import {Utility} from './modules/Utility.js';
import {Logger, LevelListAll} from './modules/Logger.js';
import {Args} from './modules/Args.js';
import {ExcelAdapter} from './modules/ExcelAdapter.js';

const logger = new Logger();
logger.config.output_level = LevelListAll.ALL;
const excel = new ExcelAdapter(logger);

excel.config.read_only = false;
excel.config.save = true;

Utility.echo('Wait!');

excel.executeExcel(Args.getArgs(), (ws_book)=>{
  excel.excelErrorNameDelete(ws_book);
});

Logger.print();
Utility.echo('Done!');
