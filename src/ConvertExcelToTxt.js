import 'babel-polyfill';
import {Utility} from './modules/Utility.js';
import {Logger, LevelListAll} from './modules/Logger.js';
import {Args} from './modules/Args.js';
import {ExcelAdapter} from './modules/ExcelAdapter.js';
import {FileSystem} from './modules/FileSystem.js';

const logger = new Logger();
logger.config.output_level = LevelListAll.ALL;
const excel = new ExcelAdapter(logger);
const fileSystem = new FileSystem(logger);

Utility.echo('Wait!');

// check args
const ar_args = Args.getArgs();
if(ar_args.length <= 2){
  Utility.echo('ERR! require over 2 args');
  WScript.Quit();
}
const st_arg_basedir = ar_args[0];
if(!FileSystem.folderExists(st_arg_basedir)){
  Utility.echo('ERR! ' + st_arg_basedir + 'is not dir');
  WScript.Quit();
}
const st_arg_outdir = ar_args[1];
if(!FileSystem.folderExists(st_arg_outdir)){
  Utility.echo('ERR! ' + st_arg_outdir + 'is not dir');
  WScript.Quit();
}
const ar_ignore_dir = ar_args.slice(2);


// get file list
fileSystem.config.ignore_dir_reg = ar_ignore_dir
const ar_filepath = fileSystem.getFiles(st_arg_basedir).map((fi)=>{return fi.path});


// convert
excel.executeExcel(ar_filepath, (ws_book)=>{
  const st_bookname = st_arg_outdir + ws_book.Name.replace(/.xls(x|m)?/, '') + '_';
  excel.excelEachSheet(ws_book, (ws_sheet)=>{
    const st_sheetfile = st_bookname + ws_sheet.Name + '.txt';

    // saveing only active sheet
    ws_sheet.Activate;
    ws_book.SaveAs(
        st_sheetfile /* FileName */
        , -4158 /* FileFormat *//* .txt */
        );
  });
});

logger.print();
Utility.echo('Done!');
