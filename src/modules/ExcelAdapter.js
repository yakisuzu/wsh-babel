import {Utility} from './Utility.js';

// ---------------
// private
// ---------------

const msg = (()=>{
  let m  ={};
  m.no_support = 'Support xls, xlsx, xlsm!';
  m.error = 'Error! {0}';

  m.excel_start = 'Excel Start!';
  m.excel_end = 'Excel quit!';

  m.excel_book_open = 'Book open {0}';
  m.excel_book_close = 'Book close {0}';

  m.excel_sheet_count = 'Sheet {Count : {0}}';
  m.excel_sheet_name = 'Sheet {Name : {0}}';

  m.excel_name_count = 'Name {Count : {0}}';
  m.excel_name_value = 'Name {Name : {0}, Value : {1}}';
  m.excel_name_delete_count = 'Delete! Name count = {0}';
  m.excel_name_delete_value = 'Delete! ' + m.excel_name_value;

  m.excel_fc_count = 'Fc {Count : {0}}';
  m.excel_fc_value = 'Fc {Formula1 : {0}, Formula2 : {1}}';
  m.excel_fc_delete_count = 'Delete! Fc count = {0}';
  m.excel_fc_delete_value = 'Delete! ' + m.excel_fc_value;

  return m;
})();

/**
 * @callback excelAdapter~fu_execute
 * @param {Object<Excel>} ws_item
 */
/**
 * @param {Object<Excel>} ws
 * @param {excelAdapter~fu_execute} fu_execute
 */
function eachItem(ws, fu_execute){
  for(let nu_ws = 1; nu_ws <= ws.Count; nu_ws++){
    fu_execute(ws.Item(nu_ws));
  }
}

// ---------------
// public
// ---------------

/**
 *
 */
class Config{
  constructor(){
    this.read_only = true;
    this.save = false;
    this.excel_use_ignore_reg = false;
    this.excel_ignore_reg = [];
    this.excel_error_reg = [/#N\/A/, /#REF!/, /[a-z,A-Z]:\\(.+\\)*.+\.xlsx?/, /\[.+\.xlsx?\]/];
  }
}

/**
 *
 */
class ExcelAdapter{

  /**
   * @constructor
   */
  constructor(logger){
    this.Logger = logger;

    this.config = new Config();
  }

  /**
   * @callback excelAdapter~fu_execute
   * @param {Worksheet} ws_sheet
   */
  /**
   * @param {Workbook} ws_book
   * @param {excelAdapter~fu_execute} fu_execute
   */
  eachSheet(ws_book, fu_execute){
    let ws_sheets = ws_book.Worksheets;
    this.Logger.trace(msg.excel_sheet_count, [ws_sheets.Count]);
    eachItem(ws_sheets, (ws_sheet)=>{
      this.Logger.trace(msg.excel_sheet_name, [ws_sheet.Name]);
      fu_execute(ws_sheet);
    });
  }

  /**
   * @param {String} st_value
   * @return {Boolean}
   */
  isErrorValue(st_value){
    if(this.config.excel_use_ignore_reg){
      for(let st_ignore of this.config.excel_ignore_reg){
        // when not found regex, value is error
        if(st_value.search(st_ignore) === -1){
          return true;
        }
      }

    }else{
      for(let st_err of this.config.excel_error_reg){
        // when contains error, value is error
        if(st_value.search(st_err) !== -1){
          return true;
        }
      }
    }
    return false;
  }

  /**
   * @param {Workbook} ws_book
   */
  excelErrorNameDelete(ws_book){
    let ws_names = ws_book.Names;
    this.Logger.trace(msg.excel_name_count, [ws_names.Count]);

    let ar_del_name = [];
    eachItem(ws_names, (ws_name)=>{
      this.Logger.trace(msg.excel_name_value, [ws_name.Name, ws_name.Value]);

      ws_name.Visible = true;

      // add delete array
      if(this.isErrorValue(ws_name.Value)){
        ar_del_name.push(ws_name);
      }
    });

    // execute error name delete
    for(let ws_del of ar_del_name){
      this.Logger.trace(msg.excel_name_delete_value, [ws_del.Name, ws_del.Value]);

      ws_del.Delete();
    }
    this.Logger.trace(msg.excel_name_delete_count, [ar_del_name.length]);
  }

  /**
   * @param {Workbook} ws_book
   */
  excelErrorFormatDelete(ws_book){
    this.eachSheet(ws_book, (ws_sheet)=>{
      let ws_fcs = ws_sheet.Cells.FormatConditions;
      this.Logger.trace(msg.excel_fc_count, [ws_fcs.Count]);

      let ar_del_fc = [];
      eachItem(ws_fcs, function(ws_fc){
        let fc = getFc(ws_fc);
        this.Logger.trace(msg.excel_fc_value, [fc.Formula1, fc.Formula2]);

        // TODO check ws_fc.Formula2
        // add delete array
        if(this.isErrorValue(fc.Formula1)){
          ar_del_fc.push(ws_fc);
        }
      });

      // execute error name delete
      for(let ws_del of ar_del_fc){
        let fc = getFc(ws_del);
        this.Logger.trace(msg.excel_fc_delete_value, [fc.Formula1, fc.Formula2]);

        ws_del.Delete();
      }
      this.Logger.trace(msg.excel_fc_delete_count, [ar_del_fc.length]);
    });

    function getFc(ws_fc){
      let f1 = '';
      let f2 = '';
      try{
        f1 = ws_fc.Formula1;
        f2 = ws_fc.Formula2;
      }catch(e){
        // If it is not set, an error is thrown
      }

      return new (class{
        constructor(){
          this.Formula1 = f1;
          this.Formula2 = f2;
        }
      })();
    }
  }

  /**
   * @callback excelAdapter~fu_execute
   * @param {Workbook} ws_book
   */
  /**
   * @param {Array<String>} ar_files
   * @param {excelAdapter~fu_execute} fu_execute
   */
  executeExcel(ar_files, fu_execute){
    let ws_excel;
    try{
      ws_excel = WScript.CreateObject('Excel.Application');
      ws_excel.Visible = false;
      this.Logger.trace(msg.excel_start);

      // repeat arg file
      for(let st_arg of ar_files){
        // ignore extention at pattern
        if(st_arg.search(/^.+\.xls(x|m)?$/) === -1){
          this.Logger.warn(msg.no_support);
          continue;
        }

        // execute execl function
        let ws_book;
        try{
          this.Logger.trace(msg.excel_book_open, [st_arg]);
          ws_book = ws_excel.Workbooks.Open(
              /* FileName */ st_arg,
              /* UpdateLinks */ 0,
              /* ReadOnly */ this.config.read_only,
              /* Format */ null,
              /* Password */ null,
              /* WriteResPassword */ null,
              /* IgnoreReadOnlyRecommended */ true
              );

          fu_execute(ws_book);

          if(this.config.save){
            ws_book.Save();
          }
          ws_book.Close(this.config.save);
          this.Logger.trace(msg.excel_book_close, [st_arg]);
        }catch(e){
          this.Logger.error(msg.error, [st_arg]);
          ws_book.Close(false);
          throw e;
        }
      }
    }catch(e){
      Utility.dump(e);
    }finally{
      try{
        if(ws_excel !== undefined){
          ws_excel.Quit();
        }
        this.Logger.trace(msg.excel_end);
      }catch(e){
        Utility.dump(e);
      }
    }
  }
}

export {ExcelAdapter};
