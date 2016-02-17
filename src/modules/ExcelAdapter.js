import {Utility} from './Utility.js';

// ---------------
// private
// ---------------

const msg = (()=>{
  const m = {};
  m.no_support = 'Support extention is only xls, xlsx, xlsm!';
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
 * @param {Object<Excel>} ws
 */
function* eachItem(ws){
  for(let nu_ws = 1; nu_ws <= ws.Count; nu_ws++){
    yield ws.Item(nu_ws);
  }
}

/**
 * @param {ExcelAdapter} self
 * @param {Object<Workbook>} ws_book
 */
function* eachSheet(self, ws_book){
  const ws_sheets = ws_book.Worksheets;
  self.logger.trace(msg.excel_sheet_count, [ws_sheets.Count]);
  for(let ws_sheet of eachItem(ws_sheets)){
    self.logger.trace(msg.excel_sheet_name, [ws_sheet.Name]);
    yield ws_sheet;
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
  constructor(log){
    this.logger = log;

    this.config = new Config();
  }

  /**
   * @param {String} st_value
   * @return {Boolean}
   */
  isErrorValue(st_value){
    if(this.config.excel_use_ignore_reg){
      // when not found regex, value is error
      if(this.config.excel_ignore_reg.some((reg)=>{
        return st_value.search(reg) === -1;
      })){
        return true;
      }

    }else{
      // when contains error, value is error
      if(this.config.excel_error_reg.some((reg)=>{
        return st_value.search(reg) !== -1;
      })){
        return true;
      }
    }
    return false;
  }

  /**
   * @param {Object<Workbook>} ws_book
   */
  excelErrorNameDelete(ws_book){
    const ws_names = ws_book.Names;
    this.logger.trace(msg.excel_name_count, [ws_names.Count]);

    const ar_del_name = [];
    for(let ws_name of eachItem(ws_names)){
      this.logger.trace(msg.excel_name_value, [ws_name.Name, ws_name.Value]);

      ws_name.Visible = true;

      // add delete array
      if(this.isErrorValue(ws_name.Value)){
        ar_del_name.push(ws_name);
      }
    }

    // execute error name delete
    for(let ws_del of ar_del_name){
      this.logger.trace(msg.excel_name_delete_value, [ws_del.Name, ws_del.Value]);

      ws_del.Delete();
    }
    this.logger.trace(msg.excel_name_delete_count, [ar_del_name.length]);
  }

  /**
   * @param {Object<Workbook>} ws_book
   */
  excelErrorFormatDelete(ws_book){
    const getFc = (ws_fc)=>{
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

    for(let ws_sheet of eachSheet(this, ws_book)){
      const ws_fcs = ws_sheet.Cells.FormatConditions;
      this.logger.trace(msg.excel_fc_count, [ws_fcs.Count]);

      const ar_del_fc = [];
      for(let ws_fc of eachItem(ws_fcs)){
        const fc = getFc(ws_fc);
        this.logger.trace(msg.excel_fc_value, [fc.Formula1, fc.Formula2]);

        // TODO check ws_fc.Formula2
        // add delete array
        if(this.isErrorValue(fc.Formula1)){
          ar_del_fc.push(ws_fc);
        }
      }

      // execute error name delete
      for(let ws_del of ar_del_fc){
        const fc = getFc(ws_del);
        this.logger.trace(msg.excel_fc_delete_value, [fc.Formula1, fc.Formula2]);

        ws_del.Delete();
      }
      this.logger.trace(msg.excel_fc_delete_count, [ar_del_fc.length]);
    }
  }

  /**
   * @callback excelAdapter~fu_execute
   * @param {Object<Worksheet>} ws_sheet
   */
  /**
   * @param {Object<Workbook>} ws_book
   * @param {excelAdapter~fu_execute} fu_execute
   */
  excelEachSheet(ws_book, fu_execute){
    for(let ws_sheet of eachSheet(this, ws_book)){
      fu_execute(ws_sheet);
    }
  }

  /**
   * @callback excelAdapter~fu_execute
   * @param {String} st_sheetname
   * @param {Number} nu_row
   * @param {Number} nu_col
   */
  /**
   * @param {Object<Workbook>} ws_book
   * @param {excelAdapter~fu_execute} fu_execute
   */
  excelEachCell(ws_book, fu_execute){
    for(let ws_sheet of eachSheet(this, ws_book)){
      const ROW_MIN = ws_sheet.UsedRange.Row;
      const ROW_MAX = ROW_MIN + ws_sheet.UsedRange.Rows.Count - 1;
      const COL_MIN = ws_sheet.UsedRange.Columns;
      const COL_MAX = COL_MIN + ws_sheet.UsedRange.Columns.Count - 1;
      const SHEET_NAME = ws_sheet.Name;

      for(let nu_row = ROW_MIN; nu_row <= ROW_MAX; nu_row++){
        for(let nu_col = COL_MIN; nu_col <= COL_MAX; nu_col++){
          fu_execute(SHEET_NAME, nu_row, nu_col);
        }
      }
    }
  }

  /**
   * @callback excelAdapter~fu_execute
   * @param {Object<Workbook>} ws_book
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
      this.logger.trace(msg.excel_start);

      // repeat arg file
      for(let st_arg of ar_files){
        // ignore extention at pattern
        if(st_arg.search(/^.+\.xls(x|m)?$/) === -1){
          this.logger.warn(msg.no_support);
          continue;
        }

        // execute execl function
        let ws_book;
        try{
          this.logger.trace(msg.excel_book_open, [st_arg]);
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
          this.logger.trace(msg.excel_book_close, [st_arg]);
        }catch(e){
          this.logger.error(msg.error, [st_arg]);
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
        this.logger.trace(msg.excel_end);
      }catch(e){
        Utility.dump(e);
      }
    }
  }
}

export {ExcelAdapter};
