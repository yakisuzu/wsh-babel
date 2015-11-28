module.excelAdapter = (function(){
  var mod = {};

  // import
  var utility = module.utility;
  utility.checkImport('ExcelAdapter', 'logger');
  var logger = module.logger;

  /**
   * @constructor
   */
  mod.ExcelAdapter = function(){
    this.excel_use_ignore_reg = false;
    this.excel_ignore_reg = [];
    this.excel_error_reg = [/#N\/A/, /#REF!/, /[a-z,A-Z]:\\(.+\\)*.+\.xlsx?/, /\[.+\.xlsx?\]/];
  };

  (function(p){
    /**
     * @param {String} st_value
     * @return {Boolean}
     */
    p.isErrorValue = function(st_value){
      var self = this;

      if(self.excel_use_ignore_reg){
        for(var i = 0; i < self.excel_ignore_reg.length; i++){
          var st_ignore = self.excel_ignore_reg[i];

          // when not found regex, value is error
          if(st_value.search(st_ignore) === -1){
            return true;
          }
        }

      }else{
        for(var i = 0; i < self.excel_error_reg.length; i++){
          var st_err = self.excel_error_reg[i];

          // when contains error, value is error
          if(st_value.search(st_err) !== -1){
            return true;
          }
        }
      }
      return false;
    };

    /**
     * @param {Workbook} ws_book
     */
    p.excelErrorNameDelete = function(ws_book){
      var self = this;

      var ws_names = ws_book.Names;
      logger.traceBuild(getMsg().excel_name_count, [ws_names.Count]);

      var ar_del_name = [];
      mod.eachItem(ws_names, function(ws_name){
        logger.traceBuild(getMsg().excel_name_value, [ws_name.Name, ws_name.Value]);

        ws_name.Visible = true;

        // add delete array
        if(p.isErrorValue.call(self, ws_name.Value)){
          ar_del_name.push(ws_name);
        }
      });

      // execute error name delete
      logger.traceBuild(getMsg().excel_name_hit_count, [ar_del_name.length]);
      while(ar_del_name.length !== 0){
        var ws_del = ar_del_name.pop();
        logger.traceBuild(getMsg().excel_name_delete_value, [ws_del.Name, ws_del.Value]);

        ws_del.Delete();
      }
    };

    /**
     * @param {Workbook} ws_book
     */
    p.excelErrorFormatDelete = function(ws_book){
      var self = this;

      mod.eachSheet(ws_book, function(ws_sheet){
        var ws_fcs = ws_sheet.Cells.FormatConditions;
        logger.traceBuild(getMsg().excel_fc_count, [ws_fcs.Count]);

        var ar_del_fc = [];
        mod.eachItem(ws_fcs, function(ws_fc){
          var fc = getFc(ws_fc);
          logger.traceBuild(getMsg().excel_fc_value, [fc.Formula1, fc.Formula2]);

          // TODO check ws_fc.Formula2
          // add delete array
          if(p.isErrorValue.call(self, fc.Formula1)){
            ar_del_fc.push(ws_fc);
          }
        });

        // execute error name delete
        logger.traceBuild(getMsg().excel_fc_hit_count, [ar_del_fc.length]);
        while(ar_del_fc.length !== 0){
          var ws_del = ar_del_fc.pop();
          var fc = getFc(ws_del);
          logger.traceBuild(getMsg().excel_fc_delete_value, [fc.Formula1, fc.Formula2]);

          ws_del.Delete();
        }
      });

      function getFc(ws_fc){
        var f1 = '';
        var f2 = '';
        try{
          f1 = ws_fc.Formula1;
          f2 = ws_fc.Formula2;
        }catch(e){
          // If it is not set, an error is thrown
        }

        return {'Formula1' : f1, 'Formula2' : f2};
      }
    };

  })(mod.ExcelAdapter.prototype);

  /**
   * @callback excelAdapter~fu_execute
   * @param {Workbook} ws_book
   */
  /**
   * @param {Array<String>} ar_files
   * @param {excelAdapter~fu_execute} fu_execute
   */
  mod.executeExcel = function(ar_files, fu_execute){
    var ws_excel;
    try{
      ws_excel = openExcel();

      // repeat arg file
      for(var i = 0; i < ar_files.length; i++){
        var st_arg = ar_files[i];
        // ignore extention at pattern
        if(st_arg.search(/^.+\.xlsx?$/) === -1){
          logger.warn(getMsg().no_support);
          continue;
        }

        // execute execl function
        var ws_book;
        try{
          ws_book = ws_excel.Workbooks.Open(
              /* FileName */ st_arg,
              /* UpdateLinks */ 0,
              /* ReadOnly */ false,
              /* Format */ null,
              /* Password */ null,
              /* WriteResPassword */ null,
              /* IgnoreReadOnlyRecommended */ true
              );
          logger.traceBuild(getMsg().excel_book_open, [st_arg]);

          fu_execute(ws_book);

          ws_book.Close(true);
          logger.traceBuild(getMsg().excel_book_close, [st_arg]);
        }catch(e){
          logger.errorBuild(getMsg().error, [st_arg]);
          ws_book.Close(false);
          throw e;
        }
      }
    }catch(e){
      utility.dump(e);
    }finally{
      try{
        closeExcel(ws_excel);
      }catch(e){
        utility.dump(e);
      }
    }
  };


  /**
   * @callback excelAdapter~fu_execute
   * @param {Object<Excel>} ws_item
   */
  /**
   * @param {Object<Excel>} ws
   * @param {excelAdapter~fu_execute} fu_execute
   */
  mod.eachItem = function(ws, fu_execute){
    for(var nu_ws = 1; nu_ws <= ws.Count; nu_ws++){
      fu_execute(ws.Item(nu_ws));
    }
  };

  /**
   * @callback excelAdapter~fu_execute
   * @param {Worksheet} ws_sheet
   */
  /**
   * @param {Workbook} ws_book
   * @param {excelAdapter~fu_execute} fu_execute
   */
  mod.eachSheet = function(ws_book, fu_execute){
    var ws_sheets = ws_book.Worksheets;
    logger.traceBuild(getMsg().excel_sheet_count, [ws_sheets.Count]);
    mod.eachItem(ws_sheets, function(ws_sheet){
      logger.traceBuild(getMsg().excel_sheet_name, [ws_sheet.Name]);
      fu_execute(ws_sheet);
    });
  };

  /**
   * private
   * @return {Excel}
   */
  function openExcel(){
    var ws_excel;
    try{
      ws_excel = WScript.CreateObject('Excel.Application');
      ws_excel.Visible = false;
    }catch(e){
      throw e;
    }
    logger.trace(getMsg().excel_start);
    return ws_excel;
  }

  /**
   * private
   * @param {Excel} ws_excel
   */
  function closeExcel(ws_excel){
    try{
      if(ws_excel !== undefined){
        ws_excel.Quit();
      }
    }catch(e){
      throw e;
    }

    logger.trace(getMsg().excel_end);
  }

  /**
   * private
   * @return {Object}
   */
  function getMsg(){
    return (function(){
      var m  ={};
      m.no_support = 'Support xls or xlsx!';
      m.error = 'Error! {0}';

      m.excel_start = 'Excel Start!';
      m.excel_end = 'Excel quit!';

      m.excel_book_open = 'Book open {0}';
      m.excel_book_close = 'Book close {0}';

      m.excel_sheet_count = 'Sheet {Count : {0}}';
      m.excel_sheet_name = 'Sheet {Name : {0}}';

      m.excel_name_count = 'Name {Count : {0}}';
      m.excel_name_value = 'Name {Name : {0}, Value : {1}}';
      m.excel_name_hit_count = 'Name hit count = {0}';
      m.excel_name_delete_value = 'Delete! ' + m.excel_name_value;

      m.excel_fc_count = 'Fc {Count : {0}}';
      m.excel_fc_value = 'Fc {Formula1 : {0}, Formula2 : {1}}';
      m.excel_fc_hit_count = 'Fc hit count = {0}';
      m.excel_fc_delete_value = 'Delete! ' + m.excel_fc_value;

      return m;
    })();
  }

  return mod;
})();

