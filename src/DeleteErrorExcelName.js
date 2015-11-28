<job>
  <script language="JScript" src="./module/Module.js"></script>
  <script language="JScript" src="./module/Logger.js"></script>
  <script language="JScript" src="./module/Args.js"></script>
  <script language="JScript" src="./module/ExcelAdapter.js"></script>
  <script language="JScript">
    // import
    var utility = module.utility;
    var logger = module.logger;
    var args = module.args;
    var excelAdapter = module.excelAdapter;
    var excel = new excelAdapter.ExcelAdapter();

    var msg = {};
    msg.start = 'Wait!';
    msg.end =  'Done!';

    logger.set.outputLevel(logger.level.ALL);

    logger.info(msg.start);
    logger.print();

    excelAdapter.executeExcel(
      args.getArgs()
      , function(ws_book){
        excel.excelErrorNameDelete(ws_book);
      }
    );

    logger.info(msg.end);
    logger.print();
  </script>
</job>
