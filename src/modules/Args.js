module.args = (function(){
  var mod = {};

  // import
  var utility = module.utility;
  utility.checkImport('Args', 'logger');
  var logger = module.logger;

  /**
   * @return {Array<String>}
   */
  mod.getArgs = function(){
    var ws_args = WScript.Arguments;
    if(ws_args.Length === 0){
      logger.info(getMsg().no_args);
      logger.print();
      WScript.Quit();
    }

    var ar_args = [];
    for(var nu_arg = 0; nu_arg < ws_args.Length; nu_arg++){
      var st_arg = ws_args.Item(nu_arg);
      logger.trace(st_arg);
      ar_args.push(st_arg);
    }
    return ar_args;
  };

  /**
   * private
   * @return {Object}
   */
  function getMsg(){
    return (function(){
      var m  ={};
      m.no_args = 'Please drag & drop any file!';
      return m;
    })();
  }

  return mod;
})();

