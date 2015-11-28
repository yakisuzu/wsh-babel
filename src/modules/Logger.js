// TODO setting format
// TODO show line no
module.logger = (function(){
  var mod = {};

  // import
  var utility = module.utility;

  // static private
  var ar_output_stock = [];

  /**
   * Object
   */
  mod.level = (function(){
    var level = {};
    level.ALL = 0;
    for(var key in getLevelList()){
      level[key] = getLevelList()[key];
    }
    level.OFF = 9;
    return level;
  })();

  // static private
  var ob_output_setting = (function(){
    var o = {};
    o.output_level = mod.level.INFO;
    o.header = function(st_level){return '[' + st_level + ']';};
    o.linefeed = '\n';
    o.output = function(st_msg){utility.echo(st_msg);};
    return o;
  })();

  /**
   * Declare the module in a loop
   * @param {String} st_text
   */
  for(var key in getLevelList()){
    mod[key.toLowerCase()] = (function(key){
      return function(st_text){
        outputPush(key, st_text);
      };
    })(key);
  }

  /**
   * Declare the module in a loop
   * @param {String} st_msg
   * @param {Array<Object>} st_args
   */
  for(var key in getLevelList()){
    mod[key.toLowerCase() + 'Build'] = (function(key){
      return function(st_msg, ar_args){
        outputPush(key, utility.buildMsg(st_msg, ar_args));
      };
    })(key);
  }

  /**
   * @param {void}
   */
  mod.print = function(){
    var st_output_string = '';
    while(ar_output_stock.length !== 0){
      var ob_output = ar_output_stock.shift();
      st_output_string += ob_output_setting.header(ob_output.level) + ob_output.text + ob_output_setting.linefeed;
    }

    if(st_output_string !== ''){
      ob_output_setting.output(st_output_string);
    }
  };

  /**
   * Object
   */
  mod.set = (function(setting){
    var mods = {};

    /**
     * @param {Number} nu_level_value
     */
    mods.outputLevel = function(nu_level_value){
      setting.output_level = nu_level_value;
    }

    /**
     * @callback logger.set~fu_header
     * @param {String} st_level
     */
    /**
     * @param {logger.set~fu_header} fu_header
     */
    mods.header = function(fu_header){
      setting.header = fu_header;
    };

    /**
     *@param {String} st_linefeed
     */
    mods.linefeed = function(st_linefeed){
      setting.linefeed = st_linefeed;
    };

    /**
     * @callback logger.set~fu_output
     * @param {String} st_msg
     */
    /**
     * @param {logger.set~fu_output} fu_output
     */
    mods.output = function(fu_output){
      setting.output = fu_output;
    };

    return mods;
  })(ob_output_setting);

  /**
   * private
   * @return {Object}
   */
  function getLevelList(){
    return (function(){
      var list = {};
      list.TRACE = 1;
      list.DEBUG = 2;
      list.INFO = 3;
      list.WARN = 4;
      list.ERROR = 5;
      list.FATAL = 6;
      return list;
    })();
  }

  /**
   * private
   * @param {String} st_level
   * @param {String} st_text
   */
  function outputPush(st_level, st_text){
    if(ob_output_setting.output_level <= getLevelList()[st_level]){
      ar_output_stock.push({'level' : st_level, 'text' : st_text});
    }
  }

  return mod;
})();

