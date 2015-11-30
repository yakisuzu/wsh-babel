import {Utility} from './Utility.js';

// ---------------
// private
// ---------------

const LevelList = (()=>{
  let level = {};
  level.TRACE = 1;
  level.DEBUG = 2;
  level.INFO = 3;
  level.WARN = 4;
  level.ERROR = 5;
  level.FATAL = 6;
  return level;
})();

const LevelListAll = (()=>{
  let level = {};
  level.ALL = 0;
  for(let key in LevelList){
    level[key] = LevelList[key];
  }
  level.OFF = 9;
  return level;
})();

let config = new (class{
  constructor(){
    this.output_level = LevelList.INFO;
    this.header = (st_level)=>{return '[' + st_level + ']';};
    this.linefeed = '\n';
    this.output = (st_msg)=>{Utility.echo(st_msg);};
  }
})();

/**
 *
 */
class OutputText{
  constructor(st_level, st_text){
    this.level = st_level;
    this.text = st_text;
  }
}

let ar_output_stock = [];
/**
 * @param {String} st_level
 * @param {String} st_text
 */
function outputPush(st_level, st_text){
  if(config.output_level <= LevelList[st_level]){
    ar_output_stock.push(new OutputText(st_level, st_text));
  }
}

// ---------------
// public
// ---------------

/**
 * TODO setting format
 * TODO show line no
 */
class Logger{

  static getLevel(){
    return LevelListAll;
  }

  static getConfig(){
    return config;
  }

  /**
   * @param {String}
   * @param {Array<String>}
   */
  static trace(st_msg, ar_args=[]){outputPush('trace', Utility.buildMsg(st_msg, ar_args));}
  static debug(st_msg, ar_args=[]){outputPush('debug', Utility.buildMsg(st_msg, ar_args));}
  static info (st_msg, ar_args=[]){outputPush('info', Utility.buildMsg(st_msg, ar_args));}
  static warn (st_msg, ar_args=[]){outputPush('warn', Utility.buildMsg(st_msg, ar_args));}
  static error(st_msg, ar_args=[]){outputPush('error', Utility.buildMsg(st_msg, ar_args));}
  static fatal(st_msg, ar_args=[]){outputPush('fatal', Utility.buildMsg(st_msg, ar_args));}

  /**
   * @param {void}
   */
  static print(){
    let st_output_string = '';
    for(let outputText of ar_output_stock){
      st_output_string += config.header(outputText.level) + outputText.text + config.linefeed;
    }

    if(st_output_string !== ''){
      config.output(st_output_string);
    }
  }
}

export {Logger};
