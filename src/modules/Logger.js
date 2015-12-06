import {Utility} from './Utility.js';

// ---------------
// private
// ---------------

const LevelList = (()=>{
  const level = {};
  level.TRACE = 1;
  level.DEBUG = 2;
  level.INFO = 3;
  level.WARN = 4;
  level.ERROR = 5;
  level.FATAL = 6;
  return level;
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
const outputPush = (st_level, st_text)=>{
  if(LoggerStaticConfig.output_level <= LevelList[st_level]){
    ar_output_stock.push(new OutputText(st_level, st_text));
  }
}

// ---------------
// public
// ---------------

const LevelListAll = (()=>{
  const level = {};
  level.ALL = 0;
  for(let key in LevelList){
    level[key] = LevelList[key];
  }
  level.OFF = 9;
  return level;
})();

const LoggerStaticConfig = (()=>{
  const c = {};
  c.output_level = LevelListAll.INFO;
  c.header = (st_level)=>{return '[' + st_level + ']';};
  c.linefeed = '\n';
  c.output = (st_msg)=>{Utility.echo(st_msg);};
  return c;
})();

/**
 * TODO setting format
 * TODO show line no
 */
class Logger{

  /**
   * @param {String}
   * @param {Array<String>}
   */
  static trace(st_msg, ar_args=[]){outputPush('TRACE', Utility.buildMsg(st_msg, ar_args));}
  static debug(st_msg, ar_args=[]){outputPush('DEBUG', Utility.buildMsg(st_msg, ar_args));}
  static info (st_msg, ar_args=[]){outputPush('INFO', Utility.buildMsg(st_msg, ar_args));}
  static warn (st_msg, ar_args=[]){outputPush('WARN', Utility.buildMsg(st_msg, ar_args));}
  static error(st_msg, ar_args=[]){outputPush('ERROR', Utility.buildMsg(st_msg, ar_args));}
  static fatal(st_msg, ar_args=[]){outputPush('FATAL', Utility.buildMsg(st_msg, ar_args));}

  /**
   *
   */
  static print(){
    LoggerStaticConfig.output(
      ar_output_stock.map((outputText)=>{
        return LoggerStaticConfig.header(outputText.level) + outputText.text
      }).join(LoggerStaticConfig.linefeed)
    );
    ar_output_stock = [];
  }
}

export {Logger, LoggerStaticConfig, LevelListAll};
