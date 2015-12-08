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

/**
 *
 */
class Config{
  constructor(){
    this.output_level = LevelListAll.INFO;
    this.header = (st_level)=>{return '[' + st_level + ']';};
    this.linefeed = '\n';
    this.output = (st_msg)=>{Utility.echo(st_msg);};
  }
}

/**
 * TODO setting format
 * TODO show line no
 */
class Logger{

  /**
   * @constructor
   */
  constructor(){
    this.config = new Config();
    this.output_stock = [];
  }

  /**
   * @param {String}
   * @param {Array<String>}
   */
  trace(st_msg, ar_args=[]){outputPush('TRACE', Utility.buildMsg(st_msg, ar_args));}
  debug(st_msg, ar_args=[]){outputPush('DEBUG', Utility.buildMsg(st_msg, ar_args));}
  info (st_msg, ar_args=[]){outputPush('INFO',  Utility.buildMsg(st_msg, ar_args));}
  warn (st_msg, ar_args=[]){outputPush('WARN',  Utility.buildMsg(st_msg, ar_args));}
  error(st_msg, ar_args=[]){outputPush('ERROR', Utility.buildMsg(st_msg, ar_args));}
  fatal(st_msg, ar_args=[]){outputPush('FATAL', Utility.buildMsg(st_msg, ar_args));}

  /**
   * @param {String} st_level
   * @param {String} st_text
   */
  outputPush(st_level, st_text){
    if(this.config.output_level <= LevelList[st_level]){
      this.output_stock.push(new OutputText(st_level, st_text));
    }
  }

  /**
   *
   */
  print(){
    this.config.output(
      this.output_stock.map((outputText)=>{
        return this.config.header(outputText.level) + outputText.text
      }).join(this.config.linefeed)
    );
    this.output_stock = [];
  }
}

export {Logger, LevelListAll};
