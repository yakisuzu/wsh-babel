// ---------------
// private
// ---------------

const msg = (()=>{
  let m = {};
  m.not_import = '{0} has not been imported into the {1} module';
  m.not_support = '{0} class not support';
  m.dump_object = 'key : {0}, class : {1}';
  m.dump_array = 'index : {0}, class : {1}';
  m.dump_value = 'value : {0}, class : {1}';
  m.dump_error = 'name : {0}, message : {1}';
  return m;
})();

// ---------------
// public
// ---------------

/**
 *
 */
class Utility{

  /**
   * @param {Object} o
   */
  static echo(o){
    WScript.Echo(o);
  }

  /**
   * @param {String} st_msg
   * @param {Array<String>} ar_args
   */
  static buildMsg(st_msg, ar_args){
    let st_build = st_msg;
    ar_args.forEach((st_arg, i, self)=>{
      st_build = st_build.replace('{' + i + '}', st_arg);
    });
    return st_build;
  }

  /**
   * @param {Object} o
   * @return {String}
   */
  static getClass(o){
    let st_class =  Object.prototype.toString.apply(o);
    return st_class.replace(/\[object /, '').replace(/\]/, '');
  }

  /**
   * @param {Object} object
   */
  static dump(object){
    (function dumpR(object, st_pac_base){
      let st_class = getClass(object);
      let st_pac = (st_pac_base ? st_pac_base + '.' : '');

      switch(st_class){
        case 'Object':
          for(let key in object){
            let value = '';
            try{
              value = object[key];
            }catch(e){}
            echo(buildMsg(msg.dump_object, [st_pac + key, getClass(value)]));
            dumpR(value, st_pac + key);
          }
          break;

        case 'Array':
          object.forEach((value, i, self)=>{
            echo(buildMsg(msg.dump_array, [st_pac + i, getClass(value)]));
            dumpR(value, st_pac + i);
          });
          break;

        case 'Function':
          echo(object.toString());
          dumpR(object.prototype, st_pac + 'prototype');
          break;

        case 'Error':
          echo(buildMsg(msg.dump_error, [object.name, object.message]));
          break;

        case 'Boolean':
        case 'Number':
        case 'Date':
        case 'Math':
        case 'String':
        case 'RegExp':
          echo(buildMsg(msg.dump_value, [object.toString(), getClass(object)]));
          break;

        default:
          echo(buildMsg(msg.not_support, [st_class]));
      }
    })(object);
  }
}

export {Utility};
