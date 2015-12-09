import {Utility} from './Utility.js';

// ---------------
// private
// ---------------

const msg = (()=>{
  const m = {};
  m.no_args = 'Please drag & drop any file!';
  return m;
})();

// ---------------
// public
// ---------------

/**
 *
 */
class Args{

  /**
   * @return {Array<String>}
   */
  static getArgs(bo_no_args_error=true){
    const ws_args = WScript.Arguments;
    if(bo_no_args_error){
      if(ws_args.Length === 0){
        Utility.echo(msg.no_args);
        WScript.Quit();
      }
    }

    const ar_args = [];
    for(let nu_arg = 0; nu_arg < ws_args.Length; nu_arg++){
      const st_arg = ws_args.Item(nu_arg);
      ar_args.push(st_arg);
    }
    return ar_args;
  }
}

export {Args};
