// ---------------
// private
// ---------------

/**
 * @return {Object}
 */
const msg = (()=>{
  let m  ={};
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

  constructor(logger){
    this.Logger = logger
  }

  /**
   * @return {Array<String>}
   */
  getArgs(){
    let ws_args = WScript.Arguments;
    if(ws_args.Length === 0){
      this.Logger.info(msg.no_args);
      this.Logger.print();
      WScript.Quit();
    }

    let ar_args = [];
    for(let nu_arg = 0; nu_arg < ws_args.Length; nu_arg++){
      let st_arg = ws_args.Item(nu_arg);
      this.Logger.trace(st_arg);
      ar_args.push(st_arg);
    }
    return ar_args;
  }
}

export {Args};
