// ---------------
// private
// ---------------

let ws_fso = WScript.CreateObject('Scripting.FileSystemObject');

/**
 * @return {Object}
 */
const msg = (()=>{
  let m  ={};
  return m;
})();

// ---------------
// public
// ---------------

/**
 *
 */
class File{

  /**
   * @param {String}
   * @return {boolean}
   */
  exists(st_path){
    return ws_fso.FileExists(st_path);
  }

  /**
   * @return {void}
   */
  createTextFile(st_file, ar_text){
    let ws_file = ws_fso.CreateTextFile(st_file, true);
    for(let st_text of ar_text){
      ws_file.WriteLine(st_text);
    }
    ws_file.Close();
  }
}

export {File};
