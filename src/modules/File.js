// ---------------
// private
// ---------------

const ws_fso = WScript.CreateObject('Scripting.FileSystemObject');

const msg = (()=>{
  const m = {};
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
   * @param {String} st_path
   * @return {boolean}
   */
  static exists(st_path){
    return ws_fso.FileExists(st_path);
  }

  /**
   * @param {String} st_file
   * @param {Array<String>} ar_text
   */
  static createTextFile(st_file, ar_text){
    const ws_file = ws_fso.CreateTextFile(st_file, true);
    for(let st_text of ar_text){
      ws_file.WriteLine(st_text);
    }
    ws_file.Close();
  }
}

export {File};
