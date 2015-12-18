import {Utility} from './Utility.js';

// ---------------
// private
// ---------------

const ws_fso = WScript.CreateObject('Scripting.FileSystemObject');

const msg = (()=>{
  const m = {};
  m.not_exists = 'not exists {0}';
  m.yyyymmddHHMM = '{0}/{1}/{2} {3}:{4}';
  return m;
})();

/**
 *
 */
class FileInfo{
  constructor(st_path, st_date_last_modified){
    this.path = st_path;
    this.date_last_modified = st_date_last_modified;
  }
}

/**
 *
 */
const getAppended0 = (st_base)=>{
  return ('0' + st_base).slice(-2);
};

/**
 *
 */
const getFormatedDate = (dt)=>{
  return Utility.buildMsg(msg.yyyymmddHHMM, [
      dt.getFullYear()
      , getAppended0(dt.getMonth() + 1)
      , getAppended0(dt.getDate())
      , getAppended0(dt.getHours())
      , getAppended0(dt.getMinutes())
  ]);
};

/**
 * @param {Object} ws_itr
 * @return {Object}
 */
function* eachEnumerator(ws_itr){
  const enu = new Enumerator(ws_itr);
  for(enu.moveFirst(); !enu.atEnd(); enu.moveNext()){
    yield enu.item();
  }
}

// ---------------
// public
// ---------------

/**
 *
 */
class Config{
  constructor(){
    this.ignore_dir_reg = [];
  }
}

/**
 *
 */
class FileSystem{

  /**
   * @constructor
   */
  constructor(log){
    this.logger = log;

    this.config = new Config();
  }

  /**
   * @param {String} st_filepath
   * @return {boolean}
   */
  static fileExists(st_filepath){
    return ws_fso.FileExists(st_filepath);
  }

  /**
   * @param {String} st_folderpath
   * @return {boolean}
   */
  static folderExists(st_folderpath){
    return ws_fso.FolderExists(st_folderpath);
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

  /**
   * @param {String} st_base_dir
   * @return {Array<FileInfo>}
   */
  getFiles(st_base_dir){
    if(!FileSystem.folderExists(st_base_dir)){
      this.logger.error(msg.not_exists, [st_base_dir]);
      return [];
    }

    const get_files = (ws_folder)=>{
      const ar_files = [];
      for(let ws_file of eachEnumerator(ws_folder.Files)){
        const st_path = ws_file.Path;
        const st_date_last_modified = getFormatedDate(new Date(ws_file.DateLastModified));

        ar_files.push(new FileInfo(st_path, st_date_last_modified));
      }
      return ar_files;
    };

    const get_folders = (ws_folder)=>{
      let ar_files = [];
      for(let ws_folder of eachEnumerator(ws_folder.SubFolders)){
        // when found regex, skip folder
        if(this.config.ignore_dir_reg.some(
              (reg)=>{return ws_folder.Path.search(reg) !== -1;}
              )){
          continue;
        }
        ar_files = ar_files.concat(get_folders(ws_folder));
      }
      return ar_files.concat(get_files(ws_folder));
    };

    return get_folders(ws_fso.GetFolder(st_base_dir));
  }
}

export {FileSystem};
