var fs = require('fs');
var path = require('path');

var ob_entry = (function(){
  var ob_ret = {};
  var st_base_dir = './src/';
  fs.readdirSync(st_base_dir).filter(function(st_file){
    return st_file.search(/\.js$/) !== -1
  }).forEach(function(st_file){
    ob_ret[st_file] = path.resolve(path.join(st_base_dir, st_file));
  });
  return ob_ret;
})();

module.exports = {
  entry: ob_entry,
  output: {
    filename: '[name]'
  },
  module: {
    loaders: [
    {
      test: /\.jsx?$/,
      exclude: /(node_modules|bower_components)/,
      loader: 'babel'
    }
    ]
  }
}
