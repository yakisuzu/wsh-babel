module.exports = {
  entry: {
    index: "./src/index.js",
    DeleteErrorExcelName: "./src/DeleteErrorExcelName.js",
    DeleteErrorExcelFormat: "./src/DeleteErrorExcelFormat.js"
  },
  output: {
    path: "./dest/",
    filename: "[name].js"
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
