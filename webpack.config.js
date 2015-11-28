module.exports = {
  entry: {
    index: "./src/index.js"
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
