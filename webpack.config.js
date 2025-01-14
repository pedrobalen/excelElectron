const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = {
  entry: "./src/index.js",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "bundle.js",
  },
  module: {
    rules: [
      {
        test: /\.jsx?$/, // Babel loader for JavaScript and JSX
        exclude: /node_modules/,
        use: {
          loader: "babel-loader",
          options: {
            presets: ["@babel/preset-env", "@babel/preset-react"],
          },
        },
      },
      {
        test: /\.css$/, // CSS loader
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  resolve: {
    extensions: [".js", ".jsx", ".json"],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./public/index.html", // HTML template
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"), // Directory to serve static files from
    },
    port: 3000, // Port for the dev server
    open: true, // Automatically open the browser
  },
};
