const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const VueLoaderPlugin = require('vue-loader/lib/plugin')
const fs = require("fs");
const webpack = require("webpack");

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js"
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js', '.vue'],
      alias: {
        vue$: 'vue/dist/vue.js'
      }
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader", 
            options: {
              presets: ["@babel/preset-env"]
            }
          }
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.vue$/,
          loader: 'vue-loader',
          options: { esModule: true }
        },
        {
          test: /\.stylus$/,
          use: ['vue-style-loader', 'css-loader', 'stylus-loader']
        },
        {
          test: /\.css$/,
          use: ['vue-style-loader', 'css-loader']
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          use: "file-loader"
        },
        {
          test: /\.(woff2?|eot|ttf|otf)(\?.*)?$/,
          use: 'file-loader'
        }
      ]
    },
    plugins: [
      new VueLoaderPlugin(),
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new CopyWebpackPlugin([
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css"
        }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      disableHostCheck: true,
      // https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};
