/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");
const dotenv = require('dotenv');

// Load environment variables from .env file (local development) or process.env (production)
const env = dotenv.config().parsed;

// Create an object with the env variables, prioritizing process.env over .env file
const envKeys = {};
const requiredEnvVars = ['OPENAI_API_KEY', 'PINECONE_API_KEY', 'CLAUDE_API_KEY'];

requiredEnvVars.forEach(key => {
  // Use process.env first (for Vercel), then fall back to .env file
  const value = process.env[key] || (env && env[key]);
  if (value) {
    envKeys[`process.env.${key}`] = JSON.stringify(value);
  }
});

console.log("Environment variables loaded:", Object.keys(envKeys).map(k => k.replace('process.env.', '')));

const urlDev = "https://localhost:3002/";
const urlProd = "https://projectify5-0.vercel.app/"; // Vercel production deployment URL

/* global require, module, process, __dirname */

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
      functions: "./src/functions/functions.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
      fallback: {
        // Provide node module polyfills
        "process": require.resolve("process/browser"),
        "os": require.resolve("os-browserify"),
        "path": false, // No polyfill needed
        "fs": false // No polyfill needed
      }
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [

      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.js",
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "index.html",
        template: "./src/index.html",
        chunks: [],
      }),
      new webpack.DefinePlugin({
        'process.env.NODE_ENV': JSON.stringify(dev ? 'development' : 'production'),
        ...envKeys,
      }),
      new webpack.ProvidePlugin({
        process: 'process/browser',
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "./src/prompts/*",
            to: "prompts/[name][ext]"
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content
                  .toString()
                  .replace(new RegExp(urlProd, "g"), urlDev);
              } else {
                return content
                  .toString()
                  .replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },

        ],
      }),
    ],
    devServer: {
      static: [
        {
          directory: path.join(__dirname, "dist"),
          publicPath: "/"
        },
        {
          directory: path.join(__dirname, "assets"),
          publicPath: "/assets"
        }
      ],
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3002,
    },
  };

  return config;
};
