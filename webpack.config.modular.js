/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const CssMinimizerPlugin = require("css-minimizer-webpack-plugin");
const path = require("path");
const webpack = require("webpack");
const dotenv = require('dotenv');

const urlDev = "https://localhost:3002/";
const urlProd = "https://projectify5-0.vercel.app/";

/* global require, module, process, __dirname */

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  
  // Load environment variables
  const dotenvResult = dotenv.config().parsed;
  const envKeys = {};
  const requiredEnvVars = ['OPENAI_API_KEY', 'PINECONE_API_KEY', 'CLAUDE_API_KEY'];

  console.log("Webpack build mode:", options.mode);
  console.log("Using modular CSS configuration");

  requiredEnvVars.forEach(key => {
    const value = process.env[key] || (dotenvResult && dotenvResult[key]);
    if (value) {
      envKeys[`process.env.${key}`] = JSON.stringify(value);
      console.log(`✅ ${key}: Found`);
    } else {
      console.log(`❌ ${key}: NOT FOUND`);
      envKeys[`process.env.${key}`] = JSON.stringify(undefined);
    }
  });

  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      vendor: ["react", "react-dom", "core-js", "@fluentui/react"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
      functions: "./src/functions/functions.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].[contenthash].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
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
        // CSS Module Rules
        {
          test: /\.css$/,
          include: path.resolve(__dirname, "src/taskpane/styles"),
          use: [
            dev ? "style-loader" : MiniCssExtractPlugin.loader,
            {
              loader: "css-loader",
              options: {
                importLoaders: 1,
                sourceMap: true,
              },
            },
            {
              loader: "postcss-loader",
              options: {
                postcssOptions: {
                  plugins: [
                    "postcss-import", // Handles @import statements
                    "autoprefixer",   // Adds vendor prefixes
                    "postcss-nesting", // Allows nesting
                  ],
                },
              },
            },
          ],
        },
        // Legacy CSS (for backward compatibility during migration)
        {
          test: /\.css$/,
          exclude: path.resolve(__dirname, "src/taskpane/styles"),
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.js",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "vendor", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "vendor", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "vendor", "functions"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]." + (process.env.npm_lifecycle_event === "vercel-build" ? "production" : "dev") + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          {
            from: "src/auth/google/callback.html",
            to: "auth/google/callback.html",
          },
          {
            from: "src/prompts/**/*.txt",
            to: "prompts/[name][ext]",
          },
        ],
      }),
      new webpack.DefinePlugin(envKeys),
      // Extract CSS in production
      ...(!dev ? [new MiniCssExtractPlugin({
        filename: "styles/[name].[contenthash].css",
        chunkFilename: "styles/[id].[contenthash].css",
      })] : []),
    ],
    optimization: {
      minimize: !dev,
      minimizer: [
        "...", // Extends default minimizers
        new CssMinimizerPlugin(), // Minify CSS
      ],
      splitChunks: {
        chunks: "all",
        cacheGroups: {
          vendor: {
            test: /[\\/]node_modules[\\/]/,
            name: "vendor",
            priority: 10,
          },
          styles: {
            name: "styles",
            test: /\.css$/,
            chunks: "all",
            enforce: true,
          },
        },
      },
    },
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      compress: true,
      port: process.env.npm_lifecycle_event === "start" ? 3002 : 3002,
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/",
      },
      historyApiFallback: true,
      hot: true,
    },
  };

  return config;
};