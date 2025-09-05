/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const fs = require("fs");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");

const urlDev = "https://localhost:3001/";
const urlProd = "https://adraughon.github.io/SqExcel/"; // Updated to your GitHub Pages URL

/* global require, module, process, __dirname */

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

function getVersionInfo() {
  try {
    const versionData = JSON.parse(fs.readFileSync('./version.json', 'utf8'));
    return versionData;
  } catch (error) {
    console.warn('Could not read version.json, using default version');
    return { version: '1.0.0.0', description: 'Version info unavailable' };
  }
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
      functions: "./src/functions/functions.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.css$/,
          use: [
            MiniCssExtractPlugin.loader,
            "css-loader"
          ],
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
      new MiniCssExtractPlugin({
        filename: "styles.css" // Fixed filename instead of hash
      }),
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.ts",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "functions", "commands"],
        inject: true,
      }),
      new HtmlWebpackPlugin({
        filename: "index.html",
        template: "./src/index.html",
        chunks: [],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "version.json",
            to: "version.json",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              let transformedContent = content.toString();
              
              if (!dev) {
                // Replace URLs for production
                transformedContent = transformedContent.replace(new RegExp(urlDev.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), urlProd);
                
                // Update version from version.json
                const versionInfo = getVersionInfo();
                transformedContent = transformedContent.replace(
                  /<Version>[\d.]+<\/Version>/g,
                  `<Version>${versionInfo.version}</Version>`
                );
              }
              
              return transformedContent;
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization",
        "Access-Control-Allow-Credentials": "true",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3001,
    },
  };

  return config;
};
