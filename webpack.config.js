const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");
require("dotenv").config();

// Webpack supports exporting an async function (or a Promise) as config.
module.exports = async (env, argv) => {
  if (argv.mode !== "production") {
    // Silence Node deprecation warnings from dev-server dependencies (e.g., util._extend).
    process.noDeprecation = true;
  }
  const httpsOptions = await devCerts.getHttpsServerOptions();

  const devServerPort = Number(process.env.WEB_PORT) || 3000;

  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.tsx",
      commands: "./src/commands/commands.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "manifest", to: "manifest" },
          { from: "assets", to: "assets" },
        ],
      }),
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
      },
      port: devServerPort,
      server: {
        type: "https",
        options: httpsOptions,
      },
      proxy: [
        {
          context: ["/api"],
          target: "http://localhost:3001",
          secure: false,
          changeOrigin: true,
        },
      ],
      hot: false,
      liveReload: true,
    },
    devtool: "source-map",
  };
};
