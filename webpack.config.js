const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const httpsOptions = dev ? await getHttpsOptions() : undefined;

  return {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "ts-loader",
        },
        {
          test: /\.js$/,
          enforce: "pre",
          use: ["source-map-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*.png",
            to: "assets/[name][ext]",
            globOptions: { ignore: ["**/emblem-master.png"] },
          },
          { from: "manifest.xml", to: "[name][ext]" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
        ],
      }),
    ],
    devServer: dev
      ? {
          headers: { "Access-Control-Allow-Origin": "*" },
          server: { type: "https", options: httpsOptions },
          port: 3000,
        }
      : undefined,
  };
};
