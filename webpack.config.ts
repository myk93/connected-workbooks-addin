import * as path from "path";
import * as webpack from "webpack";
import HtmlWebpackPlugin from "html-webpack-plugin";
import CopyWebpackPlugin from "copy-webpack-plugin";
import { Configuration as WebpackDevServerConfiguration } from "webpack-dev-server";

interface Configuration extends webpack.Configuration {
    devServer?: WebpackDevServerConfiguration;
}

const config: Configuration = {
    entry: {
        taskpane: "./src/taskpane/index.tsx",
    },
    output: {
        path: path.resolve(__dirname, "dist"),
        filename: "[name].bundle.js",
        clean: true,
    },
    resolve: {
        extensions: [".ts", ".tsx", ".js", ".jsx"],
        fallback: {
            // Required because connected-workbooks uses Buffer.from()
            buffer: require.resolve("buffer/"),
        },
        // Prevent webpack from following file: symlinks to the real path.
        // Without this, modules in the linked package resolve against the
        // linked directory's own node_modules instead of ours.
        symlinks: false,
    },
    module: {
        rules: [
            {
                test: /\.[tj]sx?$/,
                exclude: /node_modules/,
                use: "ts-loader",
            },
            {
                test: /\.css$/,
                use: ["style-loader", "css-loader"],
            },
        ],
    },
    plugins: [
        new webpack.ProvidePlugin({
            // Make Buffer available globally for connected-workbooks
            Buffer: ["buffer", "Buffer"],
        }),
        new HtmlWebpackPlugin({
            template: "./src/taskpane/index.html",
            filename: "taskpane.html",
            chunks: ["taskpane"],
        }),
        new CopyWebpackPlugin({
            patterns: [
                {
                    from: "assets",
                    to: "assets",
                },
                {
                    from: "manifest.xml",
                    to: "manifest.xml",
                },
            ],
        }),
    ],
    externals: {
        // Prevent office-js from being bundled — it's loaded from CDN in index.html
        "office-js": "Office",
    },
    devServer: {
        hot: true,
        port: 3000,
        server: {
            type: "https",
            options: {
                // Run `npx office-addin-dev-certs install` to generate these
                ca: path.resolve(
                    process.env.USERPROFILE || process.env.HOME || "",
                    ".office-addin-dev-certs/ca.crt"
                ),
                key: path.resolve(
                    process.env.USERPROFILE || process.env.HOME || "",
                    ".office-addin-dev-certs/localhost.key"
                ),
                cert: path.resolve(
                    process.env.USERPROFILE || process.env.HOME || "",
                    ".office-addin-dev-certs/localhost.crt"
                ),
            },
        },
        headers: {
            "Access-Control-Allow-Origin": "*",
        },
    },
    mode: "development",
    devtool: "source-map",
};

export default config;
