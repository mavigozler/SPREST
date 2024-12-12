"use strict";

import path from 'path';
import { fileURLToPath } from 'url';
import HtmlWebpackPlugin from 'html-webpack-plugin';
import { Configuration } from 'webpack';

const __filename = fileURLToPath(import.meta.url); // get the resolved path to the file
const __dirname = path.dirname(__filename); // get the name of the directory

export const webpackConfig: Configuration = {
	entry: {
		app: "../../src/JSONtool.ts",
		vendor: [ "json5" ]
	},
	module: {
		rules: [
			{
				test: /\.ts?$/,
				use: "ts-loader",
				exclude: /node_modules/
			},
			{
				test: /\.css$/, // match *.css files
				use: [ "style-loader", "css-loader" ]
			},
			{
				test: /\.(jpg|png)$/, // match jpg and png
				use: {
					loader: "file-loader",
					options: { name: "[name].[ext]", outputPath: "img" }
						// output directory for images
				}
			}
		]
	},
	resolve: { extensions: [ ".tsx", ".ts", ".js" ]},
	output: {
		filename: '[name].bundle.js',
		path: path.resolve(__dirname, 'js'),
	},
	mode: 'development',
	plugins: [
		new HtmlWebpackPlugin({
			template: '../../html/JSONtool.html',
		}),
	]
};
