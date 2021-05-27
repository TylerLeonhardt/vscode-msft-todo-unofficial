/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

//@ts-check
'use strict';

//@ts-check
/** @typedef {import('webpack').Configuration} WebpackConfig **/

const path = require('path');
const webpack = require('webpack');

module.exports = /** @type WebpackConfig */ {
	context: __dirname,
	mode: 'none', // this leaves the source code as close as possible to the original (when packaging we set this to 'production')
	target: 'webworker', // extensions run in a webworker context
	entry: {
		'extension': './src/extension.ts',
		// 'test/suite/index': './src/web/test/suite/index.ts'
	},
	resolve: {
		mainFields: ['module', 'main'],
		extensions: ['.ts', '.js'], // support ts-files and js-files
		alias: {
			'./env/node': path.resolve(__dirname, 'src/clientFactories/env/browser'),
			'./authServer': path.resolve(__dirname, 'src/clientFactories/env/browser/authServer'),
			'node-fetch': path.resolve(__dirname, 'node_modules/node-fetch/browser.js'),
			'buffer': path.resolve(__dirname, 'node_modules/buffer/index.js'),
			'randombytes': path.resolve(__dirname, 'node_modules/randombytes/browser.js'),
			'stream': path.resolve(__dirname, 'node_modules/stream/index.js'),
			'uuid': path.resolve(__dirname, 'node_modules/uuid/dist/esm-browser/index.js')
		},
		fallback: {
			'assert': require.resolve('assert')
		}
	},
	module: {
		rules: [{
			test: /\.ts$/,
			exclude: /node_modules/,
			use: [
				{
					loader: 'ts-loader'
				}
			]
		}]
	},
	plugins: [
		new webpack.ProvidePlugin({
			process: 'process/browser',
		}),
	],
	externals: {
		'vscode': 'commonjs vscode', // ignored because it doesn't exist
	},
	performance: {
		hints: false
	},
	output: {
		filename: '[name].js',
		path: path.join(__dirname, './dist/web'),
		libraryTarget: 'commonjs'
	},
	devtool: 'nosources-source-map'
};