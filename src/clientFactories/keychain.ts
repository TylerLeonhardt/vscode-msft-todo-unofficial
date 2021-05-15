/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as vscode from 'vscode';

const SERVICE_ID = `microsoft-todo-unofficial.login`;

export class Keychain {
	constructor(private context: vscode.ExtensionContext) {}


	async setToken(token: string): Promise<void> {

		try {
			return await this.context.secrets.store(SERVICE_ID, token);
		} catch (e) {
			console.error(`Setting token failed: ${e}`);

			// Temporary fix for #94005
			// This happens when processes write simulatenously to the keychain, most
			// likely when trying to refresh the token. Ignore the error since additional
			// writes after the first one do not matter. Should actually be fixed upstream.
			if (e.message === 'The specified item already exists in the keychain.') {
				return;
			}

			const troubleshooting = "Troubleshooting Guide";
			const result = await vscode.window.showErrorMessage(`Writing login information to the keychain failed with error '${e.message}'.`, troubleshooting);
			if (result === troubleshooting) {
				vscode.env.openExternal(vscode.Uri.parse('https://code.visualstudio.com/docs/editor/settings-sync#_troubleshooting-keychain-issues'));
			}
		}
	}

	async getToken(): Promise<string | null | undefined> {
		try {
			return await this.context.secrets.get(SERVICE_ID);
		} catch (e) {
			// Ignore
			console.error(`Getting token failed: ${e}`);
			return Promise.resolve(undefined);
		}
	}

	async deleteToken(): Promise<void> {
		try {
			return await this.context.secrets.delete(SERVICE_ID);
		} catch (e) {
			// Ignore
			console.error(`Deleting token failed: ${e}`);
			return Promise.resolve(undefined);
		}
	}
}
