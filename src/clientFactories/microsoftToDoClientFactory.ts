import { Client } from '@microsoft/microsoft-graph-client';
import * as vscode from 'vscode';

const redirectUri = encodeURIComponent(`${vscode.env.uriScheme}://tylerleonhardt.msft-todo-unofficial/`);
const scopes = ['Tasks.ReadWrite'];
const uri = vscode.Uri.parse(`https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=a4fd7674-4ebd-4dbc-831c-338314dd459e&response_type=token&redirect_uri=${redirectUri}&response_mode=fragment&scope=${scopes}`);

export class MicrosoftToDoClientFactory implements vscode.UriHandler {
	private _token: string | undefined;

	private didAuthenticate: vscode.EventEmitter<void> = new vscode.EventEmitter();
	public onDidAuthenticate: vscode.Event<void> = this.didAuthenticate.event;

	private loginCommand = vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.login',
		() => vscode.env.openExternal(uri));

	handleUri(uri: vscode.Uri) {
		const fragmentResult = new Map<string, string>();
		uri.fragment.split('&').forEach(q => {
			const [key, value] = q.split('=');
			fragmentResult.set(key, value);
		});

		if (fragmentResult.has('access_token')) {
			this._token = fragmentResult.get('access_token')!;
			this.didAuthenticate.fire();
			return;
		}

		console.log(uri.query);
	}

	public async getClient(): Promise<Client | undefined> {
		// const session = await vscode.authentication.getSession('microsoft', scopes, { createIfNone: true });

		if (!this._token) {
			// await vscode.env.openExternal(uri);
			return;
		}

		return Client.init({
			authProvider: (done) => {
				// done(undefined, session.accessToken);
				done(undefined, this._token!);
			}
		});
	}

	public async getAll<T>(client: Client, apiPath: string): Promise<T[]> {
		let iterUri: string | null | undefined = apiPath;
		const list = new Array<T>();
		while (iterUri) {
			let res = await client.api(iterUri).get() as { '@odata.nextLink': string | null | undefined; value: T[] };
			res.value.forEach(r => list.push(r));
			iterUri = res['@odata.nextLink'];
		}

		return list;
	}
}
