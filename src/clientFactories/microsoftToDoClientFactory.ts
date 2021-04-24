import { Client } from '@microsoft/microsoft-graph-client';
import * as vscode from 'vscode';
import { AzureActiveDirectoryService } from '../AADHelper';

// const redirectUri = encodeURIComponent(`${vscode.env.uriScheme}://tylerleonhardt.msft-todo-unofficial/`);
// const scopes = ['Tasks.ReadWrite'];
// const uri = vscode.Uri.parse(`https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=a4fd7674-4ebd-4dbc-831c-338314dd459e&response_type=token&redirect_uri=${redirectUri}&response_mode=fragment&scope=${scopes}`);

export class MicrosoftToDoClientFactory {
	private _token: string | undefined;

	private didAuthenticate: vscode.EventEmitter<void> = new vscode.EventEmitter();
	public onDidAuthenticate: vscode.Event<void> = this.didAuthenticate.event;

	constructor(private aadService: AzureActiveDirectoryService) {}

	private loginCommand = vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.login',
		() => this.aadService.createSession());

	// handleUri(uri: vscode.Uri) {
	// 	const fragmentResult = new Map<string, string>();
	// 	uri.fragment.split('&').forEach(q => {
	// 		const [key, value] = q.split('=');
	// 		fragmentResult.set(key, value);
	// 	});

	// 	if (fragmentResult.has('access_token')) {
	// 		this._token = fragmentResult.get('access_token')!;
	// 		this.didAuthenticate.fire();
	// 		return;
	// 	}

	// 	console.log(uri.query);
	// }

	public async getClient(): Promise<Client | undefined> {
		const sessions = (await this.aadService.sessions)
		if (!sessions?.length) {
			return;
		}
		// const session = sessions[0]; // await vscode.authentication.getSession('microsoft', scopes, { createIfNone: true });

		// if (!this._token) {
		// 	// await vscode.env.openExternal(uri);
		// 	return;
		// }

		return Client.init({
			authProvider: (done) => {
				done(undefined, sessions[0].accessToken);
				// done(undefined, this._token!);
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
