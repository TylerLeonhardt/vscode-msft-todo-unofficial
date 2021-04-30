import { Client } from '@microsoft/microsoft-graph-client';
import * as vscode from 'vscode';
import { AzureActiveDirectoryService } from '../AADHelper';

export class MicrosoftToDoClientFactory {
	constructor(private aadService: AzureActiveDirectoryService) {}

	public async getClient(): Promise<Client | undefined> {
		const sessions = await this.aadService.sessions
		if (!sessions?.length) {
			return;
		}

		return Client.init({
			authProvider: (done) => {
				done(undefined, sessions[0].accessToken);
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
