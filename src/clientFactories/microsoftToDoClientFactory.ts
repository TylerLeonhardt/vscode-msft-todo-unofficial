import { Client } from '@microsoft/microsoft-graph-client';
import * as vscode from 'vscode';
import { AADService } from './AADService';
import { MSAService } from './MSAService';

export class MicrosoftToDoClientFactory {
	private loginType: 'msa' | 'aad' | undefined;
	constructor(private msaService: MSAService, private aadService: AADService) {}

	public async getClient(): Promise<Client | undefined> {
		const session = this.loginType === 'msa'
			? (await this.msaService.getSessions())[0]
			: (await this.aadService.getSessions())[0];
		if (!session) {
			return;
		}

		return Client.init({
			authProvider: (done) => {
				done(undefined, session.accessToken);
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

	public async clearSessions() {
		await this.msaService.clearSessions();
		await this.aadService.clearSessions();
	}

	public setLoginType(type: 'msa' | 'aad') {
		this.loginType = type;
	}
}
