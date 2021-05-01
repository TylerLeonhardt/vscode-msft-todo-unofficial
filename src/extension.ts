import * as vscode from 'vscode';
import { MicrosoftToDoClientFactory } from './clientFactories/microsoftToDoClientFactory';
import { MicrosoftToDoTreeDataProvider } from './todoProviders/microsoftToDoTreeDataProvider';
import { TaskDetailsViewProvider } from './views/taskDetailsView';
import { TaskOperations } from './commands/taskOperations';
import { ListOperations } from './commands/listOperations';
import 'isomorphic-fetch';
import { MSAService } from './clientFactories/MSAService';
import { AADService } from './clientFactories/AADService';

export async function activate(context: vscode.ExtensionContext) {
	const msaService = new MSAService(context);
	const aadService = new AADService();
	await msaService.initialize();
	const clientProvider = new MicrosoftToDoClientFactory(msaService, aadService);
	const loginType: { type: 'msa' | 'aad' } | undefined = context.globalState.get('microsoftToDoUnofficialLoginType');
	if (loginType) {
		clientProvider.setLoginType(loginType.type);
	}

	context.subscriptions.push(vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.login',
		async () => {
			const result = await vscode.window.showQuickPick(['Microsoft account', 'Work or School account']);
			switch (result) {
				case 'Microsoft account':
					await msaService.createSession();
					await context.globalState.update('microsoftToDoUnofficialLoginType', { type: 'msa' });
					clientProvider.setLoginType('msa');
					break;
				case 'Work or School account': 
					await aadService.createSession();
					await context.globalState.update('microsoftToDoUnofficialLoginType', { type: 'aad' });
					clientProvider.setLoginType('aad');
					break;
			}

			vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
		}));

	context.subscriptions.push(vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.logout',
		async () => {
			await clientProvider.clearSessions();
			vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
		}));

	const treeDataProvider = new MicrosoftToDoTreeDataProvider(clientProvider);
	const view = vscode.window.createTreeView('microsoft-todo-unoffcial.listView', {
		treeDataProvider,
		showCollapseAll: true,
		canSelectMany: true
	});

	context.subscriptions.push(view);

	const taskDetailsProvider = new TaskDetailsViewProvider(context.extensionUri, clientProvider);
	view.onDidChangeSelection(async e => {
		if (e.selection.length > 0) {
			const node = e.selection[0];
			if (node.nodeType === 'task') {
				await taskDetailsProvider.changeChosenView(node);
			}
		}
	});

	const detailsView = vscode.window.registerWebviewViewProvider(
		taskDetailsProvider.viewType,
		taskDetailsProvider
	);
	context.subscriptions.push(detailsView);
	
	const taskOps = new TaskOperations(clientProvider);
	const listOps = new ListOperations(clientProvider);
	context.subscriptions.push(taskOps);
	context.subscriptions.push(listOps);
}

// this method is called when your extension is deactivated
export function deactivate() { }
