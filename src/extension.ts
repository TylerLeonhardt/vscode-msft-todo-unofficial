import * as vscode from 'vscode';
import { MicrosoftToDoClientFactory } from './clientFactories/microsoftToDoClientFactory';
import { MicrosoftToDoTreeDataProvider } from './todoProviders/microsoftToDoTreeDataProvider';
import { TaskDetailsViewProvider } from './views/taskDetailsView';
import { TaskOperations } from './commands/taskOperations';
import { ListOperations } from './commands/listOperations';
import 'cross-fetch/polyfill';

export async function activate(context: vscode.ExtensionContext) {
	console.log('Activating!');

	const clientProvider = new MicrosoftToDoClientFactory(context.globalState);
	const loginType: { type?: 'msa' | 'microsoft' } | undefined = context.globalState.get('microsoftToDoUnofficialLoginType');
	if (loginType) {
		clientProvider.setLoginType(loginType.type);
	}

	let disposable: vscode.Disposable;
	context.subscriptions.push(vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.login',
		async () => {
			const result = await vscode.window.showQuickPick(['Microsoft account', 'Work or School account']);

			if (!result) {
				return;
			}

			const provider = result === 'Microsoft account' ? 'msa' : 'microsoft';
			await vscode.authentication.getSession('microsoft', result === 'Microsoft account' ? MicrosoftToDoClientFactory.msaScopes : MicrosoftToDoClientFactory.scopes, { createIfNone: true });
			await context.globalState.update('microsoftToDoUnofficialLoginType', { type: provider });
			clientProvider.setLoginType(provider);
			vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
		}));

	context.subscriptions.push(disposable = vscode.authentication.onDidChangeSessions((e) => clientProvider.clearLoginTypeState(e)));

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

	context.subscriptions.push(vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.showTaskDetailsView',
		async () => {
			await vscode.commands.executeCommand('setContext', 'showTaskDetailsView', true);
		}));

	context.subscriptions.push(vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.hideTaskDetailsView',
		async () => {
			await vscode.commands.executeCommand('setContext', 'showTaskDetailsView', false);
		}));
	
	const taskOps = new TaskOperations(clientProvider);
	const listOps = new ListOperations(clientProvider);
	context.subscriptions.push(taskOps);
	context.subscriptions.push(listOps);
}

// this method is called when your extension is deactivated
export function deactivate() { }
