import * as vscode from 'vscode';
import { MicrosoftToDoClientFactory } from './clientFactories/microsoftToDoClientFactory';
import { MicrosoftToDoTreeDataProvider } from './todoProviders/microsoftToDoTreeDataProvider';
import { TaskDetailsViewProvider } from './views/taskDetailsView';
import { TaskOperations } from './commands/TaskOperations';
import { ListOperations } from './commands/listOperations';
import 'isomorphic-fetch';

export async function activate(context: vscode.ExtensionContext) {
	const clientProvider = new MicrosoftToDoClientFactory();
	context.subscriptions.push(vscode.window.registerUriHandler(clientProvider));

	const treeDataProvider = new MicrosoftToDoTreeDataProvider(clientProvider);
	clientProvider.onDidAuthenticate(() => vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList'));
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
