import * as vscode from 'vscode';
import { MicrosoftToDoClientFactory } from './clientFactories/microsoftToDoClientFactory';
import { MicrosoftToDoTreeDataProvider } from './todoProviders/microsoftToDoTreeDataProvider';
import { TaskDetailsViewProvider } from './views/taskDetailsView';
import { TaskOperations } from './commands/taskOperations';
import { ListOperations } from './commands/listOperations';
import 'isomorphic-fetch';
import { AzureActiveDirectoryService } from './AADHelper';

export async function activate(context: vscode.ExtensionContext) {
	const aadService = new AzureActiveDirectoryService(context);
	await aadService.initialize();

	
	context.subscriptions.push(vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.login',
		async () => {
			await aadService.createSession();
			vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
		}));

	context.subscriptions.push(vscode.commands.registerCommand(
		'microsoft-todo-unoffcial.logout',
		async () => {
			await aadService.clearSessions();
			vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
		}));

	const clientProvider = new MicrosoftToDoClientFactory(aadService);
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
