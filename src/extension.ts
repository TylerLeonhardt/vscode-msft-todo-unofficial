import * as vscode from 'vscode';
import { MicrosoftToDoTreeDataProvider } from './todoProviders/microsoftToDoTreeDataProvider';
import { TaskCreateView } from './views/taskCreateView';
import { TaskDetailsViewProvider } from './views/taskDetailsView';

export async function activate(context: vscode.ExtensionContext) {

	const treeDataProvider = new MicrosoftToDoTreeDataProvider();
	const view = vscode.window.createTreeView('microsoft-todo', {
		treeDataProvider,
		showCollapseAll: true,
		canSelectMany: true
	});

	context.subscriptions.push(view);

	const taskDetailsProvider = new TaskDetailsViewProvider(context.extensionUri);
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

	const taskCreateProvider = new TaskCreateView(context.extensionUri);

	context.subscriptions.push(vscode.window.registerWebviewViewProvider(
		taskCreateProvider.viewType,
		taskCreateProvider
	));

	vscode.commands.registerCommand('microsoft-todo.createTask', async () => {
		await vscode.commands.executeCommand('setContext', 'showTaskCreateView', true);
		taskCreateProvider.show();
	});

	vscode.commands.registerCommand('microsoft-todo.cancelCreateTask', async () => {
		await vscode.commands.executeCommand('setContext', 'showTaskCreateView', false);
	});
}

// this method is called when your extension is deactivated
export function deactivate() {}
