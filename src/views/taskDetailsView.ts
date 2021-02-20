import { TodoTask } from '@microsoft/microsoft-graph-types';
import * as vscode from 'vscode';
import { MicrosoftToDoClientFactory } from '../clientFactories/microsoftToDoClientFactory';
import { TaskNode } from '../todoProviders/microsoftToDoTreeDataProvider';
import { getNonce, WebviewViewBase } from './WebViewBase';

export class TaskDetailsViewProvider extends WebviewViewBase implements vscode.WebviewViewProvider {
	public readonly viewType = 'microsoft-todo-unoffcial.taskDetailsView';

	private chosenTask: TaskNode | undefined;

	constructor(private readonly _extensionUri: vscode.Uri, private readonly clientFactory: MicrosoftToDoClientFactory) {
		super();
	}

	public async changeChosenView(node: TaskNode) {
		this.chosenTask = node;
		await vscode.commands.executeCommand('setContext', 'showTaskDetailsView', true);
		this.show(true);
		await this._postMessage(node);
	}

	async resolveWebviewView(webviewView: vscode.WebviewView, context: vscode.WebviewViewResolveContext<unknown>, token: vscode.CancellationToken): Promise<void> {
		this._view = webviewView;
		this._webview = webviewView.webview;
		this._webview.options = {
			// Allow scripts in the webview
			enableScripts: true,

			localResourceRoots: [
				this._extensionUri
			],
		};
		super.initialize();

		this._webview.html = this.getHtmlForWebview();

		this._disposables.push(webviewView.webview.onDidReceiveMessage(async (message) => {
			switch (message.command) {
				case 'cancel':
					await vscode.commands.executeCommand('microsoft-todo-unoffcial.closeCreateTask');
					break;
				case 'update':
					const client = await this.clientFactory.getClient();
					if (!client) {
						return await vscode.window.showErrorMessage("you're not logged in.");
					}

					const body: TodoTask = {
						title: message.body.title,
						body: {
							content: message.body.note,
							contentType: 'text'
						}
					};

					if (message.body.dueDate) {
						const [ month, day, year ] = message.body.dueDate.split('/');

						body.dueDateTime = {
							dateTime: `${year}-${month}-${day}T08:00:00.0000000`,
							timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
						};
					}

					// TODO: error handling
					await client.api(`/me/todo/lists/${message.body.listId}/tasks/${message.body.id}`).patch(body);

					await vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
					break;
			}
		}));
	}

	private getHtmlForWebview() {
		if (!this._webview) {
			throw new Error('bad state: no webview found');
		}

		// Get the local path to main script run in the webview, then convert it to a uri we can use in the webview.
		const scriptUri = this._webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'webviews', 'taskDetailsView', 'main.js'));

		// Do the same for the stylesheet.
		const styleResetUri = this._webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'webviews', 'common', 'reset.css'));
		const styleVSCodeUri = this._webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'webviews', 'common', 'vscode.css'));
		const styleMainUri = this._webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'webviews', 'taskDetailsView', 'main.css'));
		const tdpCss = this._webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'node_modules', 'tiny-date-picker', 'tiny-date-picker.min.css'));
		const tdpScript = this._webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'node_modules', 'tiny-date-picker', 'dist', 'tiny-date-picker.min.js'));
		// Use a nonce to only allow a specific script to be run.
		const nonce = getNonce();

		return `<!DOCTYPE html>
			<html lang="en">
			<head>
				<meta charset="UTF-8">
				<!--
					Use a content security policy to only allow loading images from https or from our extension directory,
					and only allow scripts that have a specific nonce.
				-->
				<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${this._webview.cspSource}; script-src 'nonce-${nonce}';">
				<meta name="viewport" content="width=device-width, initial-scale=1.0">
				<link href="${styleResetUri}" rel="stylesheet">
				<link href="${styleVSCodeUri}" rel="stylesheet">
				<Link href="${tdpCss}" rel="stylesheet">
				<link href="${styleMainUri}" rel="stylesheet">
				<title>Task details</title>
			</head>
			<body>
				<input placeholder='Add Title' type='text' class='task-title' value=''/>
				<input placeholder='No due date' type='text' class='task-duedate-input' value=''/>
				<textarea placeholder='Add Note' class='task-body'></textarea>
				<button class='update update-task' hidden>Update</button>
				<button class='update update-cancel' hidden>Cancel</button>
				<h2 class='tooltip'>Additional details:
					<span class="tooltiptext">Edit these properties in the Microsoft To-Do app</span>
				</h2>
				<h4 class='task-reminder'></h4>
				<script nonce="${nonce}" src="${tdpScript}"></script>
				<script nonce="${nonce}" src="${scriptUri}"></script>
			</body>
			</html>`;
	}
}
