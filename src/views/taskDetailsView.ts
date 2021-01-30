import * as vscode from 'vscode';
import { TaskNode } from '../todoProviders/microsoftToDoTreeDataProvider';
import { getNonce, WebviewViewBase } from './WebViewBase';

export class TaskDetailsViewProvider extends WebviewViewBase implements vscode.WebviewViewProvider {
    public readonly viewType = 'microsoft-todo.taskDetailsView';

    private chosenTask: TaskNode | undefined;

    constructor(private readonly _extensionUri: vscode.Uri) {
		super();
	}

    public async changeChosenView(node: TaskNode) {
		this.chosenTask = node;
		vscode.commands.executeCommand('setContext', 'showTaskDetailsView', true);
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

        // if (this.chosenTask) {
        //     await this.changeChosenView(this.chosenTask);
        // }
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

		// Use a nonce to only allow a specific script to be run.
		const nonce = getNonce();

		let content: string;
		if (!this.chosenTask) {
			content = `
			<h2 class="task-title"></h2>
			<p class="task-body">Select a task to see additional details.</p>`;
		} else {
			content = `
			<h2 class="task-title">${this.chosenTask?.entity.title || ""}</h2>
			<p class="task-body">${this.chosenTask?.entity.body?.content || ""}</p>`;
		}

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
				<link href="${styleMainUri}" rel="stylesheet">
				<title>Task details</title>
			</head>
			<body>
				${content}
				<script nonce="${nonce}" src="${scriptUri}"></script>
			</body>
			</html>`;
	}
}
