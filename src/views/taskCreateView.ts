import * as vscode from 'vscode';
import { TaskNode } from '../todoProviders/microsoftToDoTreeDataProvider';
import { getNonce, WebviewViewBase } from './WebViewBase';

export class TaskCreateView extends WebviewViewBase implements vscode.WebviewViewProvider {
    public readonly viewType = 'microsoft-todo.taskCreateView';

    constructor(private readonly _extensionUri: vscode.Uri) {
		super();
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

        webviewView.webview.html = this.getHtmlForWebview(webviewView.webview);

		this._disposables.push(webviewView.webview.onDidReceiveMessage(async (message) => {
			switch (message.command) {
				case 'cancel':
					await vscode.commands.executeCommand('microsoft-todo.cancelCreateTask');
					break;
				case 'create':
					console.log('api call goes here' + JSON.stringify(message));
					break;
			}
		}));
    }

    private getHtmlForWebview(webview: vscode.Webview) {
		// Get the local path to main script run in the webview, then convert it to a uri we can use in the webview.
		const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'webviews', 'taskCreateView', 'main.js'));

		// Do the same for the stylesheet.
		const styleResetUri = webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'webviews', 'common', 'reset.css'));
		const styleVSCodeUri = webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'webviews', 'common', 'vscode.css'));
		const styleMainUri = webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'webviews', 'taskCreateView', 'main.css'));

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
				<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource}; script-src 'nonce-${nonce}';">
				<meta name="viewport" content="width=device-width, initial-scale=1.0">
				<link href="${styleResetUri}" rel="stylesheet">
				<link href="${styleVSCodeUri}" rel="stylesheet">
				<link href="${styleMainUri}" rel="stylesheet">
				<title>Create task</title>
			</head>
			<body>
				CREATE FLOW
				<button class='cancel-button'>Cancel</button>
				<script nonce="${nonce}" src="${scriptUri}"></script>
			</body>
			</html>`;
	}
}
