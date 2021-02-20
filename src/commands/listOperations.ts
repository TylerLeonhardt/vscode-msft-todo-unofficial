import { TodoTaskList } from '@microsoft/microsoft-graph-types';
import * as vscode from 'vscode';
import { MicrosoftToDoClientFactory } from '../clientFactories/microsoftToDoClientFactory';
import { ListNode } from "../todoProviders/microsoftToDoTreeDataProvider";

export class ListOperations extends vscode.Disposable {
	private readonly disposables: vscode.Disposable[] = [];
	constructor(private readonly clientProvider: MicrosoftToDoClientFactory) {
		super(() => {
			this.disposables.forEach(d => d.dispose());
		});

		this.disposables.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.createList',
			() => this.createList()));

		this.disposables.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.deleteList',
			(list: ListNode | undefined, lists: ListNode[] | undefined) => this.deleteList(list, lists)));
	}

	async getList(listId: string): Promise<TodoTaskList | undefined> {
		const client = await this.clientProvider.getClient();
		if (!client) {
			await vscode.window.showErrorMessage('Not logged in');
			return;
		}

		// TODO: error handling
		return (await client.api(`/me/todo/lists/${listId}`).get()).value as TodoTaskList;
	}

	async getLists(): Promise<TodoTaskList[] | undefined> {
		const client = await this.clientProvider.getClient();
		if (!client) {
			await vscode.window.showErrorMessage('Not logged in');
			return;
		}

		// TODO: error handling
		return await this.clientProvider.getAll<TodoTaskList>(client, `/me/todo/lists`) as TodoTaskList[];
	}

	public async createList(displayName?: string): Promise<void> {
		const client = await this.clientProvider.getClient();
		if (!client) {
			await vscode.window.showErrorMessage('Not logged in');
			return;
		}

		displayName ??= await vscode.window.showInputBox({
			prompt: 'Add a List',
			placeHolder: 'Groceries',
			ignoreFocusOut: true
		});

		if (!displayName) {
			return;
		}

		await client.api('/me/todo/lists').post({
			displayName
		});

		await vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
	}

	public async deleteList(list: ListNode | undefined, lists: ListNode[] | undefined): Promise<any> {
		if (!list) {
			return;
		}

		const client = await this.clientProvider.getClient();
		if (!client) {
			return await vscode.window.showErrorMessage('Not logged in');
		}

		const expected = lists?.length ? 'Delete lists' : 'Delete list';

		const tasksFormatted = lists ? `${lists.map(t => t.entity.displayName).join('", "')}` : list.entity.displayName;
		const choice = await vscode.window.showWarningMessage(
			`"${tasksFormatted}" will be permanently deleted. You won't be able to undo this action.`,
			expected, 'Cancel');

		if (choice !== expected) {
			return;
		}

		lists ??= [list];

		const promises = lists.map(t => client.api(`/me/todo/lists/${t.entity.id}`).delete());

		// TODO: error handling
		await Promise.all(promises);

		await vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
	}
}
