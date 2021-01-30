import * as vscode from 'vscode';
import 'isomorphic-fetch';
import { Client } from "@microsoft/microsoft-graph-client";
import { Entity, TodoTask, TodoTaskList } from '@microsoft/microsoft-graph-types';

const redirectUri = encodeURIComponent(`${vscode.env.uriScheme}://tylerleonhardt.msft-todo/`);
const scopes = ['Tasks.ReadWrite'];
const uri = vscode.Uri.parse(`https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=a4fd7674-4ebd-4dbc-831c-338314dd459e&response_type=token&redirect_uri=${redirectUri}&response_mode=fragment&scope=${scopes}`);

interface ListNode {
	nodeType: 'list';
	entity: TodoTaskList;
}

enum TaskStatusType {
	completed = 'Completed',
	inProgress = 'In Progress'
}

interface StatusNode {
	nodeType: 'status';
	statusType: TaskStatusType;
	getChildren: () => Promise<TaskNode[]>;
}

export interface TaskNode {
	nodeType: 'task';
	entity: TodoTask;
	parent: ListNode;
}

interface LogInNode {
	nodeType: 'login';
}

type ToDoEntity = TaskNode | StatusNode | ListNode | LogInNode;

export class MicrosoftToDoTreeDataProvider extends vscode.Disposable implements vscode.TreeDataProvider<ToDoEntity> {
	private didChangeTreeData = new vscode.EventEmitter<void | ToDoEntity | undefined>();
	onDidChangeTreeData?: vscode.Event<void | ToDoEntity | undefined> = this.didChangeTreeData.event;

	private readonly loginCommand: vscode.Disposable;
	private readonly refreshCommand: vscode.Disposable;
	private readonly uriHandler: vscode.Disposable;

	private _token: string | null = null;
	private async getClient(): Promise<Client | undefined> {
		// const session = await vscode.authentication.getSession('microsoft', scopes, { createIfNone: true });

		if (!this._token) {
			// await vscode.env.openExternal(uri);
			return;
		}

		return Client.init({
			authProvider: (done) => {
				// done(undefined, session.accessToken);
				done(undefined, this._token);
			}
		});
	}

	constructor() {
		super(() => this.dispose());

		this.uriHandler = vscode.window.registerUriHandler({
			handleUri: (uri) => {
				const fragmentResult = new Map<string, string>();
				uri.fragment.split('&').forEach(q => {
					const [key, value] = q.split('=');
					fragmentResult.set(key, value);
				});

				if (fragmentResult.has('access_token')) {
					this._token = fragmentResult.get('access_token')!;
					this.didChangeTreeData.fire();
					return;
				}

				console.log(uri.query);
			}
		});

		this.loginCommand = vscode.commands.registerCommand(
			'microsoft-todo.login',
			() => vscode.env.openExternal(uri));

		this.refreshCommand = vscode.commands.registerCommand(
			'microsoft-todo.refreshList',
			(element?: ToDoEntity) => this.didChangeTreeData.fire(element));

		this.refreshCommand = vscode.commands.registerCommand(
			'microsoft-todo.complete',
			(element: TaskNode) => this.changeState(element));

		this.refreshCommand = vscode.commands.registerCommand(
			'microsoft-todo.uncomplete',
			(element: TaskNode) => this.changeState(element));
	}

	async changeState(element: TaskNode) {
		const client = await this.getClient();

		await client!.api(`/me/todo/lists/${element.parent.entity.id}/tasks/${element.entity.id}`).patch({
			status: element.entity.status === 'completed' ? 'notStarted' : 'completed'
		});

		this.didChangeTreeData.fire(element.parent);
	}

	getTreeItem(element: ToDoEntity): vscode.TreeItem | Thenable<vscode.TreeItem> {
		let treeItem: vscode.TreeItem;
		switch (element.nodeType) {
			case 'login':
				treeItem = new vscode.TreeItem('Log in to see your Microsoft To-Do lists');
				treeItem.command = {
					command: 'microsoft-todo.login',
					title: 'Click to login to Microsoft To-Do',
				};
				break;
			case 'list':
				treeItem = new vscode.TreeItem(element.entity.displayName || "", vscode.TreeItemCollapsibleState.Collapsed);
				treeItem.contextValue = element.nodeType;
				break;
			case 'task':
				treeItem = new vscode.TreeItem(element.entity.title || "", vscode.TreeItemCollapsibleState.None);
				break;
			case 'status':
				const collapse = element.statusType === TaskStatusType.completed
					? vscode.TreeItemCollapsibleState.Collapsed
					: vscode.TreeItemCollapsibleState.Expanded;

				treeItem = new vscode.TreeItem(element.statusType, collapse);
				break;
		}

		treeItem.contextValue = element.nodeType;
		return treeItem;
	}

	async getChildren(element?: ToDoEntity): Promise<ToDoEntity[] | undefined> {
		const client = await this.getClient();
		if (!client) {
			return; // [{ nodeType: 'login' }];
		}

		if (!element) {
			const taskLists = new Array<TodoTaskList>();

			let iterUri: string | null | undefined = '/me/todo/lists';
			while (iterUri) {
				let res = await client.api(iterUri).get() as { '@odata.nextLink': string | null | undefined; value: TodoTaskList[] };
				res.value.forEach(r => taskLists.push(r));
				iterUri = res['@odata.nextLink'];
			}

			return taskLists.map(entity => ({ nodeType: 'list', entity }));
		}

		if (element.nodeType === 'list') {

			const getTasks = async (getCompleted: boolean): Promise<TaskNode[]> => {
				const tasks = new Array<TodoTask>();

				const comparison = getCompleted ? 'eq' : 'ne';
				let iterUri: string | null | undefined = `/me/todo/lists/${(element.entity as TodoTaskList).id}/tasks?$filter= status ${comparison} 'completed'`;
				while (iterUri) {
					let res = await client.api(iterUri).get() as { '@odata.nextLink': string | null | undefined; value: TodoTask[] };
					res.value.forEach(r => tasks.push(r));
					iterUri = res['@odata.nextLink'];
				}

				return tasks.map(entity => ({
					nodeType: 'task',
					entity,
					parent: element
				}));
			};

			return [
				{
					nodeType: 'status',
					getChildren: () => getTasks(false),
					statusType: TaskStatusType.inProgress,
				},
				{
					nodeType: 'status',
					getChildren: () => getTasks(true),
					statusType: TaskStatusType.completed,
				}
			];
		}

		if (element.nodeType === 'status') {
			return await element.getChildren();
		}
	}

	public dispose() {
		this.uriHandler.dispose();
		this.loginCommand.dispose();
	}
}
