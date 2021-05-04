import * as vscode from 'vscode';
import { TodoTask, TodoTaskList } from '@microsoft/microsoft-graph-types';
import { MicrosoftToDoClientFactory } from '../clientFactories/microsoftToDoClientFactory';

export interface ListNode {
	nodeType: 'list';
	entity: TodoTaskList;
}

enum TaskStatusType {
	completed = 'Completed',
	inProgress = 'In Progress'
}

interface CreateListNode {
	nodeType: 'create-list';
}

interface ImportantListNode {
	nodeType: 'important-list';
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

type ToDoEntity = TaskNode | StatusNode | ListNode | CreateListNode | ImportantListNode;

export class MicrosoftToDoTreeDataProvider extends vscode.Disposable implements vscode.TreeDataProvider<ToDoEntity> {
	private didChangeTreeData = new vscode.EventEmitter<void | ToDoEntity | undefined>();
	onDidChangeTreeData?: vscode.Event<void | ToDoEntity | undefined> = this.didChangeTreeData.event;

	private readonly disposibles: vscode.Disposable[] = [];
	private importanceFilter = false; 
	private importantNode: ImportantListNode | undefined;

	constructor(private clientFactory: MicrosoftToDoClientFactory) {
		super(() => this.dispose());

		this.disposibles.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.refreshList',
			(element?: ToDoEntity) => this.didChangeTreeData.fire(element)));

		this.disposibles.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.complete',
			(node: TaskNode, nodes: TaskNode[] | undefined) => nodes ? this.changeCompletedState(nodes) : this.changeCompletedState([node])));

		this.disposibles.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.uncomplete',
			(node: TaskNode, nodes: TaskNode[] | undefined) => nodes ? this.changeCompletedState(nodes) : this.changeCompletedState([node])));

		this.disposibles.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.star',
			(node: TaskNode, nodes: TaskNode[] | undefined) => nodes ? this.changeImportanceState(nodes) : this.changeImportanceState([node])));

		this.disposibles.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.unstar',
			(node: TaskNode, nodes: TaskNode[] | undefined) => nodes ? this.changeImportanceState(nodes) : this.changeImportanceState([node])));

		this.disposibles.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.starFilter',
			async () => {
				this.importanceFilter = true;
				await vscode.commands.executeCommand('setContext', 'starFilter', true);
				this.didChangeTreeData.fire();
			}));

		this.disposibles.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.unstarFilter',
			async () => {
				this.importanceFilter = false;
				await vscode.commands.executeCommand('setContext', 'starFilter', false);
				this.didChangeTreeData.fire();
			}));
	}

	async changeCompletedState(nodes: TaskNode[]) {
		const client = await this.clientFactory.getClient();

		const promises = nodes.map(async n => {
			await client!.api(`/me/todo/lists/${n.parent.entity.id}/tasks/${n.entity.id}`).patch({
				status: n.entity.status === 'completed' ? 'notStarted' : 'completed'
			});
			this.didChangeTreeData.fire(n.parent);
		});

		// TODO: Error handling
		await Promise.all(promises);
	}

	async changeImportanceState(nodes: TaskNode[]) {
		const client = await this.clientFactory.getClient();

		const promises = nodes.map(async n => {
			await client!.api(`/me/todo/lists/${n.parent.entity.id}/tasks/${n.entity.id}`).patch({
				importance: n.entity.importance === 'high' ? 'normal' : 'high'
			});
			this.didChangeTreeData.fire(n.parent);
		});

		// TODO: Error handling
		await Promise.all(promises);
		if (this.importantNode) {
			this.didChangeTreeData.fire(this.importantNode);
		}
	}

	getTreeItem(element: ToDoEntity): vscode.TreeItem | Thenable<vscode.TreeItem> {
		let treeItem: vscode.TreeItem;
		switch (element.nodeType) {
			case 'create-list':
				treeItem = new vscode.TreeItem('Create a new list...');
				treeItem.command = {
					command: 'microsoft-todo-unoffcial.createList',
					title: 'Create a new list...'
				};
				break;
			case 'list':
				treeItem = new vscode.TreeItem({
					label: element.entity.displayName || '',
				}, vscode.TreeItemCollapsibleState.Collapsed);
				treeItem.contextValue = element.nodeType;
				if (element.entity.isShared) {
					treeItem.description = '👥';
				}
				break;
			case 'task':
				let label = element.entity.title || "";
				let tooltip = `*${label}*`;
				const dueDateTime = element.entity.dueDateTime;
				const highlights: [number, number][] = [];

				if (dueDateTime?.dateTime) {
					const dueStr = " DUE " + new Date(dueDateTime.dateTime).toLocaleDateString() + ' ';
					label += "  ";
					highlights.push([label.length, label.length + dueStr.length]);
					label += dueStr;
				}

				if (element.entity.body?.content) {
					tooltip += `\n\n${element.entity.body.content}`;
				}

				const treeItemLabel: vscode.TreeItemLabel = {
					label,
					highlights
				};

				treeItem = new vscode.TreeItem(treeItemLabel, vscode.TreeItemCollapsibleState.None);
				const status = element.entity.status === 'completed' ? 'completed' : 'notcompleted';
				const importance = element.entity.importance === 'high' ? 'starred' : 'notstarred';
				treeItem.contextValue = `${element.nodeType}-${status} ${element.nodeType}-${importance}`;
				treeItem.tooltip = new vscode.MarkdownString(tooltip, true);
				break;
			case 'status':
				const collapse = element.statusType === TaskStatusType.completed
					? vscode.TreeItemCollapsibleState.Collapsed
					: vscode.TreeItemCollapsibleState.Expanded;

				const statusLabel = ` ${element.statusType} `;
				treeItem = new vscode.TreeItem({
					highlights: [[0, statusLabel.length]],
					label: statusLabel
				}, collapse);
				break;
			case 'important-list':
				treeItem = new vscode.TreeItem({
					label: '⭐️ Important'
				}, vscode.TreeItemCollapsibleState.Collapsed);
				break;
		}

		return treeItem;
	}

	async getChildren(element?: ToDoEntity): Promise<ToDoEntity[] | undefined> {
		const client = await this.clientFactory.getClient();
		if (!client) {
			return;
		}

		if (!element) {
			const taskLists: TodoTaskList[] = await this.clientFactory.getAll(client, '/me/todo/lists');

			this.importantNode = { nodeType: 'important-list' };
			const nodes: ToDoEntity[] = [this.importantNode];
			taskLists.forEach(entity => nodes.push({ nodeType: 'list', entity }));
			nodes.push({ nodeType: 'create-list' });
			return nodes;
		}

		if (element.nodeType === 'list') {

			const getTasks = async (getCompleted: boolean): Promise<TaskNode[]> => {
				const comparison = getCompleted ? 'eq' : 'ne';
				let filter = `status ${comparison} 'completed'`;
				if (this.importanceFilter) {
					filter += ` and importance eq 'high'`;
				}
				const tasks: TodoTask[] = await this.clientFactory.getAll(client, `/me/todo/lists/${(element.entity as TodoTaskList).id}/tasks?$filter=${filter}`);

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

		if (element.nodeType === 'important-list') {
			const filter = `status ne 'completed' and importance eq 'high'`;
			const listEntities: TodoTaskList[] = await this.clientFactory.getAll(client, '/me/todo/lists');
			const entities: ToDoEntity[] = [];
			for (const entity of listEntities) {
				const tasks: TodoTask[] = await this.clientFactory.getAll(client, `/me/todo/lists/${entity.id}/tasks?$filter=${filter}`);

				tasks.forEach(t => {
					entities.push({
						nodeType: 'task',
						entity: t,
						parent: {
							entity,
							nodeType: 'list'
						}
					});
				});
			}

			return entities;
		}
	}

	public dispose() {
		this.disposibles.forEach(d => d.dispose());
	}
}
