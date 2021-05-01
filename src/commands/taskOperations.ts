import { TodoTask, TodoTaskList } from '@microsoft/microsoft-graph-types';
import * as vscode from 'vscode';
import { MicrosoftToDoClientFactory } from '../clientFactories/microsoftToDoClientFactory';
import { ListNode, TaskNode } from "../todoProviders/microsoftToDoTreeDataProvider";

export class TaskOperations extends vscode.Disposable {
	private readonly disposables: vscode.Disposable[] = [];
	constructor(private readonly clientProvider: MicrosoftToDoClientFactory) {
		super(() => {
			this.disposables.forEach(d => d.dispose());
		});

		this.disposables.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.createTask',
			(list: ListNode | undefined) => this.createTask(list)));

		this.disposables.push(vscode.commands.registerCommand(
			'microsoft-todo-unoffcial.deleteTask',
			(task: TaskNode | undefined, tasks: TaskNode[] | undefined) => this.deleteTask(task, tasks)));
	}

	async getTask(listId: string, taskId: string): Promise<TodoTask | undefined> {
		const client = await this.clientProvider.getClient();
		if (!client) {
			await vscode.window.showErrorMessage('Not logged in');
			return;
		}

		// TODO: error handling
		return await client.api(`/me/todo/lists/${listId}/tasks/${taskId}`).get() as TodoTask;
	}

	async getTasks(listId: string): Promise<TodoTask[] | undefined> {
		const client = await this.clientProvider.getClient();
		if (!client) {
			await vscode.window.showErrorMessage('Not logged in');
			return;
		}

		// TODO: error handling
		return await this.clientProvider.getAll<TodoTask>(client, `/me/todo/lists/${listId}/tasks`) as TodoTask[];
	}

	async createTask(list: ListNode | undefined): Promise<any> {
		const client = await this.clientProvider.getClient();
		if (!client) {
			return await vscode.window.showErrorMessage('Please log in before creating a task.');
		}

		let listId = list?.entity.id;
		if (!listId) {
			const taskLists = new Array<TodoTaskList>();
			let iterUri: string | null | undefined = '/me/todo/lists';
			while (iterUri) {
				let res = await client.api(iterUri).get() as { '@odata.nextLink': string | null | undefined; value: TodoTaskList[] };
				res.value.forEach(r => taskLists.push(r));
				iterUri = res['@odata.nextLink'];
			}


			const quickPickItems: Array<vscode.QuickPickItem & { id?: string }> = taskLists.map(l => ({
				label: l.displayName || '',
				...l
			}));

			quickPickItems.push({
				label: 'Create a new list...',
				id: 'new',
			});

			const chosen = await vscode.window.showQuickPick(quickPickItems, {
				canPickMany: false,
				ignoreFocusOut: true,
				placeHolder: 'Which list would you like to add tasks to?'
			});

			listId = (chosen as TodoTaskList)?.id;

			if (listId === 'new') {
				const displayName = await vscode.window.showInputBox({
					prompt: 'Add a List',
					placeHolder: 'Groceries',
					ignoreFocusOut: true
				});

				// The user quit the prompt
				if (!displayName) {
					return;
				}

				// TODO: Error handling
				const res = await client.api('/me/todo/lists').post({
					displayName
				});

				listId = res.id;
		
				await vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
			}
		}

		// The user quit the prompt
		if (!listId) {
			return;
		}

		const inputBoxOptions: vscode.InputBoxOptions = {
			prompt: 'Add a Task',
			placeHolder: 'Eat my veggies'
		};

		let title = await vscode.window.showInputBox(inputBoxOptions);
		inputBoxOptions.prompt = 'Add another Task';
		while (title) {
			// TODO: error handling
			await client.api(`/me/todo/lists/${listId}/tasks`).post({
				title: title,
				body: {
					content: '',
					contentType: 'text'
				}
			});

			// not awaiting on purpose
			vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
			title = await vscode.window.showInputBox(inputBoxOptions);
		}
	}

	async deleteTask(task: TaskNode | undefined, tasks: TaskNode[] | undefined): Promise<any> {
		if (!task) {
			return;
		}

		const client = await this.clientProvider.getClient();
		if (!client) {
			return await vscode.window.showErrorMessage('Not logged in');
		}

		const expected = tasks?.length ? 'Delete tasks' : 'Delete task';

		const tasksFormatted = tasks ? `${tasks.map(t => t.entity.title).join('", "')}` : task.entity.title;
		const choice = await vscode.window.showWarningMessage(
			`"${tasksFormatted}" will be permanently deleted. You won't be able to undo this action.`, { modal: true },
			expected, 'Cancel');

		if (choice !== expected) {
			return;
		}

		tasks ??= [task];

		const promises = tasks.map(t => client.api(`/me/todo/lists/${t.parent.entity.id}/tasks/${t.entity.id}`).delete());

		// TODO: error handling
		await Promise.all(promises);

		await vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList', task.parent);
	}
}
