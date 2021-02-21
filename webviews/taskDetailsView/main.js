//@ts-check

// This script will be run within the webview itself
// It cannot access the main VS Code APIs directly.
(function () {
	//@ts-ignore
	const vscode = acquireVsCodeApi();
	/** @type HTMLInputElement */
	const title = document.querySelector('.task-title');
	/** @type HTMLTextAreaElement */
	const body = document.querySelector('.task-body');
	/** @type HTMLInputElement */
	const dueDateInput = document.querySelector('.task-duedate');
	/** @type HTMLButtonElement */
	const remindDate = document.querySelector('.task-reminder');
	/** @type HTMLButtonElement */
	const updateButton = document.querySelector('.update-task');
	/** @type HTMLButtonElement */
	const cancelButton = document.querySelector('.update-cancel');
	let currentNode;

	/**
	 * @param { Date | string } date
	 * @returns { string }
	 */
	function formatDueDate(date) {
		if (typeof date === 'string' || date instanceof String) {
			date = new Date(date);
		}
		return `Due ${date.toLocaleDateString()}`;
	}

	// @ts-ignore
	TinyDatePicker(dueDateInput, {
		/**
		 * @param { Date } date
		 * @returns { string }
		 */
		format(date) {
			return formatDueDate(date);
		},

		/**
		 * @param {string} str
		 * @returns {Date}
		 */
		parse(str) {
			if (!str) {
				return new Date();
			}

			if (str.indexOf('Due ') !== -1) {
				str = str.split('Due ')[1];
			}

			const [ month, day, year ] = str.split('/');
			var date = new Date(parseInt(year, 10), parseInt(month, 10) - 1, parseInt(day, 10));
			return date || new Date();
		}
	});

	// Handle messages sent from the extension to the webview
	window.addEventListener('message', event => {
		const taskNode = event.data; // The json data that the extension sent
		changeTaskNode(taskNode);
	});

	function changeTaskNode(taskNode) {
		vscode.setState(taskNode);
		currentNode = taskNode;
		title.value = taskNode.entity.title;
		body.value = taskNode.entity.body.content;
		updateButton.hidden = true;
		cancelButton.hidden = true;

		dueDateInput.value = taskNode.entity.dueDateTime ? formatDueDate(taskNode.entity.dueDateTime.dateTime) : '';

		if (taskNode.entity.reminderDateTime) {
			remindDate.hidden = false;
			remindDate.innerHTML = `Reminder set at ${new Date(taskNode.entity.reminderDateTime.dateTime).toLocaleString()}`;
		} else {
			remindDate.hidden = true;
		}
	}

	const onkeydown = () => {
		// TODO: only do this if the content had changed
		if (updateButton.hidden) {
			updateButton.hidden = false;
			cancelButton.hidden = false;
			updateButton.onclick = () => {
				updateButton.hidden = true;
				cancelButton.hidden = true;

				
				vscode.postMessage({
					command: 'update',
					body: {
						title: title.value,
						note: body.value,
						id: currentNode.entity.id,
						listId: currentNode.parent.entity.id,
						dueDate: dueDateInput.value ? dueDateInput.value.split('Due ')[1] : ''
					}
				});
				currentNode.entity.title = title.value;
				currentNode.entity.body.content = body.value;
			};

			cancelButton.onclick = () => {
				updateButton.hidden = true;
				cancelButton.hidden = true;
				title.value = currentNode.entity.title;
				body.value = currentNode.entity.body.content;
				dueDateInput.value = currentNode.entity.dueDateTime ? formatDueDate(currentNode.entity.dueDateTime.dateTime) : '';
			};
		}
	};

	title.onkeydown = () => onkeydown();
	body.onkeydown = () => onkeydown();
	dueDateInput.onchange = () => onkeydown();

	const initialState = vscode.getState();
	if (initialState) {
		changeTaskNode(initialState);
	}

	vscode.postMessage({ command: 'ready' });
}());
