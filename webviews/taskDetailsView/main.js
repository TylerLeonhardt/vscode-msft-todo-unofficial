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
	/** @type HTMLInputElement */
	const reminderDate = document.querySelector('.task-reminder-date');
	/** @type HTMLInputElement */
	const reminderTime = document.querySelector('.task-reminder-time');
	/** @type HTMLButtonElement */
	const updateButton = document.querySelector('.update-task');
	/** @type HTMLButtonElement */
	const cancelButton = document.querySelector('.update-cancel');

	let currentNode;

	/**
	 * @param {string} prefix
	 * @param {string} formatStr
	 * @param {{dateTime: string;timeZone: string;}} graphDateTime
	 */
	function format(prefix, formatStr, graphDateTime) {
		const mo = graphDateTime.timeZone === 'UTC' ? moment.utc(graphDateTime.dateTime) : moment.tz(graphDateTime.dateTime, graphDateTime.timeZone);
		return mo.local().format(`[${prefix}]${formatStr}`);
	}

	const formatTime = (graphDateTime) => format('', 'HH:mm', graphDateTime);
	const formatDueDate = (graphDateTime) => format('Due ', 'l', graphDateTime);
	const formatReminderDate = (graphDateTime) => format('Remind at ', 'l', graphDateTime);

	// @ts-ignore
	TinyDatePicker(dueDateInput, {
		/**
		 * @param { Date } date
		 * @returns { string }
		 */
		format(date) {
			return moment(date).local().format('[Due] l');
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

			return moment(str).toDate();
		}
	});

	// @ts-ignore
	TinyDatePicker(reminderDate, {
		/**
		 * @param { Date } date
		 * @returns { string }
		 */
		format(date) {
			return moment(date).local().format('[Remind at] l');
		},

		/**
		 * @param {string} str
		 * @returns {Date}
		 */
		parse(str) {
			if (!str) {
				return new Date();
			}

			if (str.indexOf('Remind at ') !== -1) {
				str = str.split('Remind at ')[1];
			}

			return moment(str).toDate();
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

		/** @type {{dateTime: string;timeZone: string;}} */
		const due = taskNode.entity.dueDateTime;
		const remind = taskNode.entity.reminderDateTime;
		dueDateInput.value = due ? formatDueDate(due) : '';

		if (remind) {
			reminderDate.value = formatReminderDate(remind);
			reminderTime.type = 'time';
			reminderTime.value = formatTime(remind);
		} else {
			reminderDate.value = '';
			reminderTime.type = 'hidden';
			reminderTime.value = '';
		}
	}

	const onchange = () => {
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
						dueDate: dueDateInput.value ? dueDateInput.value.split('Due ')[1] : '',
						reminderDate: reminderDate.value ? reminderDate.value.split('Remind at ')[1] : '',
						reminderTime: reminderTime.value
					}
				});
				currentNode.entity.title = title.value;
				currentNode.entity.body.content = body.value;
			};

			cancelButton.onclick = () => {
				changeTaskNode(currentNode);
			};
		}
	};

	title.onkeydown = () => onchange();
	body.onkeydown = () => onchange();
	dueDateInput.onchange = () => onchange();
	reminderTime.onchange = () => onchange();
	reminderDate.onchange = () => {
		onchange();

		// handle showing the reminder time
		if (reminderDate.value) {
			reminderTime.type = 'time';
			if(!reminderTime.value) {
				reminderTime.value = moment().add(1, 'hours').format('HH:mm');
			}
		} else {
			reminderTime.type = 'hidden';
			reminderTime.value = '';
		}
	};

	const initialState = vscode.getState();
	if (initialState) {
		changeTaskNode(initialState);
	}

	vscode.postMessage({ command: 'ready' });
}());
