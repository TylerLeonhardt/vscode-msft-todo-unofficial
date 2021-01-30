//@ts-check

// This script will be run within the webview itself
// It cannot access the main VS Code APIs directly.
(function () {
    //@ts-ignore
    const vscode = acquireVsCodeApi();

    // const taskNode = vscode.getState();

    // if (taskNode) {
    //     changeTaskNode(taskNode);
    // }

    // /** @type {Array<{ value: string }>} */
    // let colors = oldState.colors;

    // updateColorList(colors);

    // Handle messages sent from the extension to the webview
    window.addEventListener('message', event => {
        const taskNode = event.data; // The json data that the extension sent
        changeTaskNode(taskNode);
    });

    function changeTaskNode(taskNode) {
        const title = document.querySelector('.task-title');
        title.innerHTML = taskNode.entity.title;
        const body = document.querySelector('.task-body');
        body.innerHTML = taskNode.entity.body.content;
        // vscode.setState(taskNode);
    }

    vscode.postMessage({ command: 'ready' });

    /**
     * @param {Array<{ value: string }>} colors
     */
    // function updateColorList(colors) {
    //     const ul = document.querySelector('.color-list');
    //     ul.textContent = '';
    //     for (const color of colors) {
    //         const li = document.createElement('li');
    //         li.className = 'color-entry';

    //         const colorPreview = document.createElement('div');
    //         colorPreview.className = 'color-preview';
    //         colorPreview.style.backgroundColor = `#${color.value}`;
    //         colorPreview.addEventListener('click', () => {
    //             onColorClicked(color.value);
    //         });
    //         li.appendChild(colorPreview);

    //         const input = document.createElement('input');
    //         input.className = 'color-input';
    //         input.type = 'text';
    //         input.value = color.value;
    //         input.addEventListener('change', (e) => {
    //             const value = e.target.value;
    //             if (!value) {
    //                 // Treat empty value as delete
    //                 colors.splice(colors.indexOf(color), 1);
    //             } else {
    //                 color.value = value;
    //             }
    //             updateColorList(colors);
    //         });
    //         li.appendChild(input);

    //         ul.appendChild(li);
    //     }

    //     // Update the saved state
    //     vscode.setState({ colors: colors });
    // }

    // /** 
    //  * @param {string} color 
    //  */
    // function onColorClicked(color) {
    //     vscode.postMessage({ type: 'colorSelected', value: color });
    // }

    // /**
    //  * @returns string
    //  */
    // function getNewCalicoColor() {
    //     const colors = ['020202', 'f1eeee', 'a85b20', 'daab70', 'efcb99'];
    //     return colors[Math.floor(Math.random() * colors.length)];
    // }

    // function addColor() {
    //     colors.push({ value: getNewCalicoColor() });
    //     updateColorList(colors);
    // }
}());
