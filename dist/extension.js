/******/ (() => { // webpackBootstrap
/******/ 	var __webpack_modules__ = ([
/* 0 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.deactivate = exports.activate = void 0;
const vscode = __webpack_require__(1);
const microsoftToDoClientFactory_1 = __webpack_require__(2);
const microsoftToDoTreeDataProvider_1 = __webpack_require__(47);
const taskDetailsView_1 = __webpack_require__(48);
const taskOperations_1 = __webpack_require__(50);
const listOperations_1 = __webpack_require__(51);
__webpack_require__(52);
const MsaAuthProvider_1 = __webpack_require__(60);
function activate(context) {
    return __awaiter(this, void 0, void 0, function* () {
        const auth = new MsaAuthProvider_1.MsaAuthProvider(context);
        yield auth.initialize();
        vscode.authentication.registerAuthenticationProvider(MsaAuthProvider_1.MsaAuthProvider.id, 'Microsoft Account (MSA)', auth, {
            supportsMultipleAccounts: false
        });
        const clientProvider = new microsoftToDoClientFactory_1.MicrosoftToDoClientFactory(context.globalState);
        const loginType = context.globalState.get('microsoftToDoUnofficialLoginType');
        if (loginType) {
            clientProvider.setLoginType(loginType.type);
        }
        let disposable;
        context.subscriptions.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.login', () => __awaiter(this, void 0, void 0, function* () {
            const result = yield vscode.window.showQuickPick(['Microsoft account', 'Work or School account']);
            if (!result) {
                return;
            }
            const provider = result === 'Microsoft account' ? 'msa' : 'microsoft';
            yield vscode.authentication.getSession(provider, microsoftToDoClientFactory_1.MicrosoftToDoClientFactory.scopes, { createIfNone: true });
            yield context.globalState.update('microsoftToDoUnofficialLoginType', { type: provider });
            clientProvider.setLoginType(provider);
            vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
        })));
        context.subscriptions.push(disposable = vscode.authentication.onDidChangeSessions((e) => clientProvider.clearLoginTypeState(e)));
        const treeDataProvider = new microsoftToDoTreeDataProvider_1.MicrosoftToDoTreeDataProvider(clientProvider);
        const view = vscode.window.createTreeView('microsoft-todo-unoffcial.listView', {
            treeDataProvider,
            showCollapseAll: true,
            canSelectMany: true
        });
        context.subscriptions.push(view);
        const taskDetailsProvider = new taskDetailsView_1.TaskDetailsViewProvider(context.extensionUri, clientProvider);
        view.onDidChangeSelection((e) => __awaiter(this, void 0, void 0, function* () {
            if (e.selection.length > 0) {
                const node = e.selection[0];
                if (node.nodeType === 'task') {
                    yield taskDetailsProvider.changeChosenView(node);
                }
            }
        }));
        const detailsView = vscode.window.registerWebviewViewProvider(taskDetailsProvider.viewType, taskDetailsProvider);
        context.subscriptions.push(detailsView);
        context.subscriptions.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.showTaskDetailsView', () => __awaiter(this, void 0, void 0, function* () {
            yield vscode.commands.executeCommand('setContext', 'showTaskDetailsView', true);
        })));
        context.subscriptions.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.hideTaskDetailsView', () => __awaiter(this, void 0, void 0, function* () {
            yield vscode.commands.executeCommand('setContext', 'showTaskDetailsView', false);
        })));
        const taskOps = new taskOperations_1.TaskOperations(clientProvider);
        const listOps = new listOperations_1.ListOperations(clientProvider);
        context.subscriptions.push(taskOps);
        context.subscriptions.push(listOps);
    });
}
exports.activate = activate;
// this method is called when your extension is deactivated
function deactivate() { }
exports.deactivate = deactivate;


/***/ }),
/* 1 */
/***/ ((module) => {

"use strict";
module.exports = require("vscode");;

/***/ }),
/* 2 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.MicrosoftToDoClientFactory = void 0;
const microsoft_graph_client_1 = __webpack_require__(3);
const vscode = __webpack_require__(1);
class MicrosoftToDoClientFactory {
    constructor(globalState) {
        this.globalState = globalState;
    }
    getClient() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.loginType) {
                return;
            }
            this.session = yield vscode.authentication.getSession(this.loginType, MicrosoftToDoClientFactory.scopes);
            if (!this.session) {
                return;
            }
            return microsoft_graph_client_1.Client.init({
                authProvider: (done) => {
                    done(undefined, this.session.accessToken);
                }
            });
        });
    }
    getAll(client, apiPath) {
        return __awaiter(this, void 0, void 0, function* () {
            let iterUri = apiPath;
            const list = new Array();
            while (iterUri) {
                let res = yield client.api(iterUri).get();
                res.value.forEach(r => list.push(r));
                iterUri = res['@odata.nextLink'];
            }
            return list;
        });
    }
    setLoginType(type) {
        this.loginType = type;
    }
    clearLoginTypeState(e) {
        return __awaiter(this, void 0, void 0, function* () {
            if (e.provider.id !== 'msa' && e.provider.id !== 'microsoft') {
                return;
            }
            // we already cleared the state
            if (!this.loginType) {
                return;
            }
            yield this.globalState.update('microsoftToDoUnofficialLoginType', {});
            this.setLoginType(undefined);
            yield vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
        });
    }
}
exports.MicrosoftToDoClientFactory = MicrosoftToDoClientFactory;
MicrosoftToDoClientFactory.scopes = ['Tasks.ReadWrite', 'offline_access', 'openid', 'profile'];


/***/ }),
/* 3 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "BatchRequestContent": () => (/* reexport safe */ _content_BatchRequestContent__WEBPACK_IMPORTED_MODULE_0__.BatchRequestContent),
/* harmony export */   "BatchResponseContent": () => (/* reexport safe */ _content_BatchResponseContent__WEBPACK_IMPORTED_MODULE_1__.BatchResponseContent),
/* harmony export */   "AuthenticationHandler": () => (/* reexport safe */ _middleware_AuthenticationHandler__WEBPACK_IMPORTED_MODULE_2__.AuthenticationHandler),
/* harmony export */   "HTTPMessageHandler": () => (/* reexport safe */ _middleware_HTTPMessageHandler__WEBPACK_IMPORTED_MODULE_3__.HTTPMessageHandler),
/* harmony export */   "RetryHandler": () => (/* reexport safe */ _middleware_RetryHandler__WEBPACK_IMPORTED_MODULE_4__.RetryHandler),
/* harmony export */   "RedirectHandler": () => (/* reexport safe */ _middleware_RedirectHandler__WEBPACK_IMPORTED_MODULE_5__.RedirectHandler),
/* harmony export */   "TelemetryHandler": () => (/* reexport safe */ _middleware_TelemetryHandler__WEBPACK_IMPORTED_MODULE_6__.TelemetryHandler),
/* harmony export */   "MiddlewareFactory": () => (/* reexport safe */ _middleware_MiddlewareFactory__WEBPACK_IMPORTED_MODULE_7__.MiddlewareFactory),
/* harmony export */   "AuthenticationHandlerOptions": () => (/* reexport safe */ _middleware_options_AuthenticationHandlerOptions__WEBPACK_IMPORTED_MODULE_8__.AuthenticationHandlerOptions),
/* harmony export */   "RetryHandlerOptions": () => (/* reexport safe */ _middleware_options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_9__.RetryHandlerOptions),
/* harmony export */   "RedirectHandlerOptions": () => (/* reexport safe */ _middleware_options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_10__.RedirectHandlerOptions),
/* harmony export */   "FeatureUsageFlag": () => (/* reexport safe */ _middleware_options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_11__.FeatureUsageFlag),
/* harmony export */   "TelemetryHandlerOptions": () => (/* reexport safe */ _middleware_options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_11__.TelemetryHandlerOptions),
/* harmony export */   "ChaosHandlerOptions": () => (/* reexport safe */ _middleware_options_ChaosHandlerOptions__WEBPACK_IMPORTED_MODULE_12__.ChaosHandlerOptions),
/* harmony export */   "ChaosStrategy": () => (/* reexport safe */ _middleware_options_ChaosStrategy__WEBPACK_IMPORTED_MODULE_13__.ChaosStrategy),
/* harmony export */   "ChaosHandler": () => (/* reexport safe */ _middleware_ChaosHandler__WEBPACK_IMPORTED_MODULE_14__.ChaosHandler),
/* harmony export */   "LargeFileUploadTask": () => (/* reexport safe */ _tasks_LargeFileUploadTask__WEBPACK_IMPORTED_MODULE_15__.LargeFileUploadTask),
/* harmony export */   "OneDriveLargeFileUploadTask": () => (/* reexport safe */ _tasks_OneDriveLargeFileUploadTask__WEBPACK_IMPORTED_MODULE_16__.OneDriveLargeFileUploadTask),
/* harmony export */   "PageIterator": () => (/* reexport safe */ _tasks_PageIterator__WEBPACK_IMPORTED_MODULE_17__.PageIterator),
/* harmony export */   "Client": () => (/* reexport safe */ _Client__WEBPACK_IMPORTED_MODULE_18__.Client),
/* harmony export */   "CustomAuthenticationProvider": () => (/* reexport safe */ _CustomAuthenticationProvider__WEBPACK_IMPORTED_MODULE_19__.CustomAuthenticationProvider),
/* harmony export */   "GraphError": () => (/* reexport safe */ _GraphError__WEBPACK_IMPORTED_MODULE_20__.GraphError),
/* harmony export */   "GraphRequest": () => (/* reexport safe */ _GraphRequest__WEBPACK_IMPORTED_MODULE_21__.GraphRequest),
/* harmony export */   "ImplicitMSALAuthenticationProvider": () => (/* reexport safe */ _ImplicitMSALAuthenticationProvider__WEBPACK_IMPORTED_MODULE_22__.ImplicitMSALAuthenticationProvider),
/* harmony export */   "MSALAuthenticationProviderOptions": () => (/* reexport safe */ _MSALAuthenticationProviderOptions__WEBPACK_IMPORTED_MODULE_23__.MSALAuthenticationProviderOptions),
/* harmony export */   "ResponseType": () => (/* reexport safe */ _ResponseType__WEBPACK_IMPORTED_MODULE_24__.ResponseType)
/* harmony export */ });
/* harmony import */ var _content_BatchRequestContent__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(4);
/* harmony import */ var _content_BatchResponseContent__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(7);
/* harmony import */ var _middleware_AuthenticationHandler__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(8);
/* harmony import */ var _middleware_HTTPMessageHandler__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(13);
/* harmony import */ var _middleware_RetryHandler__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(14);
/* harmony import */ var _middleware_RedirectHandler__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(16);
/* harmony import */ var _middleware_TelemetryHandler__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(18);
/* harmony import */ var _middleware_MiddlewareFactory__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(22);
/* harmony import */ var _middleware_options_AuthenticationHandlerOptions__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(11);
/* harmony import */ var _middleware_options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(15);
/* harmony import */ var _middleware_options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(17);
/* harmony import */ var _middleware_options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(12);
/* harmony import */ var _middleware_options_ChaosHandlerOptions__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(23);
/* harmony import */ var _middleware_options_ChaosStrategy__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(24);
/* harmony import */ var _middleware_ChaosHandler__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(25);
/* harmony import */ var _tasks_LargeFileUploadTask__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(27);
/* harmony import */ var _tasks_OneDriveLargeFileUploadTask__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(29);
/* harmony import */ var _tasks_PageIterator__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(31);
/* harmony import */ var _Client__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(32);
/* harmony import */ var _CustomAuthenticationProvider__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(33);
/* harmony import */ var _GraphError__WEBPACK_IMPORTED_MODULE_20__ = __webpack_require__(36);
/* harmony import */ var _GraphRequest__WEBPACK_IMPORTED_MODULE_21__ = __webpack_require__(34);
/* harmony import */ var _ImplicitMSALAuthenticationProvider__WEBPACK_IMPORTED_MODULE_22__ = __webpack_require__(42);
/* harmony import */ var _MSALAuthenticationProviderOptions__WEBPACK_IMPORTED_MODULE_23__ = __webpack_require__(46);
/* harmony import */ var _ResponseType__WEBPACK_IMPORTED_MODULE_24__ = __webpack_require__(38);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

























//# sourceMappingURL=index.js.map

/***/ }),
/* 4 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "BatchRequestContent": () => (/* binding */ BatchRequestContent)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(6);
/* harmony import */ var _RequestMethod__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(5);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @module BatchRequestContent
 */

/**
 * @class
 * Class for handling BatchRequestContent
 */
class BatchRequestContent {
    /**
     * @public
     * @constructor
     * Constructs a BatchRequestContent instance
     * @param {BatchRequestStep[]} [requests] - Array of requests value
     * @returns An instance of a BatchRequestContent
     */
    constructor(requests) {
        this.requests = new Map();
        if (typeof requests !== "undefined") {
            const limit = BatchRequestContent.requestLimit;
            if (requests.length > limit) {
                const error = new Error(`Maximum requests limit exceeded, Max allowed number of requests are ${limit}`);
                error.name = "Limit Exceeded Error";
                throw error;
            }
            for (const req of requests) {
                this.addRequest(req);
            }
        }
    }
    /**
     * @private
     * @static
     * Validates the dependency chain of the requests
     *
     * Note:
     * Individual requests can depend on other individual requests. Currently, requests can only depend on a single other request, and must follow one of these three patterns:
     * 1. Parallel - no individual request states a dependency in the dependsOn property.
     * 2. Serial - all individual requests depend on the previous individual request.
     * 3. Same - all individual requests that state a dependency in the dependsOn property, state the same dependency.
     * As JSON batching matures, these limitations will be removed.
     * @see {@link https://developer.microsoft.com/en-us/graph/docs/concepts/known_issues#json-batching}
     *
     * @param {Map<string, BatchRequestStep>} requests - The map of requests.
     * @returns The boolean indicating the validation status
     */
    static validateDependencies(requests) {
        const isParallel = (reqs) => {
            const iterator = reqs.entries();
            let cur = iterator.next();
            while (!cur.done) {
                const curReq = cur.value[1];
                if (curReq.dependsOn !== undefined && curReq.dependsOn.length > 0) {
                    return false;
                }
                cur = iterator.next();
            }
            return true;
        };
        const isSerial = (reqs) => {
            const iterator = reqs.entries();
            let cur = iterator.next();
            const firstRequest = cur.value[1];
            if (firstRequest.dependsOn !== undefined && firstRequest.dependsOn.length > 0) {
                return false;
            }
            let prev = cur;
            cur = iterator.next();
            while (!cur.done) {
                const curReq = cur.value[1];
                if (curReq.dependsOn === undefined || curReq.dependsOn.length !== 1 || curReq.dependsOn[0] !== prev.value[1].id) {
                    return false;
                }
                prev = cur;
                cur = iterator.next();
            }
            return true;
        };
        const isSame = (reqs) => {
            const iterator = reqs.entries();
            let cur = iterator.next();
            const firstRequest = cur.value[1];
            let dependencyId;
            if (firstRequest.dependsOn === undefined || firstRequest.dependsOn.length === 0) {
                dependencyId = firstRequest.id;
            }
            else {
                if (firstRequest.dependsOn.length === 1) {
                    const fDependencyId = firstRequest.dependsOn[0];
                    if (fDependencyId !== firstRequest.id && reqs.has(fDependencyId)) {
                        dependencyId = fDependencyId;
                    }
                    else {
                        return false;
                    }
                }
                else {
                    return false;
                }
            }
            cur = iterator.next();
            while (!cur.done) {
                const curReq = cur.value[1];
                if ((curReq.dependsOn === undefined || curReq.dependsOn.length === 0) && dependencyId !== curReq.id) {
                    return false;
                }
                if (curReq.dependsOn !== undefined && curReq.dependsOn.length !== 0) {
                    if (curReq.dependsOn.length === 1 && (curReq.id === dependencyId || curReq.dependsOn[0] !== dependencyId)) {
                        return false;
                    }
                    if (curReq.dependsOn.length > 1) {
                        return false;
                    }
                }
                cur = iterator.next();
            }
            return true;
        };
        if (requests.size === 0) {
            const error = new Error("Empty requests map, Please provide at least one request.");
            error.name = "Empty Requests Error";
            throw error;
        }
        return isParallel(requests) || isSerial(requests) || isSame(requests);
    }
    /**
     * @private
     * @static
     * @async
     * Converts Request Object instance to a JSON
     * @param {IsomorphicRequest} request - The IsomorphicRequest Object instance
     * @returns A promise that resolves to JSON representation of a request
     */
    static getRequestData(request) {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            const requestData = {
                url: "",
            };
            const hasHttpRegex = new RegExp("^https?://");
            // Stripping off hostname, port and url scheme
            requestData.url = hasHttpRegex.test(request.url) ? "/" + request.url.split(/.*?\/\/.*?\//)[1] : request.url;
            requestData.method = request.method;
            const headers = {};
            request.headers.forEach((value, key) => {
                headers[key] = value;
            });
            if (Object.keys(headers).length) {
                requestData.headers = headers;
            }
            if (request.method === _RequestMethod__WEBPACK_IMPORTED_MODULE_0__.RequestMethod.PATCH || request.method === _RequestMethod__WEBPACK_IMPORTED_MODULE_0__.RequestMethod.POST || request.method === _RequestMethod__WEBPACK_IMPORTED_MODULE_0__.RequestMethod.PUT) {
                requestData.body = yield BatchRequestContent.getRequestBody(request);
            }
            /**
             * TODO: Check any other property needs to be used from the Request object and add them
             */
            return requestData;
        });
    }
    /**
     * @private
     * @static
     * @async
     * Gets the body of a Request object instance
     * @param {IsomorphicRequest} request - The IsomorphicRequest object instance
     * @returns The Promise that resolves to a body value of a Request
     */
    static getRequestBody(request) {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            let bodyParsed = false;
            let body;
            try {
                const cloneReq = request.clone();
                body = yield cloneReq.json();
                bodyParsed = true;
            }
            catch (e) {
                // tslint:disable-line: no-empty
            }
            if (!bodyParsed) {
                try {
                    if (typeof Blob !== "undefined") {
                        const blob = yield request.blob();
                        const reader = new FileReader();
                        body = yield new Promise((resolve) => {
                            reader.addEventListener("load", () => {
                                const dataURL = reader.result;
                                /**
                                 * Some valid dataURL schemes:
                                 *  1. data:text/vnd-example+xyz;foo=bar;base64,R0lGODdh
                                 *  2. data:text/plain;charset=UTF-8;page=21,the%20data:1234,5678
                                 *  3. data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==
                                 *  4. data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==
                                 *  5. data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==
                                 * @see Syntax {@link https://en.wikipedia.org/wiki/Data_URI_scheme} for more
                                 */
                                const regex = new RegExp("^s*data:(.+?/.+?(;.+?=.+?)*)?(;base64)?,(.*)s*$");
                                const segments = regex.exec(dataURL);
                                resolve(segments[4]);
                            }, false);
                            reader.readAsDataURL(blob);
                        });
                    }
                    else if (typeof Buffer !== "undefined") {
                        const buffer = yield request.buffer();
                        body = buffer.toString("base64");
                    }
                    bodyParsed = true;
                }
                catch (e) {
                    // tslint:disable-line: no-empty
                }
            }
            return body;
        });
    }
    /**
     * @public
     * Adds a request to the batch request content
     * @param {BatchRequestStep} request - The request value
     * @returns The id of the added request
     */
    addRequest(request) {
        const limit = BatchRequestContent.requestLimit;
        if (request.id === "") {
            const error = new Error(`Id for a request is empty, Please provide an unique id`);
            error.name = "Empty Id For Request";
            throw error;
        }
        if (this.requests.size === limit) {
            const error = new Error(`Maximum requests limit exceeded, Max allowed number of requests are ${limit}`);
            error.name = "Limit Exceeded Error";
            throw error;
        }
        if (this.requests.has(request.id)) {
            const error = new Error(`Adding request with duplicate id ${request.id}, Make the id of the requests unique`);
            error.name = "Duplicate RequestId Error";
            throw error;
        }
        this.requests.set(request.id, request);
        return request.id;
    }
    /**
     * @public
     * Removes request from the batch payload and its dependencies from all dependents
     * @param {string} requestId - The id of a request that needs to be removed
     * @returns The boolean indicating removed status
     */
    removeRequest(requestId) {
        const deleteStatus = this.requests.delete(requestId);
        const iterator = this.requests.entries();
        let cur = iterator.next();
        /**
         * Removing dependencies where this request is present as a dependency
         */
        while (!cur.done) {
            const dependencies = cur.value[1].dependsOn;
            if (typeof dependencies !== "undefined") {
                const index = dependencies.indexOf(requestId);
                if (index !== -1) {
                    dependencies.splice(index, 1);
                }
                if (dependencies.length === 0) {
                    delete cur.value[1].dependsOn;
                }
            }
            cur = iterator.next();
        }
        return deleteStatus;
    }
    /**
     * @public
     * @async
     * Serialize content from BatchRequestContent instance
     * @returns The body content to make batch request
     */
    getContent() {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            const requests = [];
            const requestBody = {
                requests,
            };
            const iterator = this.requests.entries();
            let cur = iterator.next();
            if (cur.done) {
                const error = new Error("No requests added yet, Please add at least one request.");
                error.name = "Empty Payload";
                throw error;
            }
            if (!BatchRequestContent.validateDependencies(this.requests)) {
                const error = new Error(`Invalid dependency found, Dependency should be:
1. Parallel - no individual request states a dependency in the dependsOn property.
2. Serial - all individual requests depend on the previous individual request.
3. Same - all individual requests that state a dependency in the dependsOn property, state the same dependency.`);
                error.name = "Invalid Dependency";
                throw error;
            }
            while (!cur.done) {
                const requestStep = cur.value[1];
                const batchRequestData = (yield BatchRequestContent.getRequestData(requestStep.request));
                /**
                 * @see {@link https://developer.microsoft.com/en-us/graph/docs/concepts/json_batching#request-format}
                 */
                if (batchRequestData.body !== undefined && (batchRequestData.headers === undefined || batchRequestData.headers["content-type"] === undefined)) {
                    const error = new Error(`Content-type header is not mentioned for request #${requestStep.id}, For request having body, Content-type header should be mentioned`);
                    error.name = "Invalid Content-type header";
                    throw error;
                }
                batchRequestData.id = requestStep.id;
                if (requestStep.dependsOn !== undefined && requestStep.dependsOn.length > 0) {
                    batchRequestData.dependsOn = requestStep.dependsOn;
                }
                requests.push(batchRequestData);
                cur = iterator.next();
            }
            requestBody.requests = requests;
            return requestBody;
        });
    }
    /**
     * @public
     * Adds a dependency for a given dependent request
     * @param {string} dependentId - The id of the dependent request
     * @param {string} [dependencyId] - The id of the dependency request, if not specified the preceding request will be considered as a dependency
     * @returns Nothing
     */
    addDependency(dependentId, dependencyId) {
        if (!this.requests.has(dependentId)) {
            const error = new Error(`Dependent ${dependentId} does not exists, Please check the id`);
            error.name = "Invalid Dependent";
            throw error;
        }
        if (typeof dependencyId !== "undefined" && !this.requests.has(dependencyId)) {
            const error = new Error(`Dependency ${dependencyId} does not exists, Please check the id`);
            error.name = "Invalid Dependency";
            throw error;
        }
        if (typeof dependencyId !== "undefined") {
            const dependent = this.requests.get(dependentId);
            if (dependent.dependsOn === undefined) {
                dependent.dependsOn = [];
            }
            if (dependent.dependsOn.indexOf(dependencyId) !== -1) {
                const error = new Error(`Dependency ${dependencyId} is already added for the request ${dependentId}`);
                error.name = "Duplicate Dependency";
                throw error;
            }
            dependent.dependsOn.push(dependencyId);
        }
        else {
            const iterator = this.requests.entries();
            let prev;
            let cur = iterator.next();
            while (!cur.done && cur.value[1].id !== dependentId) {
                prev = cur;
                cur = iterator.next();
            }
            if (typeof prev !== "undefined") {
                const dId = prev.value[0];
                if (cur.value[1].dependsOn === undefined) {
                    cur.value[1].dependsOn = [];
                }
                if (cur.value[1].dependsOn.indexOf(dId) !== -1) {
                    const error = new Error(`Dependency ${dId} is already added for the request ${dependentId}`);
                    error.name = "Duplicate Dependency";
                    throw error;
                }
                cur.value[1].dependsOn.push(dId);
            }
            else {
                const error = new Error(`Can't add dependency ${dependencyId}, There is only a dependent request in the batch`);
                error.name = "Invalid Dependency Addition";
                throw error;
            }
        }
    }
    /**
     * @public
     * Removes a dependency for a given dependent request id
     * @param {string} dependentId - The id of the dependent request
     * @param {string} [dependencyId] - The id of the dependency request, if not specified will remove all the dependencies of that request
     * @returns The boolean indicating removed status
     */
    removeDependency(dependentId, dependencyId) {
        const request = this.requests.get(dependentId);
        if (typeof request === "undefined" || request.dependsOn === undefined || request.dependsOn.length === 0) {
            return false;
        }
        if (typeof dependencyId !== "undefined") {
            const index = request.dependsOn.indexOf(dependencyId);
            if (index === -1) {
                return false;
            }
            request.dependsOn.splice(index, 1);
            return true;
        }
        else {
            delete request.dependsOn;
            return true;
        }
    }
}
/**
 * @private
 * @static
 * Limit for number of requests {@link - https://developer.microsoft.com/en-us/graph/docs/concepts/known_issues#json-batching}
 */
BatchRequestContent.requestLimit = 20;
//# sourceMappingURL=BatchRequestContent.js.map

/***/ }),
/* 5 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "RequestMethod": () => (/* binding */ RequestMethod)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @enum
 * Enum for RequestMethods
 * @property {string} GET - The get request type
 * @property {string} PATCH - The patch request type
 * @property {string} POST - The post request type
 * @property {string} PUT - The put request type
 * @property {string} DELETE - The delete request type
 */
var RequestMethod;
(function (RequestMethod) {
    RequestMethod["GET"] = "GET";
    RequestMethod["PATCH"] = "PATCH";
    RequestMethod["POST"] = "POST";
    RequestMethod["PUT"] = "PUT";
    RequestMethod["DELETE"] = "DELETE";
})(RequestMethod || (RequestMethod = {}));
//# sourceMappingURL=RequestMethod.js.map

/***/ }),
/* 6 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "__extends": () => (/* binding */ __extends),
/* harmony export */   "__assign": () => (/* binding */ __assign),
/* harmony export */   "__rest": () => (/* binding */ __rest),
/* harmony export */   "__decorate": () => (/* binding */ __decorate),
/* harmony export */   "__param": () => (/* binding */ __param),
/* harmony export */   "__metadata": () => (/* binding */ __metadata),
/* harmony export */   "__awaiter": () => (/* binding */ __awaiter),
/* harmony export */   "__generator": () => (/* binding */ __generator),
/* harmony export */   "__createBinding": () => (/* binding */ __createBinding),
/* harmony export */   "__exportStar": () => (/* binding */ __exportStar),
/* harmony export */   "__values": () => (/* binding */ __values),
/* harmony export */   "__read": () => (/* binding */ __read),
/* harmony export */   "__spread": () => (/* binding */ __spread),
/* harmony export */   "__spreadArrays": () => (/* binding */ __spreadArrays),
/* harmony export */   "__await": () => (/* binding */ __await),
/* harmony export */   "__asyncGenerator": () => (/* binding */ __asyncGenerator),
/* harmony export */   "__asyncDelegator": () => (/* binding */ __asyncDelegator),
/* harmony export */   "__asyncValues": () => (/* binding */ __asyncValues),
/* harmony export */   "__makeTemplateObject": () => (/* binding */ __makeTemplateObject),
/* harmony export */   "__importStar": () => (/* binding */ __importStar),
/* harmony export */   "__importDefault": () => (/* binding */ __importDefault),
/* harmony export */   "__classPrivateFieldGet": () => (/* binding */ __classPrivateFieldGet),
/* harmony export */   "__classPrivateFieldSet": () => (/* binding */ __classPrivateFieldSet)
/* harmony export */ });
/*! *****************************************************************************
Copyright (c) Microsoft Corporation.

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
PERFORMANCE OF THIS SOFTWARE.
***************************************************************************** */
/* global Reflect, Promise */

var extendStatics = function(d, b) {
    extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return extendStatics(d, b);
};

function __extends(d, b) {
    extendStatics(d, b);
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
}

var __assign = function() {
    __assign = Object.assign || function __assign(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
        }
        return t;
    }
    return __assign.apply(this, arguments);
}

function __rest(s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
}

function __decorate(decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
}

function __param(paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
}

function __metadata(metadataKey, metadataValue) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(metadataKey, metadataValue);
}

function __awaiter(thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
}

function __generator(thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
}

function __createBinding(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}

function __exportStar(m, exports) {
    for (var p in m) if (p !== "default" && !exports.hasOwnProperty(p)) exports[p] = m[p];
}

function __values(o) {
    var s = typeof Symbol === "function" && Symbol.iterator, m = s && o[s], i = 0;
    if (m) return m.call(o);
    if (o && typeof o.length === "number") return {
        next: function () {
            if (o && i >= o.length) o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
    throw new TypeError(s ? "Object is not iterable." : "Symbol.iterator is not defined.");
}

function __read(o, n) {
    var m = typeof Symbol === "function" && o[Symbol.iterator];
    if (!m) return o;
    var i = m.call(o), r, ar = [], e;
    try {
        while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
    }
    catch (error) { e = { error: error }; }
    finally {
        try {
            if (r && !r.done && (m = i["return"])) m.call(i);
        }
        finally { if (e) throw e.error; }
    }
    return ar;
}

function __spread() {
    for (var ar = [], i = 0; i < arguments.length; i++)
        ar = ar.concat(__read(arguments[i]));
    return ar;
}

function __spreadArrays() {
    for (var s = 0, i = 0, il = arguments.length; i < il; i++) s += arguments[i].length;
    for (var r = Array(s), k = 0, i = 0; i < il; i++)
        for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
            r[k] = a[j];
    return r;
};

function __await(v) {
    return this instanceof __await ? (this.v = v, this) : new __await(v);
}

function __asyncGenerator(thisArg, _arguments, generator) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var g = generator.apply(thisArg, _arguments || []), i, q = [];
    return i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i;
    function verb(n) { if (g[n]) i[n] = function (v) { return new Promise(function (a, b) { q.push([n, v, a, b]) > 1 || resume(n, v); }); }; }
    function resume(n, v) { try { step(g[n](v)); } catch (e) { settle(q[0][3], e); } }
    function step(r) { r.value instanceof __await ? Promise.resolve(r.value.v).then(fulfill, reject) : settle(q[0][2], r); }
    function fulfill(value) { resume("next", value); }
    function reject(value) { resume("throw", value); }
    function settle(f, v) { if (f(v), q.shift(), q.length) resume(q[0][0], q[0][1]); }
}

function __asyncDelegator(o) {
    var i, p;
    return i = {}, verb("next"), verb("throw", function (e) { throw e; }), verb("return"), i[Symbol.iterator] = function () { return this; }, i;
    function verb(n, f) { i[n] = o[n] ? function (v) { return (p = !p) ? { value: __await(o[n](v)), done: n === "return" } : f ? f(v) : v; } : f; }
}

function __asyncValues(o) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var m = o[Symbol.asyncIterator], i;
    return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
    function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
    function settle(resolve, reject, d, v) { Promise.resolve(v).then(function(v) { resolve({ value: v, done: d }); }, reject); }
}

function __makeTemplateObject(cooked, raw) {
    if (Object.defineProperty) { Object.defineProperty(cooked, "raw", { value: raw }); } else { cooked.raw = raw; }
    return cooked;
};

function __importStar(mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result.default = mod;
    return result;
}

function __importDefault(mod) {
    return (mod && mod.__esModule) ? mod : { default: mod };
}

function __classPrivateFieldGet(receiver, privateMap) {
    if (!privateMap.has(receiver)) {
        throw new TypeError("attempted to get private field on non-instance");
    }
    return privateMap.get(receiver);
}

function __classPrivateFieldSet(receiver, privateMap, value) {
    if (!privateMap.has(receiver)) {
        throw new TypeError("attempted to set private field on non-instance");
    }
    privateMap.set(receiver, value);
    return value;
}


/***/ }),
/* 7 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "BatchResponseContent": () => (/* binding */ BatchResponseContent)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @class
 * Class that handles BatchResponseContent
 */
class BatchResponseContent {
    /**
     * @public
     * @constructor
     * Creates the BatchResponseContent instance
     * @param {BatchResponseBody} response - The response body returned for batch request from server
     * @returns An instance of a BatchResponseContent
     */
    constructor(response) {
        this.responses = new Map();
        this.update(response);
    }
    /**
     * @private
     * Creates native Response object from the json representation of it.
     * @param {KeyValuePairObject} responseJSON - The response json value
     * @returns The Response Object instance
     */
    createResponseObject(responseJSON) {
        const body = responseJSON.body;
        const options = {};
        options.status = responseJSON.status;
        if (responseJSON.statusText !== undefined) {
            options.statusText = responseJSON.statusText;
        }
        options.headers = responseJSON.headers;
        if (options.headers !== undefined && options.headers["Content-Type"] !== undefined) {
            if (options.headers["Content-Type"].split(";")[0] === "application/json") {
                const bodyString = JSON.stringify(body);
                return new Response(bodyString, options);
            }
        }
        return new Response(body, options);
    }
    /**
     * @public
     * Updates the Batch response content instance with given responses.
     * @param {BatchResponseBody} response - The response json representing batch response message
     * @returns Nothing
     */
    update(response) {
        this.nextLink = response["@odata.nextLink"];
        const responses = response.responses;
        for (let i = 0, l = responses.length; i < l; i++) {
            this.responses.set(responses[i].id, this.createResponseObject(responses[i]));
        }
    }
    /**
     * @public
     * To get the response of a request for a given request id
     * @param {string} requestId - The request id value
     * @returns The Response object instance for the particular request
     */
    getResponseById(requestId) {
        return this.responses.get(requestId);
    }
    /**
     * @public
     * To get all the responses of the batch request
     * @returns The Map of id and Response objects
     */
    getResponses() {
        return this.responses;
    }
    /**
     * @public
     * To get the iterator for the responses
     * @returns The Iterable generator for the response objects
     */
    *getResponsesIterator() {
        const iterator = this.responses.entries();
        let cur = iterator.next();
        while (!cur.done) {
            yield cur.value;
            cur = iterator.next();
        }
    }
}
//# sourceMappingURL=BatchResponseContent.js.map

/***/ }),
/* 8 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "AuthenticationHandler": () => (/* binding */ AuthenticationHandler)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(6);
/* harmony import */ var _MiddlewareControl__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(9);
/* harmony import */ var _MiddlewareUtil__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(10);
/* harmony import */ var _options_AuthenticationHandlerOptions__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(11);
/* harmony import */ var _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(12);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */





/**
 * @class
 * @implements Middleware
 * Class representing AuthenticationHandler
 */
class AuthenticationHandler {
    /**
     * @public
     * @constructor
     * Creates an instance of AuthenticationHandler
     * @param {AuthenticationProvider} authenticationProvider - The authentication provider for the authentication handler
     */
    constructor(authenticationProvider) {
        this.authenticationProvider = authenticationProvider;
    }
    /**
     * @public
     * @async
     * To execute the current middleware
     * @param {Context} context - The context object of the request
     * @returns A Promise that resolves to nothing
     */
    execute(context) {
        return tslib__WEBPACK_IMPORTED_MODULE_4__.__awaiter(this, void 0, void 0, function* () {
            try {
                let options;
                if (context.middlewareControl instanceof _MiddlewareControl__WEBPACK_IMPORTED_MODULE_0__.MiddlewareControl) {
                    options = context.middlewareControl.getMiddlewareOptions(_options_AuthenticationHandlerOptions__WEBPACK_IMPORTED_MODULE_2__.AuthenticationHandlerOptions);
                }
                let authenticationProvider;
                let authenticationProviderOptions;
                if (typeof options !== "undefined") {
                    authenticationProvider = options.authenticationProvider;
                    authenticationProviderOptions = options.authenticationProviderOptions;
                }
                if (typeof authenticationProvider === "undefined") {
                    authenticationProvider = this.authenticationProvider;
                }
                const token = yield authenticationProvider.getAccessToken(authenticationProviderOptions);
                const bearerKey = `Bearer ${token}`;
                (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_1__.appendRequestHeader)(context.request, context.options, AuthenticationHandler.AUTHORIZATION_HEADER, bearerKey);
                _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.TelemetryHandlerOptions.updateFeatureUsageFlag(context, _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.FeatureUsageFlag.AUTHENTICATION_HANDLER_ENABLED);
                return yield this.nextMiddleware.execute(context);
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * To set the next middleware in the chain
     * @param {Middleware} next - The middleware instance
     * @returns Nothing
     */
    setNext(next) {
        this.nextMiddleware = next;
    }
}
/**
 * @private
 * A member representing the authorization header name
 */
AuthenticationHandler.AUTHORIZATION_HEADER = "Authorization";
//# sourceMappingURL=AuthenticationHandler.js.map

/***/ }),
/* 9 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "MiddlewareControl": () => (/* binding */ MiddlewareControl)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @class
 * Class representing MiddlewareControl
 */
class MiddlewareControl {
    /**
     * @public
     * @constructor
     * Creates an instance of MiddlewareControl
     * @param {MiddlewareOptions[]} [middlewareOptions = []] - The array of middlewareOptions
     * @returns The instance of MiddlewareControl
     */
    constructor(middlewareOptions = []) {
        // tslint:disable-next-line:ban-types
        this.middlewareOptions = new Map();
        for (const option of middlewareOptions) {
            const fn = option.constructor;
            this.middlewareOptions.set(fn, option);
        }
    }
    /**
     * @public
     * To get the middleware option using the class of the option
     * @param {Function} fn - The class of the strongly typed option class
     * @returns The middleware option
     * @example
     * // if you wanted to return the middleware option associated with this class (MiddlewareControl)
     * // call this function like this:
     * getMiddlewareOptions(MiddlewareControl)
     */
    // tslint:disable-next-line:ban-types
    getMiddlewareOptions(fn) {
        return this.middlewareOptions.get(fn);
    }
    /**
     * @public
     * To set the middleware options using the class of the option
     * @param {Function} fn - The class of the strongly typed option class
     * @param {MiddlewareOptions} option - The strongly typed middleware option
     * @returns nothing
     */
    // tslint:disable-next-line:ban-types
    setMiddlewareOptions(fn, option) {
        this.middlewareOptions.set(fn, option);
    }
}
//# sourceMappingURL=MiddlewareControl.js.map

/***/ }),
/* 10 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "generateUUID": () => (/* binding */ generateUUID),
/* harmony export */   "getRequestHeader": () => (/* binding */ getRequestHeader),
/* harmony export */   "setRequestHeader": () => (/* binding */ setRequestHeader),
/* harmony export */   "appendRequestHeader": () => (/* binding */ appendRequestHeader),
/* harmony export */   "cloneRequestWithNewUrl": () => (/* binding */ cloneRequestWithNewUrl)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @constant
 * To generate the UUID
 * @returns The UUID string
 */
const generateUUID = () => {
    let uuid = "";
    for (let j = 0; j < 32; j++) {
        if (j === 8 || j === 12 || j === 16 || j === 20) {
            uuid += "-";
        }
        uuid += Math.floor(Math.random() * 16).toString(16);
    }
    return uuid;
};
/**
 * @constant
 * To get the request header from the request
 * @param {RequestInfo} request - The request object or the url string
 * @param {FetchOptions|undefined} options - The request options object
 * @param {string} key - The header key string
 * @returns A header value for the given key from the request
 */
const getRequestHeader = (request, options, key) => {
    let value = null;
    if (typeof Request !== "undefined" && request instanceof Request) {
        value = request.headers.get(key);
    }
    else if (typeof options !== "undefined" && options.headers !== undefined) {
        if (typeof Headers !== "undefined" && options.headers instanceof Headers) {
            value = options.headers.get(key);
        }
        else if (options.headers instanceof Array) {
            const headers = options.headers;
            for (let i = 0, l = headers.length; i < l; i++) {
                if (headers[i][0] === key) {
                    value = headers[i][1];
                    break;
                }
            }
        }
        else if (options.headers[key] !== undefined) {
            value = options.headers[key];
        }
    }
    return value;
};
/**
 * @constant
 * To set the header value to the given request
 * @param {RequestInfo} request - The request object or the url string
 * @param {FetchOptions|undefined} options - The request options object
 * @param {string} key - The header key string
 * @param {string } value - The header value string
 * @returns Nothing
 */
const setRequestHeader = (request, options, key, value) => {
    if (typeof Request !== "undefined" && request instanceof Request) {
        request.headers.set(key, value);
    }
    else if (typeof options !== "undefined") {
        if (options.headers === undefined) {
            options.headers = new Headers({
                [key]: value,
            });
        }
        else {
            if (typeof Headers !== "undefined" && options.headers instanceof Headers) {
                options.headers.set(key, value);
            }
            else if (options.headers instanceof Array) {
                let i = 0;
                const l = options.headers.length;
                for (; i < l; i++) {
                    const header = options.headers[i];
                    if (header[0] === key) {
                        header[1] = value;
                        break;
                    }
                }
                if (i === l) {
                    options.headers.push([key, value]);
                }
            }
            else {
                Object.assign(options.headers, { [key]: value });
            }
        }
    }
};
/**
 * @constant
 * To append the header value to the given request
 * @param {RequestInfo} request - The request object or the url string
 * @param {FetchOptions|undefined} options - The request options object
 * @param {string} key - The header key string
 * @param {string } value - The header value string
 * @returns Nothing
 */
const appendRequestHeader = (request, options, key, value) => {
    if (typeof Request !== "undefined" && request instanceof Request) {
        request.headers.append(key, value);
    }
    else if (typeof options !== "undefined") {
        if (options.headers === undefined) {
            options.headers = new Headers({
                [key]: value,
            });
        }
        else {
            if (typeof Headers !== "undefined" && options.headers instanceof Headers) {
                options.headers.append(key, value);
            }
            else if (options.headers instanceof Array) {
                options.headers.push([key, value]);
            }
            else if (options.headers === undefined) {
                options.headers = { [key]: value };
            }
            else if (options.headers[key] === undefined) {
                options.headers[key] = value;
            }
            else {
                options.headers[key] += `, ${value}`;
            }
        }
    }
};
/**
 * @constant
 * To clone the request with the new url
 * @param {string} url - The new url string
 * @param {Request} request - The request object
 * @returns A promise that resolves to request object
 */
const cloneRequestWithNewUrl = (newUrl, request) => tslib__WEBPACK_IMPORTED_MODULE_0__.__awaiter(undefined, void 0, void 0, function* () {
    const body = request.headers.get("Content-Type") ? yield request.blob() : yield Promise.resolve(undefined);
    const { method, headers, referrer, referrerPolicy, mode, credentials, cache, redirect, integrity, keepalive, signal } = request;
    return new Request(newUrl, { method, headers, body, referrer, referrerPolicy, mode, credentials, cache, redirect, integrity, keepalive, signal });
});
//# sourceMappingURL=MiddlewareUtil.js.map

/***/ }),
/* 11 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "AuthenticationHandlerOptions": () => (/* binding */ AuthenticationHandlerOptions)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @class
 * @implements MiddlewareOptions
 * Class representing AuthenticationHandlerOptions
 */
class AuthenticationHandlerOptions {
    /**
     * @public
     * @constructor
     * To create an instance of AuthenticationHandlerOptions
     * @param {AuthenticationProvider} [authenticationProvider] - The authentication provider instance
     * @param {AuthenticationProviderOptions} [authenticationProviderOptions] - The authentication provider options instance
     * @returns An instance of AuthenticationHandlerOptions
     */
    constructor(authenticationProvider, authenticationProviderOptions) {
        this.authenticationProvider = authenticationProvider;
        this.authenticationProviderOptions = authenticationProviderOptions;
    }
}
//# sourceMappingURL=AuthenticationHandlerOptions.js.map

/***/ }),
/* 12 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "FeatureUsageFlag": () => (/* binding */ FeatureUsageFlag),
/* harmony export */   "TelemetryHandlerOptions": () => (/* binding */ TelemetryHandlerOptions)
/* harmony export */ });
/* harmony import */ var _MiddlewareControl__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(9);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @enum
 * @property {number} NONE - The hexadecimal flag value for nothing enabled
 * @property {number} REDIRECT_HANDLER_ENABLED - The hexadecimal flag value for redirect handler enabled
 * @property {number} RETRY_HANDLER_ENABLED - The hexadecimal flag value for retry handler enabled
 * @property {number} AUTHENTICATION_HANDLER_ENABLED - The hexadecimal flag value for the authentication handler enabled
 */
var FeatureUsageFlag;
(function (FeatureUsageFlag) {
    FeatureUsageFlag[FeatureUsageFlag["NONE"] = 0] = "NONE";
    FeatureUsageFlag[FeatureUsageFlag["REDIRECT_HANDLER_ENABLED"] = 1] = "REDIRECT_HANDLER_ENABLED";
    FeatureUsageFlag[FeatureUsageFlag["RETRY_HANDLER_ENABLED"] = 2] = "RETRY_HANDLER_ENABLED";
    FeatureUsageFlag[FeatureUsageFlag["AUTHENTICATION_HANDLER_ENABLED"] = 4] = "AUTHENTICATION_HANDLER_ENABLED";
})(FeatureUsageFlag || (FeatureUsageFlag = {}));
/**
 * @class
 * @implements MiddlewareOptions
 * Class for TelemetryHandlerOptions
 */
class TelemetryHandlerOptions {
    constructor() {
        /**
         * @private
         * A member to hold the OR of feature usage flags
         */
        this.featureUsage = FeatureUsageFlag.NONE;
    }
    /**
     * @public
     * @static
     * To update the feature usage in the context object
     * @param {Context} context - The request context object containing middleware options
     * @param {FeatureUsageFlag} flag - The flag value
     * @returns nothing
     */
    static updateFeatureUsageFlag(context, flag) {
        let options;
        if (context.middlewareControl instanceof _MiddlewareControl__WEBPACK_IMPORTED_MODULE_0__.MiddlewareControl) {
            options = context.middlewareControl.getMiddlewareOptions(TelemetryHandlerOptions);
        }
        else {
            context.middlewareControl = new _MiddlewareControl__WEBPACK_IMPORTED_MODULE_0__.MiddlewareControl();
        }
        if (typeof options === "undefined") {
            options = new TelemetryHandlerOptions();
            context.middlewareControl.setMiddlewareOptions(TelemetryHandlerOptions, options);
        }
        options.setFeatureUsage(flag);
    }
    /**
     * @private
     * To set the feature usage flag
     * @param {FeatureUsageFlag} flag - The flag value
     * @returns nothing
     */
    setFeatureUsage(flag) {
        /* tslint:disable: no-bitwise */
        this.featureUsage = this.featureUsage | flag;
        /* tslint:enable: no-bitwise */
    }
    /**
     * @public
     * To get the feature usage
     * @returns A feature usage flag as hexadecimal string
     */
    getFeatureUsage() {
        return this.featureUsage.toString(16);
    }
}
//# sourceMappingURL=TelemetryHandlerOptions.js.map

/***/ }),
/* 13 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "HTTPMessageHandler": () => (/* binding */ HTTPMessageHandler)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @class
 * @implements Middleware
 * Class for HTTPMessageHandler
 */
class HTTPMessageHandler {
    /**
     * @public
     * @async
     * To execute the current middleware
     * @param {Context} context - The request context object
     * @returns A promise that resolves to nothing
     */
    execute(context) {
        return tslib__WEBPACK_IMPORTED_MODULE_0__.__awaiter(this, void 0, void 0, function* () {
            try {
                context.response = yield fetch(context.request, context.options);
                return;
            }
            catch (error) {
                throw error;
            }
        });
    }
}
//# sourceMappingURL=HTTPMessageHandler.js.map

/***/ }),
/* 14 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "RetryHandler": () => (/* binding */ RetryHandler)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(6);
/* harmony import */ var _RequestMethod__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(5);
/* harmony import */ var _MiddlewareControl__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(9);
/* harmony import */ var _MiddlewareUtil__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(10);
/* harmony import */ var _options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(15);
/* harmony import */ var _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(12);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */






/**
 * @class
 * @implements Middleware
 * Class for RetryHandler
 */
class RetryHandler {
    /**
     * @public
     * @constructor
     * To create an instance of RetryHandler
     * @param {RetryHandlerOptions} [options = new RetryHandlerOptions()] - The retry handler options value
     * @returns An instance of RetryHandler
     */
    constructor(options = new _options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.RetryHandlerOptions()) {
        this.options = options;
    }
    /**
     *
     * @private
     * To check whether the response has the retry status code
     * @param {Response} response - The response object
     * @returns Whether the response has retry status code or not
     */
    isRetry(response) {
        return RetryHandler.RETRY_STATUS_CODES.indexOf(response.status) !== -1;
    }
    /**
     * @private
     * To check whether the payload is buffered or not
     * @param {RequestInfo} request - The url string or the request object value
     * @param {FetchOptions} options - The options of a request
     * @returns Whether the payload is buffered or not
     */
    isBuffered(request, options) {
        const method = typeof request === "string" ? options.method : request.method;
        const isPutPatchOrPost = method === _RequestMethod__WEBPACK_IMPORTED_MODULE_0__.RequestMethod.PUT || method === _RequestMethod__WEBPACK_IMPORTED_MODULE_0__.RequestMethod.PATCH || method === _RequestMethod__WEBPACK_IMPORTED_MODULE_0__.RequestMethod.POST;
        if (isPutPatchOrPost) {
            const isStream = (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_2__.getRequestHeader)(request, options, "Content-Type") === "application/octet-stream";
            if (isStream) {
                return false;
            }
        }
        return true;
    }
    /**
     * @private
     * To get the delay for a retry
     * @param {Response} response - The response object
     * @param {number} retryAttempts - The current attempt count
     * @param {number} delay - The delay value in seconds
     * @returns A delay for a retry
     */
    getDelay(response, retryAttempts, delay) {
        const getRandomness = () => Number(Math.random().toFixed(3));
        const retryAfter = response.headers !== undefined ? response.headers.get(RetryHandler.RETRY_AFTER_HEADER) : null;
        let newDelay;
        if (retryAfter !== null) {
            // tslint:disable: prefer-conditional-expression
            if (Number.isNaN(Number(retryAfter))) {
                newDelay = Math.round((new Date(retryAfter).getTime() - Date.now()) / 1000);
            }
            else {
                newDelay = Number(retryAfter);
            }
            // tslint:enable: prefer-conditional-expression
        }
        else {
            // Adding randomness to avoid retrying at a same
            newDelay = retryAttempts >= 2 ? this.getExponentialBackOffTime(retryAttempts) + delay + getRandomness() : delay + getRandomness();
        }
        return Math.min(newDelay, this.options.getMaxDelay() + getRandomness());
    }
    /**
     * @private
     * To get an exponential back off value
     * @param {number} attempts - The current attempt count
     * @returns An exponential back off value
     */
    getExponentialBackOffTime(attempts) {
        return Math.round((1 / 2) * (Math.pow(2, attempts) - 1));
    }
    /**
     * @private
     * @async
     * To add delay for the execution
     * @param {number} delaySeconds - The delay value in seconds
     * @returns Nothing
     */
    sleep(delaySeconds) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            const delayMilliseconds = delaySeconds * 1000;
            return new Promise((resolve) => setTimeout(resolve, delayMilliseconds));
        });
    }
    getOptions(context) {
        let options;
        if (context.middlewareControl instanceof _MiddlewareControl__WEBPACK_IMPORTED_MODULE_1__.MiddlewareControl) {
            options = context.middlewareControl.getMiddlewareOptions(this.options.constructor);
        }
        if (typeof options === "undefined") {
            options = Object.assign(new _options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.RetryHandlerOptions(), this.options);
        }
        return options;
    }
    /**
     * @private
     * @async
     * To execute the middleware with retries
     * @param {Context} context - The context object
     * @param {number} retryAttempts - The current attempt count
     * @param {RetryHandlerOptions} options - The retry middleware options instance
     * @returns A Promise that resolves to nothing
     */
    executeWithRetry(context, retryAttempts, options) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            try {
                yield this.nextMiddleware.execute(context);
                if (retryAttempts < options.maxRetries && this.isRetry(context.response) && this.isBuffered(context.request, context.options) && options.shouldRetry(options.delay, retryAttempts, context.request, context.options, context.response)) {
                    ++retryAttempts;
                    (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_2__.setRequestHeader)(context.request, context.options, RetryHandler.RETRY_ATTEMPT_HEADER, retryAttempts.toString());
                    const delay = this.getDelay(context.response, retryAttempts, options.delay);
                    yield this.sleep(delay);
                    return yield this.executeWithRetry(context, retryAttempts, options);
                }
                else {
                    return;
                }
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * To execute the current middleware
     * @param {Context} context - The context object of the request
     * @returns A Promise that resolves to nothing
     */
    execute(context) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            try {
                const retryAttempts = 0;
                const options = this.getOptions(context);
                _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__.TelemetryHandlerOptions.updateFeatureUsageFlag(context, _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__.FeatureUsageFlag.RETRY_HANDLER_ENABLED);
                return yield this.executeWithRetry(context, retryAttempts, options);
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * To set the next middleware in the chain
     * @param {Middleware} next - The middleware instance
     * @returns Nothing
     */
    setNext(next) {
        this.nextMiddleware = next;
    }
}
/**
 * @private
 * @static
 * A list of status codes that needs to be retried
 */
RetryHandler.RETRY_STATUS_CODES = [
    429,
    503,
    504,
];
/**
 * @private
 * @static
 * A member holding the name of retry attempt header
 */
RetryHandler.RETRY_ATTEMPT_HEADER = "Retry-Attempt";
/**
 * @private
 * @static
 * A member holding the name of retry after header
 */
RetryHandler.RETRY_AFTER_HEADER = "Retry-After";
//# sourceMappingURL=RetryHandler.js.map

/***/ }),
/* 15 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "RetryHandlerOptions": () => (/* binding */ RetryHandlerOptions)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @class
 * @implements MiddlewareOptions
 * Class for RetryHandlerOptions
 */
class RetryHandlerOptions {
    /**
     * @public
     * @constructor
     * To create an instance of RetryHandlerOptions
     * @param {number} [delay = RetryHandlerOptions.DEFAULT_DELAY] - The delay value in seconds
     * @param {number} [maxRetries = RetryHandlerOptions.DEFAULT_MAX_RETRIES] - The maxRetries value
     * @param {ShouldRetry} [shouldRetry = RetryHandlerOptions.DEFAULT_SHOULD_RETRY] - The shouldRetry callback function
     * @returns An instance of RetryHandlerOptions
     */
    constructor(delay = RetryHandlerOptions.DEFAULT_DELAY, maxRetries = RetryHandlerOptions.DEFAULT_MAX_RETRIES, shouldRetry = RetryHandlerOptions.DEFAULT_SHOULD_RETRY) {
        if (delay > RetryHandlerOptions.MAX_DELAY && maxRetries > RetryHandlerOptions.MAX_MAX_RETRIES) {
            const error = new Error(`Delay and MaxRetries should not be more than ${RetryHandlerOptions.MAX_DELAY} and ${RetryHandlerOptions.MAX_MAX_RETRIES}`);
            error.name = "MaxLimitExceeded";
            throw error;
        }
        else if (delay > RetryHandlerOptions.MAX_DELAY) {
            const error = new Error(`Delay should not be more than ${RetryHandlerOptions.MAX_DELAY}`);
            error.name = "MaxLimitExceeded";
            throw error;
        }
        else if (maxRetries > RetryHandlerOptions.MAX_MAX_RETRIES) {
            const error = new Error(`MaxRetries should not be more than ${RetryHandlerOptions.MAX_MAX_RETRIES}`);
            error.name = "MaxLimitExceeded";
            throw error;
        }
        else if (delay < 0 && maxRetries < 0) {
            const error = new Error(`Delay and MaxRetries should not be negative`);
            error.name = "MinExpectationNotMet";
            throw error;
        }
        else if (delay < 0) {
            const error = new Error(`Delay should not be negative`);
            error.name = "MinExpectationNotMet";
            throw error;
        }
        else if (maxRetries < 0) {
            const error = new Error(`MaxRetries should not be negative`);
            error.name = "MinExpectationNotMet";
            throw error;
        }
        this.delay = Math.min(delay, RetryHandlerOptions.MAX_DELAY);
        this.maxRetries = Math.min(maxRetries, RetryHandlerOptions.MAX_MAX_RETRIES);
        this.shouldRetry = shouldRetry;
    }
    /**
     * @public
     * To get the maximum delay
     * @returns A maximum delay
     */
    getMaxDelay() {
        return RetryHandlerOptions.MAX_DELAY;
    }
}
/**
 * @private
 * @static
 * A member holding default delay value in seconds
 */
RetryHandlerOptions.DEFAULT_DELAY = 3;
/**
 * @private
 * @static
 * A member holding default maxRetries value
 */
RetryHandlerOptions.DEFAULT_MAX_RETRIES = 3;
/**
 * @private
 * @static
 * A member holding maximum delay value in seconds
 */
RetryHandlerOptions.MAX_DELAY = 180;
/**
 * @private
 * @static
 * A member holding maximum maxRetries value
 */
RetryHandlerOptions.MAX_MAX_RETRIES = 10;
/**
 * @private
 * A member holding default shouldRetry callback
 */
RetryHandlerOptions.DEFAULT_SHOULD_RETRY = () => true;
//# sourceMappingURL=RetryHandlerOptions.js.map

/***/ }),
/* 16 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "RedirectHandler": () => (/* binding */ RedirectHandler)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(6);
/* harmony import */ var _RequestMethod__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(5);
/* harmony import */ var _MiddlewareControl__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(9);
/* harmony import */ var _MiddlewareUtil__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(10);
/* harmony import */ var _options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(17);
/* harmony import */ var _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(12);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */






/**
 * @class
 * Class
 * @implements Middleware
 * Class representing RedirectHandler
 */
class RedirectHandler {
    /**
     * @public
     * @constructor
     * To create an instance of RedirectHandler
     * @param {RedirectHandlerOptions} [options = new RedirectHandlerOptions()] - The redirect handler options instance
     * @returns An instance of RedirectHandler
     */
    constructor(options = new _options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.RedirectHandlerOptions()) {
        this.options = options;
    }
    /**
     * @private
     * To check whether the response has the redirect status code or not
     * @param {Response} response - The response object
     * @returns A boolean representing whether the response contains the redirect status code or not
     */
    isRedirect(response) {
        return RedirectHandler.REDIRECT_STATUS_CODES.indexOf(response.status) !== -1;
    }
    /**
     * @private
     * To check whether the response has location header or not
     * @param {Response} response - The response object
     * @returns A boolean representing the whether the response has location header or not
     */
    hasLocationHeader(response) {
        return response.headers.has(RedirectHandler.LOCATION_HEADER);
    }
    /**
     * @private
     * To get the redirect url from location header in response object
     * @param {Response} response - The response object
     * @returns A redirect url from location header
     */
    getLocationHeader(response) {
        return response.headers.get(RedirectHandler.LOCATION_HEADER);
    }
    /**
     * @private
     * To check whether the given url is a relative url or not
     * @param {string} url - The url string value
     * @returns A boolean representing whether the given url is a relative url or not
     */
    isRelativeURL(url) {
        return url.indexOf("://") === -1;
    }
    /**
     * @private
     * To check whether the authorization header in the request should be dropped for consequent redirected requests
     * @param {string} requestUrl - The request url value
     * @param {string} redirectUrl - The redirect url value
     * @returns A boolean representing whether the authorization header in the request should be dropped for consequent redirected requests
     */
    shouldDropAuthorizationHeader(requestUrl, redirectUrl) {
        const schemeHostRegex = /^[A-Za-z].+?:\/\/.+?(?=\/|$)/;
        const requestMatches = schemeHostRegex.exec(requestUrl);
        let requestAuthority;
        let redirectAuthority;
        if (requestMatches !== null) {
            requestAuthority = requestMatches[0];
        }
        const redirectMatches = schemeHostRegex.exec(redirectUrl);
        if (redirectMatches !== null) {
            redirectAuthority = redirectMatches[0];
        }
        return typeof requestAuthority !== "undefined" && typeof redirectAuthority !== "undefined" && requestAuthority !== redirectAuthority;
    }
    /**
     * @private
     * @async
     * To update a request url with the redirect url
     * @param {string} redirectUrl - The redirect url value
     * @param {Context} context - The context object value
     * @returns Nothing
     */
    updateRequestUrl(redirectUrl, context) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            context.request = typeof context.request === "string" ? redirectUrl : yield (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_2__.cloneRequestWithNewUrl)(redirectUrl, context.request);
        });
    }
    /**
     * @private
     * To get the options for execution of the middleware
     * @param {Context} context - The context object
     * @returns A options for middleware execution
     */
    getOptions(context) {
        let options;
        if (context.middlewareControl instanceof _MiddlewareControl__WEBPACK_IMPORTED_MODULE_1__.MiddlewareControl) {
            options = context.middlewareControl.getMiddlewareOptions(_options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.RedirectHandlerOptions);
        }
        if (typeof options === "undefined") {
            options = Object.assign(new _options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.RedirectHandlerOptions(), this.options);
        }
        return options;
    }
    /**
     * @private
     * @async
     * To execute the next middleware and to handle in case of redirect response returned by the server
     * @param {Context} context - The context object
     * @param {number} redirectCount - The redirect count value
     * @param {RedirectHandlerOptions} options - The redirect handler options instance
     * @returns A promise that resolves to nothing
     */
    executeWithRedirect(context, redirectCount, options) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            try {
                yield this.nextMiddleware.execute(context);
                const response = context.response;
                if (redirectCount < options.maxRedirects && this.isRedirect(response) && this.hasLocationHeader(response) && options.shouldRedirect(response)) {
                    ++redirectCount;
                    if (response.status === RedirectHandler.STATUS_CODE_SEE_OTHER) {
                        context.options.method = _RequestMethod__WEBPACK_IMPORTED_MODULE_0__.RequestMethod.GET;
                        delete context.options.body;
                    }
                    else {
                        const redirectUrl = this.getLocationHeader(response);
                        if (!this.isRelativeURL(redirectUrl) && this.shouldDropAuthorizationHeader(response.url, redirectUrl)) {
                            delete context.options.headers[RedirectHandler.AUTHORIZATION_HEADER];
                        }
                        yield this.updateRequestUrl(redirectUrl, context);
                    }
                    yield this.executeWithRedirect(context, redirectCount, options);
                }
                else {
                    return;
                }
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * To execute the current middleware
     * @param {Context} context - The context object of the request
     * @returns A Promise that resolves to nothing
     */
    execute(context) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            try {
                const redirectCount = 0;
                const options = this.getOptions(context);
                context.options.redirect = RedirectHandler.MANUAL_REDIRECT;
                _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__.TelemetryHandlerOptions.updateFeatureUsageFlag(context, _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__.FeatureUsageFlag.REDIRECT_HANDLER_ENABLED);
                return yield this.executeWithRedirect(context, redirectCount, options);
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * To set the next middleware in the chain
     * @param {Middleware} next - The middleware instance
     * @returns Nothing
     */
    setNext(next) {
        this.nextMiddleware = next;
    }
}
/**
 * @private
 * @static
 * A member holding the array of redirect status codes
 */
RedirectHandler.REDIRECT_STATUS_CODES = [
    301,
    302,
    303,
    307,
    308,
];
/**
 * @private
 * @static
 * A member holding SeeOther status code
 */
RedirectHandler.STATUS_CODE_SEE_OTHER = 303;
/**
 * @private
 * @static
 * A member holding the name of the location header
 */
RedirectHandler.LOCATION_HEADER = "Location";
/**
 * @private
 * @static
 * A member representing the authorization header name
 */
RedirectHandler.AUTHORIZATION_HEADER = "Authorization";
/**
 * @private
 * @static
 * A member holding the manual redirect value
 */
RedirectHandler.MANUAL_REDIRECT = "manual";
//# sourceMappingURL=RedirectHandler.js.map

/***/ }),
/* 17 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "RedirectHandlerOptions": () => (/* binding */ RedirectHandlerOptions)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @class
 * @implements MiddlewareOptions
 * A class representing RedirectHandlerOptions
 */
class RedirectHandlerOptions {
    /**
     * @public
     * @constructor
     * To create an instance of RedirectHandlerOptions
     * @param {number} [maxRedirects = RedirectHandlerOptions.DEFAULT_MAX_REDIRECTS] - The max redirects value
     * @param {ShouldRedirect} [shouldRedirect = RedirectHandlerOptions.DEFAULT_SHOULD_RETRY] - The should redirect callback
     * @returns An instance of RedirectHandlerOptions
     */
    constructor(maxRedirects = RedirectHandlerOptions.DEFAULT_MAX_REDIRECTS, shouldRedirect = RedirectHandlerOptions.DEFAULT_SHOULD_RETRY) {
        if (maxRedirects > RedirectHandlerOptions.MAX_MAX_REDIRECTS) {
            const error = new Error(`MaxRedirects should not be more than ${RedirectHandlerOptions.MAX_MAX_REDIRECTS}`);
            error.name = "MaxLimitExceeded";
            throw error;
        }
        if (maxRedirects < 0) {
            const error = new Error(`MaxRedirects should not be negative`);
            error.name = "MinExpectationNotMet";
            throw error;
        }
        this.maxRedirects = maxRedirects;
        this.shouldRedirect = shouldRedirect;
    }
}
/**
 * @private
 * @static
 * A member holding default max redirects value
 */
RedirectHandlerOptions.DEFAULT_MAX_REDIRECTS = 5;
/**
 * @private
 * @static
 * A member holding maximum max redirects value
 */
RedirectHandlerOptions.MAX_MAX_REDIRECTS = 20;
/**
 * @private
 * A member holding default shouldRedirect callback
 */
RedirectHandlerOptions.DEFAULT_SHOULD_RETRY = () => true;
//# sourceMappingURL=RedirectHandlerOptions.js.map

/***/ }),
/* 18 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "TelemetryHandler": () => (/* binding */ TelemetryHandler)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(6);
/* harmony import */ var _GraphRequestUtil__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(19);
/* harmony import */ var _Version__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(21);
/* harmony import */ var _MiddlewareControl__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(9);
/* harmony import */ var _MiddlewareUtil__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(10);
/* harmony import */ var _options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(12);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @module TelemetryHandler
 */





/**
 * @class
 * @implements Middleware
 * Class for TelemetryHandler
 */
class TelemetryHandler {
    /**
     * @public
     * @async
     * To execute the current middleware
     * @param {Context} context - The context object of the request
     * @returns A Promise that resolves to nothing
     */
    execute(context) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            try {
                const url = typeof context.request === "string" ? context.request : context.request.url;
                if ((0,_GraphRequestUtil__WEBPACK_IMPORTED_MODULE_0__.isGraphURL)(url)) {
                    // Add telemetry only if the request url is a Graph URL.
                    // Errors are reported as in issue #265 if headers are present when redirecting to a non Graph URL
                    let clientRequestId = (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_3__.getRequestHeader)(context.request, context.options, TelemetryHandler.CLIENT_REQUEST_ID_HEADER);
                    if (!clientRequestId) {
                        clientRequestId = (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_3__.generateUUID)();
                        (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_3__.setRequestHeader)(context.request, context.options, TelemetryHandler.CLIENT_REQUEST_ID_HEADER, clientRequestId);
                    }
                    let sdkVersionValue = `${TelemetryHandler.PRODUCT_NAME}/${_Version__WEBPACK_IMPORTED_MODULE_1__.PACKAGE_VERSION}`;
                    let options;
                    if (context.middlewareControl instanceof _MiddlewareControl__WEBPACK_IMPORTED_MODULE_2__.MiddlewareControl) {
                        options = context.middlewareControl.getMiddlewareOptions(_options_TelemetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__.TelemetryHandlerOptions);
                    }
                    if (options) {
                        const featureUsage = options.getFeatureUsage();
                        sdkVersionValue += ` (${TelemetryHandler.FEATURE_USAGE_STRING}=${featureUsage})`;
                    }
                    (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_3__.appendRequestHeader)(context.request, context.options, TelemetryHandler.SDK_VERSION_HEADER, sdkVersionValue);
                }
                else {
                    // Remove telemetry headers if present during redirection.
                    delete context.options.headers[TelemetryHandler.CLIENT_REQUEST_ID_HEADER];
                    delete context.options.headers[TelemetryHandler.SDK_VERSION_HEADER];
                }
                return yield this.nextMiddleware.execute(context);
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * To set the next middleware in the chain
     * @param {Middleware} next - The middleware instance
     * @returns Nothing
     */
    setNext(next) {
        this.nextMiddleware = next;
    }
}
/**
 * @private
 * @static
 * A member holding the name of the client request id header
 */
TelemetryHandler.CLIENT_REQUEST_ID_HEADER = "client-request-id";
/**
 * @private
 * @static
 * A member holding the name of the sdk version header
 */
TelemetryHandler.SDK_VERSION_HEADER = "SdkVersion";
/**
 * @private
 * @static
 * A member holding the language prefix for the sdk version header value
 */
TelemetryHandler.PRODUCT_NAME = "graph-js";
/**
 * @private
 * @static
 * A member holding the key for the feature usage metrics
 */
TelemetryHandler.FEATURE_USAGE_STRING = "featureUsage";
//# sourceMappingURL=TelemetryHandler.js.map

/***/ }),
/* 19 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "oDataQueryNames": () => (/* binding */ oDataQueryNames),
/* harmony export */   "urlJoin": () => (/* binding */ urlJoin),
/* harmony export */   "serializeContent": () => (/* binding */ serializeContent),
/* harmony export */   "isGraphURL": () => (/* binding */ isGraphURL)
/* harmony export */ });
/* harmony import */ var _Constants__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(20);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module GraphRequestUtil
 */

/**
 * To hold list of OData query params
 */
const oDataQueryNames = ["$select", "$expand", "$orderby", "$filter", "$top", "$skip", "$skipToken", "$count"];
/**
 * To construct the URL by appending the segments with "/"
 * @param {string[]} urlSegments - The array of strings
 * @returns The constructed URL string
 */
const urlJoin = (urlSegments) => {
    const removePostSlash = (s) => s.replace(/\/+$/, "");
    const removePreSlash = (s) => s.replace(/^\/+/, "");
    const joiner = (pre, cur) => [removePostSlash(pre), removePreSlash(cur)].join("/");
    const parts = Array.prototype.slice.call(urlSegments);
    return parts.reduce(joiner);
};
/**
 * Serializes the content
 * @param {any} content - The content value that needs to be serialized
 * @returns The serialized content
 *
 * Note:
 * This conversion is required due to the following reasons:
 * Body parameter of Request method of isomorphic-fetch only accepts Blob, ArrayBuffer, FormData, TypedArrays string.
 * Node.js platform does not support Blob, FormData. Javascript File object inherits from Blob so it is also not supported in node. Therefore content of type Blob, File, FormData will only come from browsers.
 * Parallel to ArrayBuffer in javascript, node provides Buffer interface. Node's Buffer is able to send the arbitrary binary data to the server successfully for both Browser and Node platform. Whereas sending binary data via ArrayBuffer or TypedArrays was only possible using Browser. To support both Node and Browser, `serializeContent` converts TypedArrays or ArrayBuffer to `Node Buffer`.
 * If the data received is in JSON format, `serializeContent` converts the JSON to string.
 */
const serializeContent = (content) => {
    const className = content && content.constructor && content.constructor.name;
    if (className === "Buffer" || className === "Blob" || className === "File" || className === "FormData" || typeof content === "string") {
        return content;
    }
    if (className === "ArrayBuffer") {
        content = Buffer.from(content);
    }
    else if (className === "Int8Array" || className === "Int16Array" || className === "Int32Array" || className === "Uint8Array" || className === "Uint16Array" || className === "Uint32Array" || className === "Uint8ClampedArray" || className === "Float32Array" || className === "Float64Array" || className === "DataView") {
        content = Buffer.from(content.buffer);
    }
    else {
        try {
            content = JSON.stringify(content);
        }
        catch (error) {
            throw new Error("Unable to stringify the content");
        }
    }
    return content;
};
/**
 * Checks if the url is one of the service root endpoints for Microsoft Graph and Graph Explorer.
 * @param {string} url - The url to be verified
 * @returns {boolean} - Returns true if the url is a Graph URL
 */
const isGraphURL = (url) => {
    // Valid Graph URL pattern - https://graph.microsoft.com/{version}/{resource}?{query-parameters}
    // Valid Graph URL example - https://graph.microsoft.com/v1.0/
    url = url.toLowerCase();
    if (url.indexOf("https://") !== -1) {
        url = url.replace("https://", "");
        // Find where the host ends
        const startofPortNoPos = url.indexOf(":");
        const endOfHostStrPos = url.indexOf("/");
        let hostName = "";
        if (endOfHostStrPos !== -1) {
            if (startofPortNoPos !== -1 && startofPortNoPos < endOfHostStrPos) {
                hostName = url.substring(0, startofPortNoPos);
                return _Constants__WEBPACK_IMPORTED_MODULE_0__.GRAPH_URLS.has(hostName);
            }
            // Parse out the host
            hostName = url.substring(0, endOfHostStrPos);
            return _Constants__WEBPACK_IMPORTED_MODULE_0__.GRAPH_URLS.has(hostName);
        }
    }
    return false;
};
//# sourceMappingURL=GraphRequestUtil.js.map

/***/ }),
/* 20 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "GRAPH_API_VERSION": () => (/* binding */ GRAPH_API_VERSION),
/* harmony export */   "GRAPH_BASE_URL": () => (/* binding */ GRAPH_BASE_URL),
/* harmony export */   "GRAPH_URLS": () => (/* binding */ GRAPH_URLS)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module Constants
 */
/**
 * @constant
 * A Default API endpoint version for a request
 */
const GRAPH_API_VERSION = "v1.0";
/**
 * @constant
 * A Default base url for a request
 */
const GRAPH_BASE_URL = "https://graph.microsoft.com/";
/**
 * To hold list of the service root endpoints for Microsoft Graph and Graph Explorer for each national cloud.
 * Set(iterable:Object) is not supported in Internet Explorer. The consumer is recommended to use a suitable polyfill.
 */
const GRAPH_URLS = new Set(["graph.microsoft.com", "graph.microsoft.us", "dod-graph.microsoft.us", "graph.microsoft.de", "microsoftgraph.chinacloudapi.cn"]);
//# sourceMappingURL=Constants.js.map

/***/ }),
/* 21 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "PACKAGE_VERSION": () => (/* binding */ PACKAGE_VERSION)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD
/**
 * @module Version
 */
const PACKAGE_VERSION = "2.2.1";
//# sourceMappingURL=Version.js.map

/***/ }),
/* 22 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "MiddlewareFactory": () => (/* binding */ MiddlewareFactory)
/* harmony export */ });
/* harmony import */ var _AuthenticationHandler__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(8);
/* harmony import */ var _HTTPMessageHandler__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(13);
/* harmony import */ var _options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(17);
/* harmony import */ var _options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(15);
/* harmony import */ var _RedirectHandler__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(16);
/* harmony import */ var _RetryHandler__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(14);
/* harmony import */ var _TelemetryHandler__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(18);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */







/**
 * @private
 * To check whether the environment is node or not
 * @returns A boolean representing the environment is node or not
 */
const isNodeEnvironment = () => {
    return typeof process === "object" && "function" === "function";
};
/**
 * @class
 * Class containing function(s) related to the middleware pipelines.
 */
class MiddlewareFactory {
    /**
     * @public
     * @static
     * Returns the default middleware chain an array with the  middleware handlers
     * @param {AuthenticationProvider} authProvider - The authentication provider instance
     * @returns an array of the middleware handlers of the default middleware chain
     */
    static getDefaultMiddlewareChain(authProvider) {
        const middleware = [];
        const authenticationHandler = new _AuthenticationHandler__WEBPACK_IMPORTED_MODULE_0__.AuthenticationHandler(authProvider);
        const retryHandler = new _RetryHandler__WEBPACK_IMPORTED_MODULE_5__.RetryHandler(new _options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.RetryHandlerOptions());
        const telemetryHandler = new _TelemetryHandler__WEBPACK_IMPORTED_MODULE_6__.TelemetryHandler();
        const httpMessageHandler = new _HTTPMessageHandler__WEBPACK_IMPORTED_MODULE_1__.HTTPMessageHandler();
        middleware.push(authenticationHandler);
        middleware.push(retryHandler);
        if (isNodeEnvironment()) {
            const redirectHandler = new _RedirectHandler__WEBPACK_IMPORTED_MODULE_4__.RedirectHandler(new _options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_2__.RedirectHandlerOptions());
            middleware.push(redirectHandler);
        }
        middleware.push(telemetryHandler);
        middleware.push(httpMessageHandler);
        return middleware;
    }
}
//# sourceMappingURL=MiddlewareFactory.js.map

/***/ }),
/* 23 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "ChaosHandlerOptions": () => (/* binding */ ChaosHandlerOptions)
/* harmony export */ });
/* harmony import */ var _ChaosStrategy__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(24);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module ChaosHandlerOptions
 */

/**
 * Class representing ChaosHandlerOptions
 * @class
 * Class
 * @implements MiddlewareOptions
 */
class ChaosHandlerOptions {
    /**
     * @public
     * @constructor
     * To create an instance of Testing Handler Options
     * @param {ChaosStrategy} ChaosStrategy - Specifies the startegy used for the Testing Handler -> RAMDOM/MANUAL
     * @param {string} statusMessage - The Message to be returned in the response
     * @param {number?} statusCode - The statusCode to be returned in the response
     * @param {number?} chaosPercentage - The percentage of randomness/chaos in the handler
     * @param {any?} responseBody - The response body to be returned in the response
     * @returns An instance of ChaosHandlerOptions
     */
    constructor(chaosStrategy = _ChaosStrategy__WEBPACK_IMPORTED_MODULE_0__.ChaosStrategy.RANDOM, statusMessage = "Some error Happened", statusCode, chaosPercentage, responseBody) {
        this.chaosStrategy = chaosStrategy;
        this.statusCode = statusCode;
        this.statusMessage = statusMessage;
        this.chaosPercentage = chaosPercentage !== undefined ? chaosPercentage : 10;
        this.responseBody = responseBody;
        if (this.chaosPercentage > 100) {
            throw new Error("Error Pecentage can not be more than 100");
        }
    }
}
//# sourceMappingURL=ChaosHandlerOptions.js.map

/***/ }),
/* 24 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "ChaosStrategy": () => (/* binding */ ChaosStrategy)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module ChaosStrategy
 */
/**
 * Strategy used for Testing Handler
 * @enum
 */
var ChaosStrategy;
(function (ChaosStrategy) {
    ChaosStrategy[ChaosStrategy["MANUAL"] = 0] = "MANUAL";
    ChaosStrategy[ChaosStrategy["RANDOM"] = 1] = "RANDOM";
})(ChaosStrategy || (ChaosStrategy = {}));
//# sourceMappingURL=ChaosStrategy.js.map

/***/ }),
/* 25 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "ChaosHandler": () => (/* binding */ ChaosHandler)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(6);
/* harmony import */ var _MiddlewareControl__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(9);
/* harmony import */ var _MiddlewareUtil__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(10);
/* harmony import */ var _options_ChaosHandlerData__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(26);
/* harmony import */ var _options_ChaosHandlerOptions__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(23);
/* harmony import */ var _options_ChaosStrategy__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(24);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */






/**
 * Class representing ChaosHandler
 * @class
 * Class
 * @implements Middleware
 */
class ChaosHandler {
    /**
     * @public
     * @constructor
     * To create an instance of Testing Handler
     * @param {ChaosHandlerOptions} [options = new ChaosHandlerOptions()] - The testing handler options instance
     * @param manualMap - The Map passed by user containing url-statusCode info
     * @returns An instance of Testing Handler
     */
    constructor(options = new _options_ChaosHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.ChaosHandlerOptions(), manualMap) {
        this.options = options;
        this.manualMap = manualMap;
    }
    /**
     * Generates responseHeader
     * @private
     * @param {number} statusCode - the status code to be returned for the request
     * @param {string} requestID - request id
     * @param {string} requestDate - date of the request
     * @returns response Header
     */
    createResponseHeaders(statusCode, requestID, requestDate) {
        const responseHeader = new Headers();
        responseHeader.append("Cache-Control", "no-store");
        responseHeader.append("request-id", requestID);
        responseHeader.append("client-request-id", requestID);
        responseHeader.append("x-ms-ags-diagnostic", "");
        responseHeader.append("Date", requestDate);
        responseHeader.append("Strict-Transport-Security", "");
        if (statusCode === 429) {
            // throttling case has to have a timeout scenario
            responseHeader.append("retry-after", "300");
        }
        return responseHeader;
    }
    /**
     * Generates responseBody
     * @private
     * @param {number} statusCode - the status code to be returned for the request
     * @param {string} statusMessage - the status message to be returned for the request
     * @param {string} requestID - request id
     * @param {string} requestDate - date of the request
     * @param {any?} requestBody - the request body to be returned for the request
     * @returns response body
     */
    createResponseBody(statusCode, statusMessage, requestID, requestDate, responseBody) {
        if (responseBody) {
            return responseBody;
        }
        let body;
        if (statusCode >= 400) {
            const codeMessage = _options_ChaosHandlerData__WEBPACK_IMPORTED_MODULE_2__.httpStatusCode[statusCode];
            const errMessage = statusMessage;
            body = {
                error: {
                    code: codeMessage,
                    message: errMessage,
                    innerError: {
                        "request-id": requestID,
                        date: requestDate,
                    },
                },
            };
        }
        else {
            body = {};
        }
        return body;
    }
    /**
     * creates a response
     * @private
     * @param {ChaosHandlerOptions} ChaosHandlerOptions - The ChaosHandlerOptions object
     * @param {Context} context - Contains the context of the request
     */
    createResponse(chaosHandlerOptions, context) {
        try {
            let responseBody;
            let responseHeader;
            let requestID;
            let requestDate;
            const requestURL = context.request;
            requestID = (0,_MiddlewareUtil__WEBPACK_IMPORTED_MODULE_1__.generateUUID)();
            requestDate = new Date();
            responseHeader = this.createResponseHeaders(chaosHandlerOptions.statusCode, requestID, requestDate.toString());
            responseBody = this.createResponseBody(chaosHandlerOptions.statusCode, chaosHandlerOptions.statusMessage, requestID, requestDate.toString(), chaosHandlerOptions.responseBody);
            const init = { url: requestURL, status: chaosHandlerOptions.statusCode, statusText: chaosHandlerOptions.statusMessage, headers: responseHeader };
            context.response = new Response(responseBody, init);
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * Decides whether to send the request to the graph or not
     * @private
     * @param {ChaosHandlerOptions} chaosHandlerOptions - A ChaosHandlerOptions object
     * @param {Context} context - Contains the context of the request
     * @returns nothing
     */
    sendRequest(chaosHandlerOptions, context) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            try {
                this.setStatusCode(chaosHandlerOptions, context.request, context.options.method);
                if (!chaosHandlerOptions.statusCode) {
                    yield this.nextMiddleware.execute(context);
                }
                else {
                    this.createResponse(chaosHandlerOptions, context);
                }
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * Fetches a random status code for the RANDOM mode from the predefined array
     * @private
     * @param {string} requestMethod - the API method for the request
     * @returns a random status code from a given set of status codes
     */
    getRandomStatusCode(requestMethod) {
        try {
            const statusCodeArray = _options_ChaosHandlerData__WEBPACK_IMPORTED_MODULE_2__.methodStatusCode[requestMethod];
            return statusCodeArray[Math.floor(Math.random() * statusCodeArray.length)];
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * To fetch the relative URL out of the complete URL using a predefined regex pattern
     * @private
     * @param {string} urlMethod - the complete URL
     * @returns the string as relative URL
     */
    getRelativeURL(urlMethod) {
        const pattern = /https?:\/\/graph\.microsoft\.com\/[^/]+(.+?)(\?|$)/;
        let relativeURL;
        if (pattern.exec(urlMethod) !== null) {
            relativeURL = pattern.exec(urlMethod)[1];
        }
        return relativeURL;
    }
    /**
     * To fetch the status code from the map(if needed), then returns response by calling createResponse
     * @private
     * @param {ChaosHandlerOptions} ChaosHandlerOptions - The ChaosHandlerOptions object
     * @param {string} requestURL - the URL for the request
     * @param {string} requestMethod - the API method for the request
     */
    setStatusCode(chaosHandlerOptions, requestURL, requestMethod) {
        try {
            if (chaosHandlerOptions.chaosStrategy === _options_ChaosStrategy__WEBPACK_IMPORTED_MODULE_4__.ChaosStrategy.MANUAL) {
                if (chaosHandlerOptions.statusCode === undefined) {
                    // manual mode with no status code, can be a global level or request level without statusCode
                    const relativeURL = this.getRelativeURL(requestURL);
                    if (this.manualMap.get(relativeURL) !== undefined) {
                        // checking Manual Map for exact match
                        if (this.manualMap.get(relativeURL).get(requestMethod) !== undefined) {
                            chaosHandlerOptions.statusCode = this.manualMap.get(relativeURL).get(requestMethod);
                        }
                        // else statusCode would be undefined
                    }
                    else {
                        // checking for regex match if exact match doesn't work
                        this.manualMap.forEach((value, key) => {
                            const regexURL = new RegExp(key + "$");
                            if (regexURL.test(relativeURL)) {
                                if (this.manualMap.get(key).get(requestMethod) !== undefined) {
                                    chaosHandlerOptions.statusCode = this.manualMap.get(key).get(requestMethod);
                                }
                                // else statusCode would be undefined
                            }
                        });
                    }
                    // Case of redirection or request url not in map ---> statusCode would be undefined
                }
            }
            else {
                // Handling the case of Random here
                if (Math.floor(Math.random() * 100) < chaosHandlerOptions.chaosPercentage) {
                    chaosHandlerOptions.statusCode = this.getRandomStatusCode(requestMethod);
                }
                // else statusCode would be undefined
            }
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * To get the options for execution of the middleware
     * @private
     * @param {Context} context - The context object
     * @returns options for middleware execution
     */
    getOptions(context) {
        let options;
        if (context.middlewareControl instanceof _MiddlewareControl__WEBPACK_IMPORTED_MODULE_0__.MiddlewareControl) {
            options = context.middlewareControl.getMiddlewareOptions(_options_ChaosHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.ChaosHandlerOptions);
        }
        if (typeof options === "undefined") {
            options = Object.assign(new _options_ChaosHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.ChaosHandlerOptions(), this.options);
        }
        return options;
    }
    /**
     * To execute the current middleware
     * @public
     * @async
     * @param {Context} context - The context object of the request
     * @returns A Promise that resolves to nothing
     */
    execute(context) {
        return tslib__WEBPACK_IMPORTED_MODULE_5__.__awaiter(this, void 0, void 0, function* () {
            try {
                const chaosHandlerOptions = this.getOptions(context);
                return yield this.sendRequest(chaosHandlerOptions, context);
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * To set the next middleware in the chain
     * @param {Middleware} next - The middleware instance
     * @returns Nothing
     */
    setNext(next) {
        this.nextMiddleware = next;
    }
}
//# sourceMappingURL=ChaosHandler.js.map

/***/ }),
/* 26 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "methodStatusCode": () => (/* binding */ methodStatusCode),
/* harmony export */   "httpStatusCode": () => (/* binding */ httpStatusCode)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module ChaosHandlerData
 */
/**
 * Contains RequestMethod to corresponding array of possible status codes, used for Random mode
 */
const methodStatusCode = {
    GET: [429, 500, 502, 503, 504],
    POST: [429, 500, 502, 503, 504, 507],
    PUT: [429, 500, 502, 503, 504, 507],
    PATCH: [429, 500, 502, 503, 504],
    DELETE: [429, 500, 502, 503, 504, 507],
};
/**
 * Contains statusCode to statusMessage map
 */
const httpStatusCode = {
    100: "Continue",
    101: "Switching Protocols",
    102: "Processing",
    103: "Early Hints",
    200: "OK",
    201: "Created",
    202: "Accepted",
    203: "Non-Authoritative Information",
    204: "No Content",
    205: "Reset Content",
    206: "Partial Content",
    207: "Multi-Status",
    208: "Already Reported",
    226: "IM Used",
    300: "Multiple Choices",
    301: "Moved Permanently",
    302: "Found",
    303: "See Other",
    304: "Not Modified",
    305: "Use Proxy",
    307: "Temporary Redirect",
    308: "Permanent Redirect",
    400: "Bad Request",
    401: "Unauthorized",
    402: "Payment Required",
    403: "Forbidden",
    404: "Not Found",
    405: "Method Not Allowed",
    406: "Not Acceptable",
    407: "Proxy Authentication Required",
    408: "Request Timeout",
    409: "Conflict",
    410: "Gone",
    411: "Length Required",
    412: "Precondition Failed",
    413: "Payload Too Large",
    414: "URI Too Long",
    415: "Unsupported Media Type",
    416: "Range Not Satisfiable",
    417: "Expectation Failed",
    421: "Misdirected Request",
    422: "Unprocessable Entity",
    423: "Locked",
    424: "Failed Dependency",
    425: "Too Early",
    426: "Upgrade Required",
    428: "Precondition Required",
    429: "Too Many Requests",
    431: "Request Header Fields Too Large",
    451: "Unavailable For Legal Reasons",
    500: "Internal Server Error",
    501: "Not Implemented",
    502: "Bad Gateway",
    503: "Service Unavailable",
    504: "Gateway Timeout",
    505: "HTTP Version Not Supported",
    506: "Variant Also Negotiates",
    507: "Insufficient Storage",
    508: "Loop Detected",
    510: "Not Extended",
    511: "Network Authentication Required",
};
//# sourceMappingURL=ChaosHandlerData.js.map

/***/ }),
/* 27 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "LargeFileUploadTask": () => (/* binding */ LargeFileUploadTask)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(6);
/* harmony import */ var _Range__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(28);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */


/**
 * @class
 * Class representing LargeFileUploadTask
 */
class LargeFileUploadTask {
    /**
     * @public
     * @constructor
     * Constructs a LargeFileUploadTask
     * @param {Client} client - The GraphClient instance
     * @param {FileObject} file - The FileObject holding details of a file that needs to be uploaded
     * @param {LargeFileUploadSession} uploadSession - The upload session to which the upload has to be done
     * @param {LargeFileUploadTaskOptions} options - The upload task options
     * @returns An instance of LargeFileUploadTask
     */
    constructor(client, file, uploadSession, options = {}) {
        /**
         * @private
         * Default value for the rangeSize
         */
        this.DEFAULT_FILE_SIZE = 5 * 1024 * 1024;
        this.client = client;
        this.file = file;
        if (options.rangeSize === undefined) {
            options.rangeSize = this.DEFAULT_FILE_SIZE;
        }
        this.options = options;
        this.uploadSession = uploadSession;
        this.nextRange = new _Range__WEBPACK_IMPORTED_MODULE_0__.Range(0, this.options.rangeSize - 1);
    }
    /**
     * @public
     * @static
     * @async
     * Makes request to the server to create an upload session
     * @param {Client} client - The GraphClient instance
     * @param {any} payload - The payload that needs to be sent
     * @param {KeyValuePairObjectStringNumber} headers - The headers that needs to be sent
     * @returns The promise that resolves to LargeFileUploadSession
     */
    static createUploadSession(client, requestUrl, payload, headers = {}) {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            try {
                const session = yield client
                    .api(requestUrl)
                    .headers(headers)
                    .post(payload);
                const largeFileUploadSession = {
                    url: session.uploadUrl,
                    expiry: new Date(session.expirationDateTime),
                };
                return largeFileUploadSession;
            }
            catch (err) {
                throw err;
            }
        });
    }
    /**
     * @private
     * Parses given range string to the Range instance
     * @param {string[]} ranges - The ranges value
     * @returns The range instance
     */
    parseRange(ranges) {
        const rangeStr = ranges[0];
        if (typeof rangeStr === "undefined" || rangeStr === "") {
            return new _Range__WEBPACK_IMPORTED_MODULE_0__.Range();
        }
        const firstRange = rangeStr.split("-");
        const minVal = parseInt(firstRange[0], 10);
        let maxVal = parseInt(firstRange[1], 10);
        if (Number.isNaN(maxVal)) {
            maxVal = this.file.size - 1;
        }
        return new _Range__WEBPACK_IMPORTED_MODULE_0__.Range(minVal, maxVal);
    }
    /**
     * @private
     * Updates the expiration date and the next range
     * @param {UploadStatusResponse} response - The response of the upload status
     * @returns Nothing
     */
    updateTaskStatus(response) {
        this.uploadSession.expiry = new Date(response.expirationDateTime);
        this.nextRange = this.parseRange(response.nextExpectedRanges);
    }
    /**
     * @public
     * Gets next range that needs to be uploaded
     * @returns The range instance
     */
    getNextRange() {
        if (this.nextRange.minValue === -1) {
            return this.nextRange;
        }
        const minVal = this.nextRange.minValue;
        let maxValue = minVal + this.options.rangeSize - 1;
        if (maxValue >= this.file.size) {
            maxValue = this.file.size - 1;
        }
        return new _Range__WEBPACK_IMPORTED_MODULE_0__.Range(minVal, maxValue);
    }
    /**
     * @public
     * Slices the file content to the given range
     * @param {Range} range - The range value
     * @returns The sliced ArrayBuffer or Blob
     */
    sliceFile(range) {
        const blob = this.file.content.slice(range.minValue, range.maxValue + 1);
        return blob;
    }
    /**
     * @public
     * @async
     * Uploads file to the server in a sequential order by slicing the file
     * @returns The promise resolves to uploaded response
     */
    upload() {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            try {
                while (true) {
                    const nextRange = this.getNextRange();
                    if (nextRange.maxValue === -1) {
                        const err = new Error("Task with which you are trying to upload is already completed, Please check for your uploaded file");
                        err.name = "Invalid Session";
                        throw err;
                    }
                    const fileSlice = this.sliceFile(nextRange);
                    const response = yield this.uploadSlice(fileSlice, nextRange, this.file.size);
                    // Upon completion of upload process incase of onedrive, driveItem is returned, which contains id
                    if (response.id !== undefined) {
                        return response;
                    }
                    else {
                        this.updateTaskStatus(response);
                    }
                }
            }
            catch (err) {
                throw err;
            }
        });
    }
    /**
     * @public
     * @async
     * Uploads given slice to the server
     * @param {ArrayBuffer | Blob | File} fileSlice - The file slice
     * @param {Range} range - The range value
     * @param {number} totalSize - The total size of a complete file
     */
    uploadSlice(fileSlice, range, totalSize) {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            try {
                return yield this.client
                    .api(this.uploadSession.url)
                    .headers({
                    "Content-Length": `${range.maxValue - range.minValue + 1}`,
                    "Content-Range": `bytes ${range.minValue}-${range.maxValue}/${totalSize}`,
                })
                    .put(fileSlice);
            }
            catch (err) {
                throw err;
            }
        });
    }
    /**
     * @public
     * @async
     * Deletes upload session in the server
     * @returns The promise resolves to cancelled response
     */
    cancel() {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            try {
                return yield this.client.api(this.uploadSession.url).delete();
            }
            catch (err) {
                throw err;
            }
        });
    }
    /**
     * @public
     * @async
     * Gets status for the upload session
     * @returns The promise resolves to the status enquiry response
     */
    getStatus() {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            try {
                const response = yield this.client.api(this.uploadSession.url).get();
                this.updateTaskStatus(response);
                return response;
            }
            catch (err) {
                throw err;
            }
        });
    }
    /**
     * @public
     * @async
     * Resumes upload session and continue uploading the file from the last sent range
     * @returns The promise resolves to the uploaded response
     */
    resume() {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            try {
                yield this.getStatus();
                return yield this.upload();
            }
            catch (err) {
                throw err;
            }
        });
    }
}
//# sourceMappingURL=LargeFileUploadTask.js.map

/***/ }),
/* 28 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "Range": () => (/* binding */ Range)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module Range
 */
/**
 * @class
 * Class representing Range
 */
class Range {
    /**
     * @public
     * @constructor
     * Creates a range for given min and max values
     * @param {number} [minVal = -1] - The minimum value.
     * @param {number} [maxVal = -1] - The maximum value.
     * @returns An instance of a Range
     */
    constructor(minVal = -1, maxVal = -1) {
        this.minValue = minVal;
        this.maxValue = maxVal;
    }
}
//# sourceMappingURL=Range.js.map

/***/ }),
/* 29 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "OneDriveLargeFileUploadTask": () => (/* binding */ OneDriveLargeFileUploadTask)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(6);
/* harmony import */ var _LargeFileUploadTask__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(27);
/* harmony import */ var _OneDriveLargeFileUploadTaskUtil__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(30);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */



/**
 * @class
 * Class representing OneDriveLargeFileUploadTask
 */
class OneDriveLargeFileUploadTask extends _LargeFileUploadTask__WEBPACK_IMPORTED_MODULE_0__.LargeFileUploadTask {
    /**
     * @public
     * @constructor
     * Constructs a OneDriveLargeFileUploadTask
     * @param {Client} client - The GraphClient instance
     * @param {FileObject} file - The FileObject holding details of a file that needs to be uploaded
     * @param {LargeFileUploadSession} uploadSession - The upload session to which the upload has to be done
     * @param {LargeFileUploadTaskOptions} options - The upload task options
     * @returns An instance of OneDriveLargeFileUploadTask
     */
    constructor(client, file, uploadSession, options) {
        super(client, file, uploadSession, options);
    }
    /**
     * @private
     * @static
     * Constructs the create session url for Onedrive
     * @param {string} fileName - The name of the file
     * @param {path} [path = OneDriveLargeFileUploadTask.DEFAULT_UPLOAD_PATH] - The path for the upload
     * @returns The constructed create session url
     */
    static constructCreateSessionUrl(fileName, path = OneDriveLargeFileUploadTask.DEFAULT_UPLOAD_PATH) {
        fileName = fileName.trim();
        path = path.trim();
        if (path === "") {
            path = "/";
        }
        if (path[0] !== "/") {
            path = `/${path}`;
        }
        if (path[path.length - 1] !== "/") {
            path = `${path}/`;
        }
        // we choose to encode each component of the file path separately because when encoding full URI
        // with encodeURI, special characters like # or % in the file name doesn't get encoded as desired
        return `/me/drive/root:${path
            .split("/")
            .map((p) => encodeURIComponent(p))
            .join("/")}${encodeURIComponent(fileName)}:/createUploadSession`;
    }
    /**
     * @public
     * @static
     * @async
     * Creates a OneDriveLargeFileUploadTask
     * @param {Client} client - The GraphClient instance
     * @param {Blob | Buffer | File} file - File represented as Blob, Buffer or File
     * @param {OneDriveLargeFileUploadOptions} options - The options for upload task
     * @returns The promise that will be resolves to OneDriveLargeFileUploadTask instance
     */
    static create(client, file, options) {
        return tslib__WEBPACK_IMPORTED_MODULE_2__.__awaiter(this, void 0, void 0, function* () {
            const name = options.fileName;
            let content;
            let size;
            if (typeof Blob !== "undefined" && file instanceof Blob) {
                content = new File([file], name);
                size = content.size;
            }
            else if (typeof File !== "undefined" && file instanceof File) {
                content = file;
                size = content.size;
            }
            else if (typeof Buffer !== "undefined" && file instanceof Buffer) {
                const b = file;
                size = b.byteLength - b.byteOffset;
                content = b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
            }
            try {
                const requestUrl = OneDriveLargeFileUploadTask.constructCreateSessionUrl(options.fileName, options.path);
                const session = yield OneDriveLargeFileUploadTask.createUploadSession(client, requestUrl, options.fileName);
                const rangeSize = (0,_OneDriveLargeFileUploadTaskUtil__WEBPACK_IMPORTED_MODULE_1__.getValidRangeSize)(options.rangeSize);
                const fileObj = {
                    name,
                    content,
                    size,
                };
                return new OneDriveLargeFileUploadTask(client, fileObj, session, {
                    rangeSize,
                });
            }
            catch (err) {
                throw err;
            }
        });
    }
    /**
     * @public
     * @static
     * @async
     * Makes request to the server to create an upload session
     * @param {Client} client - The GraphClient instance
     * @param {string} requestUrl - The URL to create the upload session
     * @param {string} fileName - The name of a file to upload, (with extension)
     * @returns The promise that resolves to LargeFileUploadSession
     */
    static createUploadSession(client, requestUrl, fileName) {
        const _super = Object.create(null, {
            createUploadSession: { get: () => super.createUploadSession }
        });
        return tslib__WEBPACK_IMPORTED_MODULE_2__.__awaiter(this, void 0, void 0, function* () {
            const payload = {
                item: {
                    "@microsoft.graph.conflictBehavior": "rename",
                    name: fileName,
                },
            };
            try {
                return _super.createUploadSession.call(this, client, requestUrl, payload);
            }
            catch (err) {
                throw err;
            }
        });
    }
    /**
     * @public
     * Commits upload session to end uploading
     * @param {string} requestUrl - The URL to commit the upload session
     * @returns The promise resolves to committed response
     */
    commit(requestUrl) {
        return tslib__WEBPACK_IMPORTED_MODULE_2__.__awaiter(this, void 0, void 0, function* () {
            try {
                const payload = {
                    name: this.file.name,
                    "@microsoft.graph.conflictBehavior": "rename",
                    "@microsoft.graph.sourceUrl": this.uploadSession.url,
                };
                return yield this.client.api(requestUrl).put(payload);
            }
            catch (err) {
                throw err;
            }
        });
    }
}
/**
 * @private
 * @static
 * Default path for the file being uploaded
 */
OneDriveLargeFileUploadTask.DEFAULT_UPLOAD_PATH = "/";
//# sourceMappingURL=OneDriveLargeFileUploadTask.js.map

/***/ }),
/* 30 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "getValidRangeSize": () => (/* binding */ getValidRangeSize)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module OneDriveLargeFileUploadTaskUtil
 */
/**
 * @constant
 * Default value for the rangeSize
 * Recommended size is between 5 - 10 MB {@link https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/driveitem_createuploadsession#best-practices}
 */
const DEFAULT_FILE_SIZE = 5 * 1024 * 1024;
/**
 * @constant
 * Rounds off the given value to a multiple of 320 KB
 * @param {number} value - The value
 * @returns The rounded off value
 */
const roundTo320KB = (value) => {
    if (value > 320 * 1024) {
        value = Math.floor(value / (320 * 1024)) * 320 * 1024;
    }
    return value;
};
/**
 * @constant
 * Get the valid rangeSize for a file slicing (validity is based on the constrains mentioned in here
 * {@link https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/driveitem_createuploadsession#upload-bytes-to-the-upload-session})
 *
 * @param {number} [rangeSize = DEFAULT_FILE_SIZE] - The rangeSize value.
 * @returns The valid rangeSize
 */
const getValidRangeSize = (rangeSize = DEFAULT_FILE_SIZE) => {
    const sixtyMB = 60 * 1024 * 1024;
    if (rangeSize > sixtyMB) {
        rangeSize = sixtyMB;
    }
    return roundTo320KB(rangeSize);
};
//# sourceMappingURL=OneDriveLargeFileUploadTaskUtil.js.map

/***/ }),
/* 31 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "PageIterator": () => (/* binding */ PageIterator)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @class
 * Class for PageIterator
 */
class PageIterator {
    /**
     * @public
     * @constructor
     * Creates new instance for PageIterator
     * @param {Client} client - The graph client instance
     * @param {PageCollection} pageCollection - The page collection object
     * @param {PageIteratorCallback} callBack - The callback function
     * @param {GraphRequestOptions} requestOptions - The request options
     * @returns An instance of a PageIterator
     */
    constructor(client, pageCollection, callback, requestOptions) {
        this.client = client;
        this.collection = pageCollection.value;
        this.nextLink = pageCollection["@odata.nextLink"];
        this.deltaLink = pageCollection["@odata.deltaLink"];
        this.callback = callback;
        this.complete = false;
        this.requestOptions = requestOptions;
    }
    /**
     * @private
     * Iterates over a collection by enqueuing entries one by one and kicking the callback with the enqueued entry
     * @returns A boolean indicating the continue flag to process next page
     */
    iterationHelper() {
        if (this.collection === undefined) {
            return false;
        }
        let advance = true;
        while (advance && this.collection.length !== 0) {
            const item = this.collection.shift();
            advance = this.callback(item);
        }
        return advance;
    }
    /**
     * @private
     * @async
     * Helper to make a get request to fetch next page with nextLink url and update the page iterator instance with the returned response
     * @returns A promise that resolves to a response data with next page collection
     */
    fetchAndUpdateNextPageData() {
        return tslib__WEBPACK_IMPORTED_MODULE_0__.__awaiter(this, void 0, void 0, function* () {
            try {
                let graphRequest = this.client.api(this.nextLink);
                if (this.requestOptions) {
                    if (this.requestOptions.headers) {
                        graphRequest = graphRequest.headers(this.requestOptions.headers);
                    }
                    if (this.requestOptions.middlewareOptions) {
                        graphRequest = graphRequest.middlewareOptions(this.requestOptions.middlewareOptions);
                    }
                    if (this.requestOptions.options) {
                        graphRequest = graphRequest.options(this.requestOptions.options);
                    }
                }
                const response = yield graphRequest.get();
                this.collection = response.value;
                this.nextLink = response["@odata.nextLink"];
                this.deltaLink = response["@odata.deltaLink"];
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * Getter to get the deltaLink in the current response
     * @returns A deltaLink which is being used to make delta requests in future
     */
    getDeltaLink() {
        return this.deltaLink;
    }
    /**
     * @public
     * @async
     * Iterates over the collection and kicks callback for each item on iteration. Fetches next set of data through nextLink and iterates over again
     * This happens until the nextLink is drained out or the user responds with a red flag to continue from callback
     * @returns A Promise that resolves to nothing on completion and throws error incase of any discrepancy.
     */
    iterate() {
        return tslib__WEBPACK_IMPORTED_MODULE_0__.__awaiter(this, void 0, void 0, function* () {
            try {
                let advance = this.iterationHelper();
                while (advance) {
                    if (this.nextLink !== undefined) {
                        yield this.fetchAndUpdateNextPageData();
                        advance = this.iterationHelper();
                    }
                    else {
                        advance = false;
                    }
                }
                if (this.nextLink === undefined && this.collection.length === 0) {
                    this.complete = true;
                }
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * To resume the iteration
     * Note: This internally calls the iterate method, It's just for more readability.
     * @returns A Promise that resolves to nothing on completion and throws error incase of any discrepancy
     */
    resume() {
        return tslib__WEBPACK_IMPORTED_MODULE_0__.__awaiter(this, void 0, void 0, function* () {
            try {
                return this.iterate();
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * To get the completeness status of the iterator
     * @returns Boolean indicating the completeness
     */
    isComplete() {
        return this.complete;
    }
}
//# sourceMappingURL=PageIterator.js.map

/***/ }),
/* 32 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "Client": () => (/* binding */ Client)
/* harmony export */ });
/* harmony import */ var _Constants__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(20);
/* harmony import */ var _CustomAuthenticationProvider__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(33);
/* harmony import */ var _GraphRequest__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(34);
/* harmony import */ var _HTTPClient__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(39);
/* harmony import */ var _HTTPClientFactory__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(40);
/* harmony import */ var _ValidatePolyFilling__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(41);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module Client
 */






class Client {
    /**
     * @private
     * @constructor
     * Creates an instance of Client
     * @param {ClientOptions} clientOptions - The options to instantiate the client object
     */
    constructor(clientOptions) {
        /**
         * @private
         * A member which stores the Client instance options
         */
        this.config = {
            baseUrl: _Constants__WEBPACK_IMPORTED_MODULE_0__.GRAPH_BASE_URL,
            debugLogging: false,
            defaultVersion: _Constants__WEBPACK_IMPORTED_MODULE_0__.GRAPH_API_VERSION,
        };
        try {
            (0,_ValidatePolyFilling__WEBPACK_IMPORTED_MODULE_5__.validatePolyFilling)();
        }
        catch (error) {
            throw error;
        }
        for (const key in clientOptions) {
            if (clientOptions.hasOwnProperty(key)) {
                this.config[key] = clientOptions[key];
            }
        }
        let httpClient;
        if (clientOptions.authProvider !== undefined && clientOptions.middleware !== undefined) {
            const error = new Error();
            error.name = "AmbiguityInInitialization";
            error.message = "Unable to Create Client, Please provide either authentication provider for default middleware chain or custom middleware chain not both";
            throw error;
        }
        else if (clientOptions.authProvider !== undefined) {
            httpClient = _HTTPClientFactory__WEBPACK_IMPORTED_MODULE_4__.HTTPClientFactory.createWithAuthenticationProvider(clientOptions.authProvider);
        }
        else if (clientOptions.middleware !== undefined) {
            httpClient = new _HTTPClient__WEBPACK_IMPORTED_MODULE_3__.HTTPClient(...[].concat(clientOptions.middleware));
        }
        else {
            const error = new Error();
            error.name = "InvalidMiddlewareChain";
            error.message = "Unable to Create Client, Please provide either authentication provider for default middleware chain or custom middleware chain";
            throw error;
        }
        this.httpClient = httpClient;
    }
    /**
     * @public
     * @static
     * To create a client instance with options and initializes the default middleware chain
     * @param {Options} options - The options for client instance
     * @returns The Client instance
     */
    static init(options) {
        const clientOptions = {};
        for (const i in options) {
            if (options.hasOwnProperty(i)) {
                clientOptions[i] = i === "authProvider" ? new _CustomAuthenticationProvider__WEBPACK_IMPORTED_MODULE_1__.CustomAuthenticationProvider(options[i]) : options[i];
            }
        }
        return Client.initWithMiddleware(clientOptions);
    }
    /**
     * @public
     * @static
     * To create a client instance with the Client Options
     * @param {ClientOptions} clientOptions - The options object for initializing the client
     * @returns The Client instance
     */
    static initWithMiddleware(clientOptions) {
        try {
            return new Client(clientOptions);
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * @public
     * Entry point to make requests
     * @param {string} path - The path string value
     * @returns The graph request instance
     */
    api(path) {
        return new _GraphRequest__WEBPACK_IMPORTED_MODULE_2__.GraphRequest(this.httpClient, this.config, path);
    }
}
//# sourceMappingURL=Client.js.map

/***/ }),
/* 33 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "CustomAuthenticationProvider": () => (/* binding */ CustomAuthenticationProvider)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @class
 * Class representing CustomAuthenticationProvider
 * @extends AuthenticationProvider
 */
class CustomAuthenticationProvider {
    /**
     * @public
     * @constructor
     * Creates an instance of CustomAuthenticationProvider
     * @param {AuthProviderCallback} provider - An authProvider function
     * @returns An instance of CustomAuthenticationProvider
     */
    constructor(provider) {
        this.provider = provider;
    }
    /**
     * @public
     * @async
     * To get the access token
     * @returns The promise that resolves to an access token
     */
    getAccessToken() {
        return tslib__WEBPACK_IMPORTED_MODULE_0__.__awaiter(this, void 0, void 0, function* () {
            return new Promise((resolve, reject) => {
                this.provider((error, accessToken) => {
                    if (accessToken) {
                        resolve(accessToken);
                    }
                    else {
                        reject(error);
                    }
                });
            });
        });
    }
}
//# sourceMappingURL=CustomAuthenticationProvider.js.map

/***/ }),
/* 34 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "GraphRequest": () => (/* binding */ GraphRequest)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(6);
/* harmony import */ var _GraphErrorHandler__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(35);
/* harmony import */ var _GraphRequestUtil__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(19);
/* harmony import */ var _GraphResponseHandler__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(37);
/* harmony import */ var _middleware_MiddlewareControl__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(9);
/* harmony import */ var _RequestMethod__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(5);
/* harmony import */ var _ResponseType__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(38);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */







/**
 * @class
 * A Class representing GraphRequest
 */
class GraphRequest {
    /* tslint:enable: variable-name */
    /**
     * @public
     * @constructor
     * Creates an instance of GraphRequest
     * @param {HTTPClient} httpClient - The HTTPClient instance
     * @param {ClientOptions} config - The options for making request
     * @param {string} path - A path string
     */
    constructor(httpClient, config, path) {
        /**
         * @private
         * Parses the path string and creates URLComponents out of it
         * @param {string} path - The request path string
         * @returns Nothing
         */
        this.parsePath = (path) => {
            // Strips out the base of the url if they passed in
            if (path.indexOf("https://") !== -1) {
                path = path.replace("https://", "");
                // Find where the host ends
                const endOfHostStrPos = path.indexOf("/");
                if (endOfHostStrPos !== -1) {
                    // Parse out the host
                    this.urlComponents.host = "https://" + path.substring(0, endOfHostStrPos);
                    // Strip the host from path
                    path = path.substring(endOfHostStrPos + 1, path.length);
                }
                // Remove the following version
                const endOfVersionStrPos = path.indexOf("/");
                if (endOfVersionStrPos !== -1) {
                    // Parse out the version
                    this.urlComponents.version = path.substring(0, endOfVersionStrPos);
                    // Strip version from path
                    path = path.substring(endOfVersionStrPos + 1, path.length);
                }
            }
            // Strip out any leading "/"
            if (path.charAt(0) === "/") {
                path = path.substr(1);
            }
            const queryStrPos = path.indexOf("?");
            if (queryStrPos === -1) {
                // No query string
                this.urlComponents.path = path;
            }
            else {
                this.urlComponents.path = path.substr(0, queryStrPos);
                // Capture query string into oDataQueryParams and otherURLQueryParams
                const queryParams = path.substring(queryStrPos + 1, path.length).split("&");
                for (const queryParam of queryParams) {
                    this.parseQueryParameter(queryParam);
                }
            }
        };
        this.httpClient = httpClient;
        this.config = config;
        this.urlComponents = {
            host: this.config.baseUrl,
            version: this.config.defaultVersion,
            oDataQueryParams: {},
            otherURLQueryParams: {},
            otherURLQueryOptions: [],
        };
        this._headers = {};
        this._options = {};
        this._middlewareOptions = [];
        this.parsePath(path);
    }
    /**
     * @private
     * Adds the query parameter as comma separated values
     * @param {string} propertyName - The name of a property
     * @param {string|string[]} propertyValue - The vale of a property
     * @param {IArguments} additionalProperties - The additional properties
     * @returns Nothing
     */
    addCsvQueryParameter(propertyName, propertyValue, additionalProperties) {
        // If there are already $propertyName value there, append a ","
        this.urlComponents.oDataQueryParams[propertyName] = this.urlComponents.oDataQueryParams[propertyName] ? this.urlComponents.oDataQueryParams[propertyName] + "," : "";
        let allValues = [];
        if (additionalProperties.length > 1 && typeof propertyValue === "string") {
            allValues = Array.prototype.slice.call(additionalProperties);
        }
        else if (typeof propertyValue === "string") {
            allValues.push(propertyValue);
        }
        else {
            allValues = allValues.concat(propertyValue);
        }
        this.urlComponents.oDataQueryParams[propertyName] += allValues.join(",");
    }
    /**
     * @private
     * Builds the full url from the URLComponents to make a request
     * @returns The URL string that is qualified to make a request to graph endpoint
     */
    buildFullUrl() {
        const url = (0,_GraphRequestUtil__WEBPACK_IMPORTED_MODULE_1__.urlJoin)([this.urlComponents.host, this.urlComponents.version, this.urlComponents.path]) + this.createQueryString();
        if (this.config.debugLogging) {
            console.log(url); // tslint:disable-line: no-console
        }
        return url;
    }
    /**
     * @private
     * Builds the query string from the URLComponents
     * @returns The Constructed query string
     */
    createQueryString() {
        // Combining query params from oDataQueryParams and otherURLQueryParams
        const urlComponents = this.urlComponents;
        const query = [];
        if (Object.keys(urlComponents.oDataQueryParams).length !== 0) {
            for (const property in urlComponents.oDataQueryParams) {
                if (urlComponents.oDataQueryParams.hasOwnProperty(property)) {
                    query.push(property + "=" + urlComponents.oDataQueryParams[property]);
                }
            }
        }
        if (Object.keys(urlComponents.otherURLQueryParams).length !== 0) {
            for (const property in urlComponents.otherURLQueryParams) {
                if (urlComponents.otherURLQueryParams.hasOwnProperty(property)) {
                    query.push(property + "=" + urlComponents.otherURLQueryParams[property]);
                }
            }
        }
        if (urlComponents.otherURLQueryOptions.length !== 0) {
            for (const str of urlComponents.otherURLQueryOptions) {
                query.push(str);
            }
        }
        return query.length > 0 ? "?" + query.join("&") : "";
    }
    /**
     * @private
     * Parses the query parameters to set the urlComponents property of the GraphRequest object
     * @param {string|KeyValuePairObjectStringNumber} queryDictionaryOrString - The query parameter
     * @returns The same GraphRequest instance that is being called with
     */
    parseQueryParameter(queryDictionaryOrString) {
        if (typeof queryDictionaryOrString === "string") {
            if (queryDictionaryOrString.charAt(0) === "?") {
                queryDictionaryOrString = queryDictionaryOrString.substring(1);
            }
            if (queryDictionaryOrString.indexOf("&") !== -1) {
                const queryParams = queryDictionaryOrString.split("&");
                for (const str of queryParams) {
                    this.parseQueryParamenterString(str);
                }
            }
            else {
                this.parseQueryParamenterString(queryDictionaryOrString);
            }
        }
        else if (queryDictionaryOrString.constructor === Object) {
            for (const key in queryDictionaryOrString) {
                if (queryDictionaryOrString.hasOwnProperty(key)) {
                    this.setURLComponentsQueryParamater(key, queryDictionaryOrString[key]);
                }
            }
        }
        return this;
    }
    /**
     * @private
     * Parses the query parameter of string type to set the urlComponents property of the GraphRequest object
     * @param {string} queryParameter - the query parameters
     * returns nothing
     */
    parseQueryParamenterString(queryParameter) {
        /* The query key-value pair must be split on the first equals sign to avoid errors in parsing nested query parameters.
                 Example-> "/me?$expand=home($select=city)" */
        if (this.isValidQueryKeyValuePair(queryParameter)) {
            const indexOfFirstEquals = queryParameter.indexOf("=");
            const paramKey = queryParameter.substring(0, indexOfFirstEquals);
            const paramValue = queryParameter.substring(indexOfFirstEquals + 1);
            this.setURLComponentsQueryParamater(paramKey, paramValue);
        }
        else {
            /* Push values which are not of key-value structure.
            Example-> Handle an invalid input->.query(test), .query($select($select=name)) and let the Graph API respond with the error in the URL*/
            this.urlComponents.otherURLQueryOptions.push(queryParameter);
        }
    }
    /**
     * @private
     * Sets values into the urlComponents property of GraphRequest object.
     * @param {string} paramKey - the query parameter key
     * @param {string} paramValue - the query paramter value
     * @returns nothing
     */
    setURLComponentsQueryParamater(paramKey, paramValue) {
        if (_GraphRequestUtil__WEBPACK_IMPORTED_MODULE_1__.oDataQueryNames.indexOf(paramKey) !== -1) {
            const currentValue = this.urlComponents.oDataQueryParams[paramKey];
            const isValueAppendable = currentValue && (paramKey === "$expand" || paramKey === "$select" || paramKey === "$orderby");
            this.urlComponents.oDataQueryParams[paramKey] = isValueAppendable ? currentValue + "," + paramValue : paramValue;
        }
        else {
            this.urlComponents.otherURLQueryParams[paramKey] = paramValue;
        }
    }
    /**
     * @private
     * Check if the query parameter string has a valid key-value structure
     * @param {string} queryString - the query parameter string. Example -> "name=value"
     * #returns true if the query string has a valid key-value structure else false
     */
    isValidQueryKeyValuePair(queryString) {
        const indexofFirstEquals = queryString.indexOf("=");
        if (indexofFirstEquals === -1) {
            return false;
        }
        const indexofOpeningParanthesis = queryString.indexOf("(");
        if (indexofOpeningParanthesis !== -1 && queryString.indexOf("(") < indexofFirstEquals) {
            // Example -> .query($select($expand=true));
            return false;
        }
        return true;
    }
    /**
     * @private
     * Updates the custom headers and options for a request
     * @param {FetchOptions} options - The request options object
     * @returns Nothing
     */
    updateRequestOptions(options) {
        const optionsHeaders = Object.assign({}, options.headers);
        if (this.config.fetchOptions !== undefined) {
            const fetchOptions = Object.assign({}, this.config.fetchOptions);
            Object.assign(options, fetchOptions);
            if (typeof this.config.fetchOptions.headers !== undefined) {
                options.headers = Object.assign({}, this.config.fetchOptions.headers);
            }
        }
        Object.assign(options, this._options);
        if (options.headers !== undefined) {
            Object.assign(optionsHeaders, options.headers);
        }
        Object.assign(optionsHeaders, this._headers);
        options.headers = optionsHeaders;
    }
    /**
     * @private
     * @async
     * Adds the custom headers and options to the request and makes the HTTPClient send request call
     * @param {RequestInfo} request - The request url string or the Request object value
     * @param {FetchOptions} options - The options to make a request
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the response content
     */
    send(request, options, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            let rawResponse;
            const middlewareControl = new _middleware_MiddlewareControl__WEBPACK_IMPORTED_MODULE_3__.MiddlewareControl(this._middlewareOptions);
            this.updateRequestOptions(options);
            try {
                const context = yield this.httpClient.sendRequest({
                    request,
                    options,
                    middlewareControl,
                });
                rawResponse = context.response;
                const response = yield _GraphResponseHandler__WEBPACK_IMPORTED_MODULE_2__.GraphResponseHandler.getResponse(rawResponse, this._responseType, callback);
                return response;
            }
            catch (error) {
                let statusCode;
                if (typeof rawResponse !== "undefined") {
                    statusCode = rawResponse.status;
                }
                const gError = yield _GraphErrorHandler__WEBPACK_IMPORTED_MODULE_0__.GraphErrorHandler.getError(error, statusCode, callback);
                throw gError;
            }
        });
    }
    /**
     * @private
     * Checks if the content-type is present in the _headers property. If not present, defaults the content-type to application/json
     * @param none
     * @returns nothing
     */
    setHeaderContentType() {
        if (!this._headers) {
            this.header("Content-Type", "application/json");
            return;
        }
        const headerKeys = Object.keys(this._headers);
        for (const headerKey of headerKeys) {
            if (headerKey.toLowerCase() === "content-type") {
                return;
            }
        }
        // Default the content-type to application/json in case the content-type is not present in the header
        this.header("Content-Type", "application/json");
    }
    /**
     * @public
     * Sets the custom header for a request
     * @param {string} headerKey - A header key
     * @param {string} headerValue - A header value
     * @returns The same GraphRequest instance that is being called with
     */
    header(headerKey, headerValue) {
        this._headers[headerKey] = headerValue;
        return this;
    }
    /**
     * @public
     * Sets the custom headers for a request
     * @param {KeyValuePairObjectStringNumber | HeadersInit} headers - The request headers
     * @returns The same GraphRequest instance that is being called with
     */
    headers(headers) {
        for (const key in headers) {
            if (headers.hasOwnProperty(key)) {
                this._headers[key] = headers[key];
            }
        }
        return this;
    }
    /**
     * @public
     * Sets the option for making a request
     * @param {string} key - The key value
     * @param {any} value - The value
     * @returns The same GraphRequest instance that is being called with
     */
    option(key, value) {
        this._options[key] = value;
        return this;
    }
    /**
     * @public
     * Sets the options for making a request
     * @param {{ [key: string]: any }} options - The options key value pair
     * @returns The same GraphRequest instance that is being called with
     */
    options(options) {
        for (const key in options) {
            if (options.hasOwnProperty(key)) {
                this._options[key] = options[key];
            }
        }
        return this;
    }
    /**
     * @public
     * Sets the middleware options for a request
     * @param {MiddlewareOptions[]} options - The array of middleware options
     * @returns The same GraphRequest instance that is being called with
     */
    middlewareOptions(options) {
        this._middlewareOptions = options;
        return this;
    }
    /**
     * @public
     * Sets the api endpoint version for a request
     * @param {string} version - The version value
     * @returns The same GraphRequest instance that is being called with
     */
    version(version) {
        this.urlComponents.version = version;
        return this;
    }
    /**
     * @public
     * Sets the api endpoint version for a request
     * @param {ResponseType} responseType - The response type value
     * @returns The same GraphRequest instance that is being called with
     */
    responseType(responseType) {
        this._responseType = responseType;
        return this;
    }
    /**
     * @public
     * To add properties for select OData Query param
     * @param {string|string[]} properties - The Properties value
     * @returns The same GraphRequest instance that is being called with, after adding the properties for $select query
     */
    /*
     * Accepts .select("displayName,birthday")
     *     and .select(["displayName", "birthday"])
     *     and .select("displayName", "birthday")
     *
     */
    select(properties) {
        this.addCsvQueryParameter("$select", properties, arguments);
        return this;
    }
    /**
     * @public
     * To add properties for expand OData Query param
     * @param {string|string[]} properties - The Properties value
     * @returns The same GraphRequest instance that is being called with, after adding the properties for $expand query
     */
    expand(properties) {
        this.addCsvQueryParameter("$expand", properties, arguments);
        return this;
    }
    /**
     * @public
     * To add properties for orderby OData Query param
     * @param {string|string[]} properties - The Properties value
     * @returns The same GraphRequest instance that is being called with, after adding the properties for $orderby query
     */
    orderby(properties) {
        this.addCsvQueryParameter("$orderby", properties, arguments);
        return this;
    }
    /**
     * @public
     * To add query string for filter OData Query param. The request URL accepts only one $filter Odata Query option and its value is set to the most recently passed filter query string.
     * @param {string} filterStr - The filter query string
     * @returns The same GraphRequest instance that is being called with, after adding the $filter query
     */
    filter(filterStr) {
        this.urlComponents.oDataQueryParams.$filter = filterStr;
        return this;
    }
    /**
     * @public
     * To add criterion for search OData Query param. The request URL accepts only one $search Odata Query option and its value is set to the most recently passed search criterion string.
     * @param {string} searchStr - The search criterion string
     * @returns The same GraphRequest instance that is being called with, after adding the $search query criteria
     */
    search(searchStr) {
        this.urlComponents.oDataQueryParams.$search = searchStr;
        return this;
    }
    /**
     * @public
     * To add number for top OData Query param. The request URL accepts only one $top Odata Query option and its value is set to the most recently passed number value.
     * @param {number} n - The number value
     * @returns The same GraphRequest instance that is being called with, after adding the number for $top query
     */
    top(n) {
        this.urlComponents.oDataQueryParams.$top = n;
        return this;
    }
    /**
     * @public
     * To add number for skip OData Query param. The request URL accepts only one $skip Odata Query option and its value is set to the most recently passed number value.
     * @param {number} n - The number value
     * @returns The same GraphRequest instance that is being called with, after adding the number for the $skip query
     */
    skip(n) {
        this.urlComponents.oDataQueryParams.$skip = n;
        return this;
    }
    /**
     * @public
     * To add token string for skipToken OData Query param. The request URL accepts only one $skipToken Odata Query option and its value is set to the most recently passed token value.
     * @param {string} token - The token value
     * @returns The same GraphRequest instance that is being called with, after adding the token string for $skipToken query option
     */
    skipToken(token) {
        this.urlComponents.oDataQueryParams.$skipToken = token;
        return this;
    }
    /**
     * @public
     * To add boolean for count OData Query param. The URL accepts only one $count Odata Query option and its value is set to the most recently passed boolean value.
     * @param {boolean} isCount - The count boolean
     * @returns The same GraphRequest instance that is being called with, after adding the boolean value for the $count query option
     */
    count(isCount = false) {
        this.urlComponents.oDataQueryParams.$count = isCount.toString();
        return this;
    }
    /**
     * @public
     * Appends query string to the urlComponent
     * @param {string|KeyValuePairObjectStringNumber} queryDictionaryOrString - The query value
     * @returns The same GraphRequest instance that is being called with, after appending the query string to the url component
     */
    /*
     * Accepts .query("displayName=xyz")
     *     and .select({ name: "value" })
     */
    query(queryDictionaryOrString) {
        return this.parseQueryParameter(queryDictionaryOrString);
    }
    /**
     * @public
     * @async
     * Makes a http request with GET method
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the get response
     */
    get(callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            const url = this.buildFullUrl();
            const options = {
                method: _RequestMethod__WEBPACK_IMPORTED_MODULE_4__.RequestMethod.GET,
            };
            try {
                const response = yield this.send(url, options, callback);
                return response;
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Makes a http request with POST method
     * @param {any} content - The content that needs to be sent with the request
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the post response
     */
    post(content, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            const url = this.buildFullUrl();
            const options = {
                method: _RequestMethod__WEBPACK_IMPORTED_MODULE_4__.RequestMethod.POST,
                body: (0,_GraphRequestUtil__WEBPACK_IMPORTED_MODULE_1__.serializeContent)(content),
            };
            const className = content && content.constructor && content.constructor.name;
            if (className === "FormData") {
                // Content-Type headers should not be specified in case the of FormData type content
                options.headers = {};
            }
            else {
                this.setHeaderContentType();
                options.headers = this._headers;
            }
            try {
                const response = yield this.send(url, options, callback);
                return response;
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Alias for Post request call
     * @param {any} content - The content that needs to be sent with the request
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the post response
     */
    create(content, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            try {
                return yield this.post(content, callback);
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Makes http request with PUT method
     * @param {any} content - The content that needs to be sent with the request
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the put response
     */
    put(content, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            const url = this.buildFullUrl();
            this.setHeaderContentType();
            const options = {
                method: _RequestMethod__WEBPACK_IMPORTED_MODULE_4__.RequestMethod.PUT,
                body: (0,_GraphRequestUtil__WEBPACK_IMPORTED_MODULE_1__.serializeContent)(content),
            };
            try {
                const response = yield this.send(url, options, callback);
                return response;
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Makes http request with PATCH method
     * @param {any} content - The content that needs to be sent with the request
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the patch response
     */
    patch(content, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            const url = this.buildFullUrl();
            this.setHeaderContentType();
            const options = {
                method: _RequestMethod__WEBPACK_IMPORTED_MODULE_4__.RequestMethod.PATCH,
                body: (0,_GraphRequestUtil__WEBPACK_IMPORTED_MODULE_1__.serializeContent)(content),
            };
            try {
                const response = yield this.send(url, options, callback);
                return response;
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Alias for PATCH request
     * @param {any} content - The content that needs to be sent with the request
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the patch response
     */
    update(content, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            try {
                return yield this.patch(content, callback);
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Makes http request with DELETE method
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the delete response
     */
    delete(callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            const url = this.buildFullUrl();
            const options = {
                method: _RequestMethod__WEBPACK_IMPORTED_MODULE_4__.RequestMethod.DELETE,
            };
            try {
                const response = yield this.send(url, options, callback);
                return response;
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Alias for delete request call
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the delete response
     */
    del(callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            try {
                return yield this.delete(callback);
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Makes a http request with GET method to read response as a stream.
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the getStream response
     */
    getStream(callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            const url = this.buildFullUrl();
            const options = {
                method: _RequestMethod__WEBPACK_IMPORTED_MODULE_4__.RequestMethod.GET,
            };
            this.responseType(_ResponseType__WEBPACK_IMPORTED_MODULE_5__.ResponseType.STREAM);
            try {
                const stream = yield this.send(url, options, callback);
                return stream;
            }
            catch (error) {
                throw error;
            }
        });
    }
    /**
     * @public
     * @async
     * Makes a http request with GET method to read response as a stream.
     * @param {any} stream - The stream instance
     * @param {GraphRequestCallback} [callback] - The callback function to be called in response with async call
     * @returns A promise that resolves to the putStream response
     */
    putStream(stream, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_6__.__awaiter(this, void 0, void 0, function* () {
            const url = this.buildFullUrl();
            const options = {
                method: _RequestMethod__WEBPACK_IMPORTED_MODULE_4__.RequestMethod.PUT,
                headers: {
                    "Content-Type": "application/octet-stream",
                },
                body: stream,
            };
            try {
                const response = yield this.send(url, options, callback);
                return response;
            }
            catch (error) {
                throw error;
            }
        });
    }
}
//# sourceMappingURL=GraphRequest.js.map

/***/ }),
/* 35 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "GraphErrorHandler": () => (/* binding */ GraphErrorHandler)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(6);
/* harmony import */ var _GraphError__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(36);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @module GraphErrorHandler
 */

/**
 * @class
 * Class for GraphErrorHandler
 */
class GraphErrorHandler {
    /**
     * @private
     * @static
     * Populates the GraphError instance with Error instance values
     * @param {Error} error - The error returned by graph service or some native error
     * @param {number} [statusCode] - The status code of the response
     * @returns The GraphError instance
     */
    static constructError(error, statusCode) {
        const gError = new _GraphError__WEBPACK_IMPORTED_MODULE_0__.GraphError(statusCode, "", error);
        if (error.name !== undefined) {
            gError.code = error.name;
        }
        gError.body = error.toString();
        gError.date = new Date();
        return gError;
    }
    /**
     * @private
     * @static
     * @async
     * Populates the GraphError instance from the Error returned by graph service
     * @param {any} error - The error returned by graph service or some native error
     * @param {number} statusCode - The status code of the response
     * @returns A promise that resolves to GraphError instance
     *
     * Example error for https://graph.microsoft.com/v1.0/me/events?$top=3&$search=foo
     * {
     *      "error": {
     *          "code": "SearchEvents",
     *          "message": "The parameter $search is not currently supported on the Events resource.",
     *          "innerError": {
     *              "request-id": "b31c83fd-944c-4663-aa50-5d9ceb367e19",
     *              "date": "2016-11-17T18:37:45"
     *          }
     *      }
     *  }
     */
    static constructErrorFromResponse(error, statusCode) {
        error = error.error;
        const gError = new _GraphError__WEBPACK_IMPORTED_MODULE_0__.GraphError(statusCode, error.message);
        gError.code = error.code;
        if (error.innerError !== undefined) {
            gError.requestId = error.innerError["request-id"];
            gError.date = new Date(error.innerError.date);
        }
        try {
            gError.body = JSON.stringify(error);
        }
        catch (error) {
            // tslint:disable-line: no-empty
        }
        return gError;
    }
    /**
     * @public
     * @static
     * @async
     * To get the GraphError object
     * @param {any} [error = null] - The error returned by graph service or some native error
     * @param {number} [statusCode = -1] - The status code of the response
     * @param {GraphRequestCallback} [callback] - The graph request callback function
     * @returns A promise that resolves to GraphError instance
     */
    static getError(error = null, statusCode = -1, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            let gError;
            if (error && error.error) {
                gError = GraphErrorHandler.constructErrorFromResponse(error, statusCode);
            }
            else if (typeof Error !== "undefined" && error instanceof Error) {
                gError = GraphErrorHandler.constructError(error, statusCode);
            }
            else {
                gError = new _GraphError__WEBPACK_IMPORTED_MODULE_0__.GraphError(statusCode);
            }
            if (typeof callback === "function") {
                callback(gError, null);
            }
            else {
                return gError;
            }
        });
    }
}
//# sourceMappingURL=GraphErrorHandler.js.map

/***/ }),
/* 36 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "GraphError": () => (/* binding */ GraphError)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module GraphError
 */
/**
 * @class
 * Class for GraphError
 * @NOTE: This is NOT what is returned from the Graph
 * GraphError is created from parsing JSON errors returned from the graph
 * Some fields are renamed ie, "request-id" => requestId so you can use dot notation
 */
class GraphError extends Error {
    /**
     * @public
     * @constructor
     * Creates an instance of GraphError
     * @param {number} [statusCode = -1] - The status code of the error
     * @returns An instance of GraphError
     */
    constructor(statusCode = -1, message, baseError) {
        super(message || (baseError && baseError.message));
        // https://github.com/Microsoft/TypeScript/wiki/Breaking-Changes#extending-built-ins-like-error-array-and-map-may-no-longer-work
        Object.setPrototypeOf(this, GraphError.prototype);
        this.statusCode = statusCode;
        this.code = null;
        this.requestId = null;
        this.date = new Date();
        this.body = null;
        this.stack = baseError ? baseError.stack : this.stack;
    }
}
//# sourceMappingURL=GraphError.js.map

/***/ }),
/* 37 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "DocumentType": () => (/* binding */ DocumentType),
/* harmony export */   "GraphResponseHandler": () => (/* binding */ GraphResponseHandler)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(6);
/* harmony import */ var _ResponseType__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(38);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */


/**
 * @enum
 * Enum for document types
 * @property {string} TEXT_HTML - The text/html content type
 * @property {string} TEXT_XML - The text/xml content type
 * @property {string} APPLICATION_XML - The application/xml content type
 * @property {string} APPLICATION_XHTML - The application/xhml+xml content type
 */
var DocumentType;
(function (DocumentType) {
    DocumentType["TEXT_HTML"] = "text/html";
    DocumentType["TEXT_XML"] = "text/xml";
    DocumentType["APPLICATION_XML"] = "application/xml";
    DocumentType["APPLICATION_XHTML"] = "application/xhtml+xml";
})(DocumentType || (DocumentType = {}));
/**
 * @enum
 * Enum for Content types
 * @property {string} TEXT_PLAIN - The text/plain content type
 * @property {string} APPLICATION_JSON - The application/json content type
 */
var ContentType;
(function (ContentType) {
    ContentType["TEXT_PLAIN"] = "text/plain";
    ContentType["APPLICATION_JSON"] = "application/json";
})(ContentType || (ContentType = {}));
/**
 * @enum
 * Enum for Content type regex
 * @property {string} DOCUMENT - The regex to match document content types
 * @property {string} IMAGE - The regex to match image content types
 */
var ContentTypeRegexStr;
(function (ContentTypeRegexStr) {
    ContentTypeRegexStr["DOCUMENT"] = "^(text\\/(html|xml))|(application\\/(xml|xhtml\\+xml))$";
    ContentTypeRegexStr["IMAGE"] = "^image\\/.+";
})(ContentTypeRegexStr || (ContentTypeRegexStr = {}));
/**
 * @class
 * Class for GraphResponseHandler
 */
class GraphResponseHandler {
    /**
     * @private
     * @static
     * To parse Document response
     * @param {Response} rawResponse - The response object
     * @param {DocumentType} type - The type to which the document needs to be parsed
     * @returns A promise that resolves to a document content
     */
    static parseDocumentResponse(rawResponse, type) {
        try {
            if (typeof DOMParser !== "undefined") {
                return new Promise((resolve, reject) => {
                    rawResponse.text().then((xmlString) => {
                        try {
                            const parser = new DOMParser();
                            const xmlDoc = parser.parseFromString(xmlString, type);
                            resolve(xmlDoc);
                        }
                        catch (error) {
                            reject(error);
                        }
                    });
                });
            }
            else {
                return Promise.resolve(rawResponse.body);
            }
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * @private
     * @static
     * @async
     * To convert the native Response to response content
     * @param {Response} rawResponse - The response object
     * @param {ResponseType} [responseType] - The response type value
     * @returns A promise that resolves to the converted response content
     */
    static convertResponse(rawResponse, responseType) {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            if (rawResponse.status === 204) {
                // NO CONTENT
                return Promise.resolve();
            }
            let responseValue;
            try {
                switch (responseType) {
                    case _ResponseType__WEBPACK_IMPORTED_MODULE_0__.ResponseType.ARRAYBUFFER:
                        responseValue = yield rawResponse.arrayBuffer();
                        break;
                    case _ResponseType__WEBPACK_IMPORTED_MODULE_0__.ResponseType.BLOB:
                        responseValue = yield rawResponse.blob();
                        break;
                    case _ResponseType__WEBPACK_IMPORTED_MODULE_0__.ResponseType.DOCUMENT:
                        responseValue = yield GraphResponseHandler.parseDocumentResponse(rawResponse, DocumentType.TEXT_XML);
                        break;
                    case _ResponseType__WEBPACK_IMPORTED_MODULE_0__.ResponseType.JSON:
                        responseValue = yield rawResponse.json();
                        break;
                    case _ResponseType__WEBPACK_IMPORTED_MODULE_0__.ResponseType.STREAM:
                        responseValue = yield Promise.resolve(rawResponse.body);
                        break;
                    case _ResponseType__WEBPACK_IMPORTED_MODULE_0__.ResponseType.TEXT:
                        responseValue = yield rawResponse.text();
                        break;
                    default:
                        const contentType = rawResponse.headers.get("Content-type");
                        if (contentType !== null) {
                            const mimeType = contentType.split(";")[0];
                            if (new RegExp(ContentTypeRegexStr.DOCUMENT).test(mimeType)) {
                                responseValue = yield GraphResponseHandler.parseDocumentResponse(rawResponse, mimeType);
                            }
                            else if (new RegExp(ContentTypeRegexStr.IMAGE).test(mimeType)) {
                                responseValue = rawResponse.blob();
                            }
                            else if (mimeType === ContentType.TEXT_PLAIN) {
                                responseValue = yield rawResponse.text();
                            }
                            else if (mimeType === ContentType.APPLICATION_JSON) {
                                responseValue = yield rawResponse.json();
                            }
                            else {
                                responseValue = Promise.resolve(rawResponse.body);
                            }
                        }
                        else {
                            /**
                             * RFC specification {@link https://tools.ietf.org/html/rfc7231#section-3.1.1.5} says:
                             *  A sender that generates a message containing a payload body SHOULD
                             *  generate a Content-Type header field in that message unless the
                             *  intended media type of the enclosed representation is unknown to the
                             *  sender.  If a Content-Type header field is not present, the recipient
                             *  MAY either assume a media type of "application/octet-stream"
                             *  ([RFC2046], Section 4.5.1) or examine the data to determine its type.
                             *
                             *  So assuming it as a stream type so returning the body.
                             */
                            responseValue = Promise.resolve(rawResponse.body);
                        }
                        break;
                }
            }
            catch (error) {
                throw error;
            }
            return responseValue;
        });
    }
    /**
     * @public
     * @static
     * @async
     * To get the parsed response
     * @param {Response} rawResponse - The response object
     * @param {ResponseType} [responseType] - The response type value
     * @param {GraphRequestCallback} [callback] - The graph request callback function
     * @returns The parsed response
     */
    static getResponse(rawResponse, responseType, callback) {
        return tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter(this, void 0, void 0, function* () {
            try {
                if (responseType === _ResponseType__WEBPACK_IMPORTED_MODULE_0__.ResponseType.RAW) {
                    return Promise.resolve(rawResponse);
                }
                else {
                    const response = yield GraphResponseHandler.convertResponse(rawResponse, responseType);
                    if (rawResponse.ok) {
                        // Status Code 2XX
                        if (typeof callback === "function") {
                            callback(null, response);
                        }
                        else {
                            return response;
                        }
                    }
                    else {
                        // NOT OK Response
                        throw response;
                    }
                }
            }
            catch (error) {
                throw error;
            }
        });
    }
}
//# sourceMappingURL=GraphResponseHandler.js.map

/***/ }),
/* 38 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "ResponseType": () => (/* binding */ ResponseType)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @enum
 * Enum for ResponseType values
 * @property {string} ARRAYBUFFER - To download response content as an [ArrayBuffer]{@link https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/ArrayBuffer}
 * @property {string} BLOB - To download content as a [binary/blob] {@link https://developer.mozilla.org/en-US/docs/Web/API/Blob}
 * @property {string} DOCUMENT - This downloads content as a document or stream
 * @property {string} JSON - To download response content as a json
 * @property {string} STREAM - To download response as a [stream]{@link https://nodejs.org/api/stream.html}
 * @property {string} TEXT - For downloading response as a text
 */
var ResponseType;
(function (ResponseType) {
    ResponseType["ARRAYBUFFER"] = "arraybuffer";
    ResponseType["BLOB"] = "blob";
    ResponseType["DOCUMENT"] = "document";
    ResponseType["JSON"] = "json";
    ResponseType["RAW"] = "raw";
    ResponseType["STREAM"] = "stream";
    ResponseType["TEXT"] = "text";
})(ResponseType || (ResponseType = {}));
//# sourceMappingURL=ResponseType.js.map

/***/ }),
/* 39 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "HTTPClient": () => (/* binding */ HTTPClient)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @class
 * Class representing HTTPClient
 */
class HTTPClient {
    /**
     * @public
     * @constructor
     * Creates an instance of a HTTPClient
     * @param {...Middleware} middleware - The first middleware of the middleware chain or a sequence of all the Middleware handlers
     */
    constructor(...middleware) {
        if (!middleware || !middleware.length) {
            const error = new Error();
            error.name = "InvalidMiddlewareChain";
            error.message = "Please provide a default middleware chain or custom middleware chain";
            throw error;
        }
        this.setMiddleware(...middleware);
    }
    /**
     * @private
     * Processes the middleware parameter passed to set this.middleware property
     * The calling function should validate if middleware is not undefined or not empty.
     * @param {...Middleware} middleware - The middleware passed
     * @returns Nothing
     */
    setMiddleware(...middleware) {
        if (middleware.length > 1) {
            this.parseMiddleWareArray(middleware);
        }
        else {
            this.middleware = middleware[0];
        }
    }
    /**
     * @private
     * Processes the middleware array to construct the chain
     * and sets this.middleware property to the first middlware handler of the array
     * The calling function should validate if middleware is not undefined or not empty
     * @param {Middleware[]} middlewareArray - The array of middleware handlers
     * @returns Nothing
     */
    parseMiddleWareArray(middlewareArray) {
        middlewareArray.forEach((element, index) => {
            if (index < middlewareArray.length - 1) {
                element.setNext(middlewareArray[index + 1]);
            }
        });
        this.middleware = middlewareArray[0];
    }
    /**
     * @public
     * @async
     * To send the request through the middleware chain
     * @param {Context} context - The context of a request
     * @returns A promise that resolves to the Context
     */
    sendRequest(context) {
        return tslib__WEBPACK_IMPORTED_MODULE_0__.__awaiter(this, void 0, void 0, function* () {
            try {
                if (typeof context.request === "string" && context.options === undefined) {
                    const error = new Error();
                    error.name = "InvalidRequestOptions";
                    error.message = "Unable to execute the middleware, Please provide valid options for a request";
                    throw error;
                }
                yield this.middleware.execute(context);
                return context;
            }
            catch (error) {
                throw error;
            }
        });
    }
}
//# sourceMappingURL=HTTPClient.js.map

/***/ }),
/* 40 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "HTTPClientFactory": () => (/* binding */ HTTPClientFactory)
/* harmony export */ });
/* harmony import */ var _HTTPClient__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(39);
/* harmony import */ var _middleware_AuthenticationHandler__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(8);
/* harmony import */ var _middleware_HTTPMessageHandler__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(13);
/* harmony import */ var _middleware_options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(17);
/* harmony import */ var _middleware_options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(15);
/* harmony import */ var _middleware_RedirectHandler__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(16);
/* harmony import */ var _middleware_RetryHandler__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(14);
/* harmony import */ var _middleware_TelemetryHandler__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(18);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @module HTTPClientFactory
 */








/**
 * @private
 * To check whether the environment is node or not
 * @returns A boolean representing the environment is node or not
 */
const isNodeEnvironment = () => {
    return typeof process === "object" && "function" === "function";
};
/**
 * @class
 * Class representing HTTPClientFactory
 */
class HTTPClientFactory {
    /**
     * @public
     * @static
     * Creates HTTPClient with default middleware chain
     * @param {AuthenticationProvider} authProvider - The authentication provider instance
     * @returns A HTTPClient instance
     *
     * NOTE: These are the things that we need to remember while doing modifications in the below default pipeline.
     * 		* HTTPMessageHander should be the last one in the middleware pipeline, because this makes the actual network call of the request
     * 		* TelemetryHandler should be the one prior to the last middleware in the chain, because this is the one which actually collects and appends the usage flag and placing this handler 	*		  before making the actual network call ensures that the usage of all features are recorded in the flag.
     * 		* The best place for AuthenticationHandler is in the starting of the pipeline, because every other handler might have to work for multiple times for a request but the auth token for
     * 		  them will remain same. For example, Retry and Redirect handlers might be working multiple times for a request based on the response but their auth token would remain same.
     */
    static createWithAuthenticationProvider(authProvider) {
        const authenticationHandler = new _middleware_AuthenticationHandler__WEBPACK_IMPORTED_MODULE_1__.AuthenticationHandler(authProvider);
        const retryHandler = new _middleware_RetryHandler__WEBPACK_IMPORTED_MODULE_6__.RetryHandler(new _middleware_options_RetryHandlerOptions__WEBPACK_IMPORTED_MODULE_4__.RetryHandlerOptions());
        const telemetryHandler = new _middleware_TelemetryHandler__WEBPACK_IMPORTED_MODULE_7__.TelemetryHandler();
        const httpMessageHandler = new _middleware_HTTPMessageHandler__WEBPACK_IMPORTED_MODULE_2__.HTTPMessageHandler();
        authenticationHandler.setNext(retryHandler);
        if (isNodeEnvironment()) {
            const redirectHandler = new _middleware_RedirectHandler__WEBPACK_IMPORTED_MODULE_5__.RedirectHandler(new _middleware_options_RedirectHandlerOptions__WEBPACK_IMPORTED_MODULE_3__.RedirectHandlerOptions());
            retryHandler.setNext(redirectHandler);
            redirectHandler.setNext(telemetryHandler);
        }
        else {
            retryHandler.setNext(telemetryHandler);
        }
        telemetryHandler.setNext(httpMessageHandler);
        return HTTPClientFactory.createWithMiddleware(authenticationHandler);
    }
    /**
     * @public
     * @static
     * Creates a middleware chain with the given one
     * @property {...Middleware} middleware - The first middleware of the middleware chain or a sequence of all the Middleware handlers
     * @returns A HTTPClient instance
     */
    static createWithMiddleware(...middleware) {
        // Middleware should not empty or undefined. This is check is present in the HTTPClient constructor.
        return new _HTTPClient__WEBPACK_IMPORTED_MODULE_0__.HTTPClient(...middleware);
    }
}
//# sourceMappingURL=HTTPClientFactory.js.map

/***/ }),
/* 41 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "validatePolyFilling": () => (/* binding */ validatePolyFilling)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @constant
 * @function
 * Validates availability of Promise and fetch in global context
 * @returns The true in case the Promise and fetch available, otherwise throws error
 */
const validatePolyFilling = () => {
    if (typeof Promise === "undefined" && typeof fetch === "undefined") {
        const error = new Error("Library cannot function without Promise and fetch. So, please provide polyfill for them.");
        error.name = "PolyFillNotAvailable";
        throw error;
    }
    else if (typeof Promise === "undefined") {
        const error = new Error("Library cannot function without Promise. So, please provide polyfill for it.");
        error.name = "PolyFillNotAvailable";
        throw error;
    }
    else if (typeof fetch === "undefined") {
        const error = new Error("Library cannot function without fetch. So, please provide polyfill for it.");
        error.name = "PolyFillNotAvailable";
        throw error;
    }
    return true;
};
//# sourceMappingURL=ValidatePolyFilling.js.map

/***/ }),
/* 42 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "ImplicitMSALAuthenticationProvider": () => (/* binding */ ImplicitMSALAuthenticationProvider)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/* harmony import */ var msal__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(43);
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * @module ImplicitMSALAuthenticationProvider
 */

/**
 * @class
 * Class representing ImplicitMSALAuthenticationProvider
 * @extends AuthenticationProvider
 */
class ImplicitMSALAuthenticationProvider {
    /**
     * @public
     * @constructor
     * Creates an instance of ImplicitMSALAuthenticationProvider
     * @param {UserAgentApplication} msalApplication - An instance of MSAL UserAgentApplication
     * @param {MSALAuthenticationProviderOptions} options - An instance of MSALAuthenticationProviderOptions
     * @returns An instance of ImplicitMSALAuthenticationProvider
     */
    constructor(msalApplication, options) {
        this.options = options;
        this.msalApplication = msalApplication;
    }
    /**
     * @public
     * @async
     * To get the access token
     * @param {AuthenticationProviderOptions} authenticationProviderOptions - The authentication provider options object
     * @returns The promise that resolves to an access token
     */
    getAccessToken(authenticationProviderOptions) {
        return tslib__WEBPACK_IMPORTED_MODULE_0__.__awaiter(this, void 0, void 0, function* () {
            const options = authenticationProviderOptions;
            let scopes;
            if (typeof options !== "undefined") {
                scopes = options.scopes;
            }
            if (typeof scopes === "undefined" || scopes.length === 0) {
                scopes = this.options.scopes;
            }
            if (scopes.length === 0) {
                const error = new Error();
                error.name = "EmptyScopes";
                error.message = "Scopes cannot be empty, Please provide a scopes";
                throw error;
            }
            if (this.msalApplication.getAccount()) {
                const tokenRequest = {
                    scopes,
                };
                try {
                    const authResponse = yield this.msalApplication.acquireTokenSilent(tokenRequest);
                    return authResponse.accessToken;
                }
                catch (error) {
                    if (error instanceof msal__WEBPACK_IMPORTED_MODULE_1__.InteractionRequiredAuthError) {
                        try {
                            const authResponse = yield this.msalApplication.acquireTokenPopup(tokenRequest);
                            return authResponse.accessToken;
                        }
                        catch (error) {
                            throw error;
                        }
                    }
                    else {
                        throw error;
                    }
                }
            }
            else {
                try {
                    const tokenRequest = {
                        scopes,
                    };
                    yield this.msalApplication.loginPopup(tokenRequest);
                    const authResponse = yield this.msalApplication.acquireTokenSilent(tokenRequest);
                    return authResponse.accessToken;
                }
                catch (error) {
                    throw error;
                }
            }
        });
    }
}
//# sourceMappingURL=ImplicitMSALAuthenticationProvider.js.map

/***/ }),
/* 43 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "InteractionRequiredAuthErrorMessage": () => (/* binding */ InteractionRequiredAuthErrorMessage),
/* harmony export */   "InteractionRequiredAuthError": () => (/* binding */ InteractionRequiredAuthError)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/* harmony import */ var _ServerError__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(44);
/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */


var InteractionRequiredAuthErrorMessage = {
    interactionRequired: {
        code: "interaction_required"
    },
    consentRequired: {
        code: "consent_required"
    },
    loginRequired: {
        code: "login_required"
    },
};
/**
 * Error thrown when the user is required to perform an interactive token request.
 */
var InteractionRequiredAuthError = /** @class */ (function (_super) {
    tslib__WEBPACK_IMPORTED_MODULE_0__.__extends(InteractionRequiredAuthError, _super);
    function InteractionRequiredAuthError(errorCode, errorMessage) {
        var _this = _super.call(this, errorCode, errorMessage) || this;
        _this.name = "InteractionRequiredAuthError";
        Object.setPrototypeOf(_this, InteractionRequiredAuthError.prototype);
        return _this;
    }
    InteractionRequiredAuthError.isInteractionRequiredError = function (errorString) {
        var interactionRequiredCodes = [
            InteractionRequiredAuthErrorMessage.interactionRequired.code,
            InteractionRequiredAuthErrorMessage.consentRequired.code,
            InteractionRequiredAuthErrorMessage.loginRequired.code
        ];
        return errorString && interactionRequiredCodes.indexOf(errorString) > -1;
    };
    InteractionRequiredAuthError.createLoginRequiredAuthError = function (errorDesc) {
        return new InteractionRequiredAuthError(InteractionRequiredAuthErrorMessage.loginRequired.code, errorDesc);
    };
    InteractionRequiredAuthError.createInteractionRequiredAuthError = function (errorDesc) {
        return new InteractionRequiredAuthError(InteractionRequiredAuthErrorMessage.interactionRequired.code, errorDesc);
    };
    InteractionRequiredAuthError.createConsentRequiredAuthError = function (errorDesc) {
        return new InteractionRequiredAuthError(InteractionRequiredAuthErrorMessage.consentRequired.code, errorDesc);
    };
    return InteractionRequiredAuthError;
}(_ServerError__WEBPACK_IMPORTED_MODULE_1__.ServerError));

//# sourceMappingURL=InteractionRequiredAuthError.js.map

/***/ }),
/* 44 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "ServerErrorMessage": () => (/* binding */ ServerErrorMessage),
/* harmony export */   "ServerError": () => (/* binding */ ServerError)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/* harmony import */ var _AuthError__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(45);
/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */


var ServerErrorMessage = {
    serverUnavailable: {
        code: "server_unavailable",
        desc: "Server is temporarily unavailable."
    },
    unknownServerError: {
        code: "unknown_server_error"
    },
};
/**
 * Error thrown when there is an error with the server code, for example, unavailability.
 */
var ServerError = /** @class */ (function (_super) {
    tslib__WEBPACK_IMPORTED_MODULE_0__.__extends(ServerError, _super);
    function ServerError(errorCode, errorMessage) {
        var _this = _super.call(this, errorCode, errorMessage) || this;
        _this.name = "ServerError";
        Object.setPrototypeOf(_this, ServerError.prototype);
        return _this;
    }
    ServerError.createServerUnavailableError = function () {
        return new ServerError(ServerErrorMessage.serverUnavailable.code, ServerErrorMessage.serverUnavailable.desc);
    };
    ServerError.createUnknownServerError = function (errorDesc) {
        return new ServerError(ServerErrorMessage.unknownServerError.code, errorDesc);
    };
    return ServerError;
}(_AuthError__WEBPACK_IMPORTED_MODULE_1__.AuthError));

//# sourceMappingURL=ServerError.js.map

/***/ }),
/* 45 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "AuthErrorMessage": () => (/* binding */ AuthErrorMessage),
/* harmony export */   "AuthError": () => (/* binding */ AuthError)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

var AuthErrorMessage = {
    unexpectedError: {
        code: "unexpected_error",
        desc: "Unexpected error in authentication."
    },
    noWindowObjectError: {
        code: "no_window_object",
        desc: "No window object available. Details:"
    }
};
/**
 * General error class thrown by the MSAL.js library.
 */
var AuthError = /** @class */ (function (_super) {
    tslib__WEBPACK_IMPORTED_MODULE_0__.__extends(AuthError, _super);
    function AuthError(errorCode, errorMessage) {
        var _this = _super.call(this, errorMessage) || this;
        Object.setPrototypeOf(_this, AuthError.prototype);
        _this.errorCode = errorCode;
        _this.errorMessage = errorMessage;
        _this.name = "AuthError";
        return _this;
    }
    AuthError.createUnexpectedError = function (errDesc) {
        return new AuthError(AuthErrorMessage.unexpectedError.code, AuthErrorMessage.unexpectedError.desc + ": " + errDesc);
    };
    AuthError.createNoWindowObjectError = function (errDesc) {
        return new AuthError(AuthErrorMessage.noWindowObjectError.code, AuthErrorMessage.noWindowObjectError.desc + " " + errDesc);
    };
    return AuthError;
}(Error));

//# sourceMappingURL=AuthError.js.map

/***/ }),
/* 46 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "MSALAuthenticationProviderOptions": () => (/* binding */ MSALAuthenticationProviderOptions)
/* harmony export */ });
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * @class
 * @implements AuthenticationProviderOptions
 * Class representing MSALAuthenticationProviderOptions
 */
class MSALAuthenticationProviderOptions {
    /**
     * @public
     * @constructor
     * To create an instance of MSALAuthenticationProviderOptions
     * @param {string[]} scopes - An array of scopes
     * @returns An instance of MSALAuthenticationProviderOptions
     */
    constructor(scopes) {
        this.scopes = scopes;
    }
}
//# sourceMappingURL=MSALAuthenticationProviderOptions.js.map

/***/ }),
/* 47 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.MicrosoftToDoTreeDataProvider = void 0;
const vscode = __webpack_require__(1);
var TaskStatusType;
(function (TaskStatusType) {
    TaskStatusType["completed"] = "Completed";
    TaskStatusType["inProgress"] = "In Progress";
})(TaskStatusType || (TaskStatusType = {}));
class MicrosoftToDoTreeDataProvider extends vscode.Disposable {
    constructor(clientFactory) {
        super(() => this.dispose());
        this.clientFactory = clientFactory;
        this.didChangeTreeData = new vscode.EventEmitter();
        this.onDidChangeTreeData = this.didChangeTreeData.event;
        this.disposibles = [];
        this.importanceFilter = false;
        this.disposibles.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.refreshList', (element) => this.didChangeTreeData.fire(element)));
        this.disposibles.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.complete', (node, nodes) => nodes ? this.changeCompletedState(nodes) : this.changeCompletedState([node])));
        this.disposibles.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.uncomplete', (node, nodes) => nodes ? this.changeCompletedState(nodes) : this.changeCompletedState([node])));
        this.disposibles.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.star', (node, nodes) => nodes ? this.changeImportanceState(nodes) : this.changeImportanceState([node])));
        this.disposibles.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.unstar', (node, nodes) => nodes ? this.changeImportanceState(nodes) : this.changeImportanceState([node])));
        this.disposibles.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.starFilter', () => __awaiter(this, void 0, void 0, function* () {
            this.importanceFilter = true;
            yield vscode.commands.executeCommand('setContext', 'starFilter', true);
            this.didChangeTreeData.fire();
        })));
        this.disposibles.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.unstarFilter', () => __awaiter(this, void 0, void 0, function* () {
            this.importanceFilter = false;
            yield vscode.commands.executeCommand('setContext', 'starFilter', false);
            this.didChangeTreeData.fire();
        })));
    }
    changeCompletedState(nodes) {
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientFactory.getClient();
            const promises = nodes.map((n) => __awaiter(this, void 0, void 0, function* () {
                yield client.api(`/me/todo/lists/${n.parent.entity.id}/tasks/${n.entity.id}`).patch({
                    status: n.entity.status === 'completed' ? 'notStarted' : 'completed'
                });
                this.didChangeTreeData.fire(n.parent);
            }));
            // TODO: Error handling
            yield Promise.all(promises);
        });
    }
    changeImportanceState(nodes) {
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientFactory.getClient();
            const promises = nodes.map((n) => __awaiter(this, void 0, void 0, function* () {
                yield client.api(`/me/todo/lists/${n.parent.entity.id}/tasks/${n.entity.id}`).patch({
                    importance: n.entity.importance === 'high' ? 'normal' : 'high'
                });
                this.didChangeTreeData.fire(n.parent);
            }));
            // TODO: Error handling
            yield Promise.all(promises);
            if (this.importantNode) {
                this.didChangeTreeData.fire(this.importantNode);
            }
        });
    }
    getTreeItem(element) {
        var _a;
        let treeItem;
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
                const highlights = [];
                if (dueDateTime === null || dueDateTime === void 0 ? void 0 : dueDateTime.dateTime) {
                    const dueStr = " DUE " + new Date(dueDateTime.dateTime).toLocaleDateString() + ' ';
                    label += "  ";
                    highlights.push([label.length, label.length + dueStr.length]);
                    label += dueStr;
                }
                if ((_a = element.entity.body) === null || _a === void 0 ? void 0 : _a.content) {
                    tooltip += `\n\n${element.entity.body.content}`;
                }
                const treeItemLabel = {
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
    getChildren(element) {
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientFactory.getClient();
            if (!client) {
                return;
            }
            if (!element) {
                const taskLists = yield this.clientFactory.getAll(client, '/me/todo/lists');
                this.importantNode = { nodeType: 'important-list' };
                const nodes = [this.importantNode];
                taskLists.forEach(entity => nodes.push({ nodeType: 'list', entity }));
                nodes.push({ nodeType: 'create-list' });
                return nodes;
            }
            if (element.nodeType === 'list') {
                const getTasks = (getCompleted) => __awaiter(this, void 0, void 0, function* () {
                    const comparison = getCompleted ? 'eq' : 'ne';
                    let filter = `status ${comparison} 'completed'`;
                    if (this.importanceFilter) {
                        filter += ` and importance eq 'high'`;
                    }
                    const tasks = yield this.clientFactory.getAll(client, `/me/todo/lists/${element.entity.id}/tasks?$filter=${filter}`);
                    return tasks.map(entity => ({
                        nodeType: 'task',
                        entity,
                        parent: element
                    }));
                });
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
                return yield element.getChildren();
            }
            if (element.nodeType === 'important-list') {
                const filter = `status ne 'completed' and importance eq 'high'`;
                const listEntities = yield this.clientFactory.getAll(client, '/me/todo/lists');
                const entities = [];
                for (const entity of listEntities) {
                    const tasks = yield this.clientFactory.getAll(client, `/me/todo/lists/${entity.id}/tasks?$filter=${filter}`);
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
        });
    }
    dispose() {
        this.disposibles.forEach(d => d.dispose());
    }
}
exports.MicrosoftToDoTreeDataProvider = MicrosoftToDoTreeDataProvider;


/***/ }),
/* 48 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.TaskDetailsViewProvider = void 0;
const vscode = __webpack_require__(1);
const WebViewBase_1 = __webpack_require__(49);
class TaskDetailsViewProvider extends WebViewBase_1.WebviewViewBase {
    constructor(_extensionUri, clientFactory) {
        super();
        this._extensionUri = _extensionUri;
        this.clientFactory = clientFactory;
        this.viewType = 'microsoft-todo-unoffcial.taskDetailsView';
    }
    changeChosenView(node) {
        return __awaiter(this, void 0, void 0, function* () {
            this.chosenTask = node;
            yield vscode.commands.executeCommand('setContext', 'showTaskDetailsView', true);
            this.show(true);
            yield this._postMessage(node);
        });
    }
    resolveWebviewView(webviewView, context, token) {
        const _super = Object.create(null, {
            initialize: { get: () => super.initialize }
        });
        return __awaiter(this, void 0, void 0, function* () {
            this._view = webviewView;
            this._webview = webviewView.webview;
            this._webview.options = {
                // Allow scripts in the webview
                enableScripts: true,
                localResourceRoots: [
                    this._extensionUri
                ],
            };
            _super.initialize.call(this);
            this._webview.html = this.getHtmlForWebview();
            this._disposables.push(webviewView.webview.onDidReceiveMessage((message) => __awaiter(this, void 0, void 0, function* () {
                switch (message.command) {
                    case 'cancel':
                        yield vscode.commands.executeCommand('microsoft-todo-unoffcial.closeCreateTask');
                        break;
                    case 'update':
                        const client = yield this.clientFactory.getClient();
                        if (!client) {
                            return yield vscode.window.showErrorMessage("you're not logged in.");
                        }
                        const body = {
                            title: message.body.title,
                            body: {
                                content: message.body.note,
                                contentType: 'text'
                            }
                        };
                        if (message.body.dueDate) {
                            const [month, day, year] = message.body.dueDate.split('/');
                            body.dueDateTime = {
                                dateTime: `${year}-${month}-${day}T08:00:00.0000000`,
                                timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
                            };
                        }
                        else {
                            body.dueDateTime = null;
                        }
                        if (message.body.reminderDate) {
                            if (!message.body.reminderTime) {
                                vscode.window.showErrorMessage('You need to specify a time when adding a Reminder.');
                                return;
                            }
                            const [month, day, year] = message.body.reminderDate.split('/');
                            body.reminderDateTime = {
                                dateTime: `${year}-${month}-${day}T${message.body.reminderTime}:00.0000000`,
                                timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
                            };
                        }
                        else {
                            body.reminderDateTime = null;
                        }
                        // TODO: error handling
                        yield client.api(`/me/todo/lists/${message.body.listId}/tasks/${message.body.id}`).patch(body);
                        yield vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
                        break;
                }
            })));
        });
    }
    getHtmlForWebview() {
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
        const momentScript = this._webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'node_modules', 'moment', 'min', 'moment.min.js'));
        const momentTimezoneScript = this._webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'node_modules', 'moment-timezone', 'builds', 'moment-timezone-with-data-10-year-range.min.js'));
        // Use a nonce to only allow a specific script to be run.
        const nonce = WebViewBase_1.getNonce();
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
				<div class='task-reminder-form'>
					<input placeholder='Add Reminder' type='text' class='task-reminder-date' value=''/>
					<input type='hidden' class='task-reminder-time' />
				</div>
				<input placeholder='Add Due Date' type='text' class='task-duedate' value=''/>
				<label for="task-body">Note</label>
				<textarea placeholder='Add Note' class='task-body'></textarea>
				<button class='update update-task' hidden>Update</button>
				<button class='update update-cancel' hidden>Cancel</button>
				<script nonce="${nonce}" src="${tdpScript}"></script>
				<script nonce="${nonce}" src="${momentScript}"></script>
				<script nonce="${nonce}" src="${momentTimezoneScript}"></script>
				<script nonce="${nonce}" src="${scriptUri}"></script>
			</body>
			</html>`;
    }
}
exports.TaskDetailsViewProvider = TaskDetailsViewProvider;


/***/ }),
/* 49 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.WebviewViewBase = exports.WebviewBase = exports.getNonce = void 0;
const vscode = __webpack_require__(1);
function getNonce() {
    let text = '';
    const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (let i = 0; i < 32; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
}
exports.getNonce = getNonce;
class WebviewBase {
    constructor() {
        this._disposables = [];
        this._onIsReady = new vscode.EventEmitter();
        // eslint-disable-next-line @typescript-eslint/naming-convention
        this.MESSAGE_UNHANDLED = 'message not handled';
        this._waitForReady = new Promise(resolve => {
            const disposable = this._onIsReady.event(() => {
                disposable.dispose();
                resolve();
            });
        });
    }
    initialize() {
        var _a;
        (_a = this._webview) === null || _a === void 0 ? void 0 : _a.onDidReceiveMessage((message) => __awaiter(this, void 0, void 0, function* () {
            yield this._onDidReceiveMessage(message);
        }), null, this._disposables);
    }
    _onDidReceiveMessage(message) {
        return __awaiter(this, void 0, void 0, function* () {
            switch (message.command) {
                case 'ready':
                    this._onIsReady.fire();
                    return;
                default:
                    return this.MESSAGE_UNHANDLED;
            }
        });
    }
    _postMessage(message) {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            // Without the following ready check, we can end up in a state where the message handler in the webview
            // isn't ready for any of the messages we post.
            yield this._waitForReady;
            (_a = this._webview) === null || _a === void 0 ? void 0 : _a.postMessage(message);
        });
    }
    _replyMessage(originalMessage, message) {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            const reply = {
                seq: originalMessage.req,
                res: message
            };
            (_a = this._webview) === null || _a === void 0 ? void 0 : _a.postMessage(reply);
        });
    }
    _throwError(originalMessage, error) {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            const reply = {
                seq: originalMessage.req,
                err: error
            };
            (_a = this._webview) === null || _a === void 0 ? void 0 : _a.postMessage(reply);
        });
    }
    dispose() {
        this._disposables.forEach(d => d.dispose());
    }
}
exports.WebviewBase = WebviewBase;
class WebviewViewBase extends WebviewBase {
    show(preserveFocus) {
        if (this._view) {
            this._view.show(preserveFocus);
        }
        else {
            vscode.commands.executeCommand(`${this.viewType}.focus`);
        }
    }
}
exports.WebviewViewBase = WebviewViewBase;


/***/ }),
/* 50 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.TaskOperations = void 0;
const vscode = __webpack_require__(1);
class TaskOperations extends vscode.Disposable {
    constructor(clientProvider) {
        super(() => {
            this.disposables.forEach(d => d.dispose());
        });
        this.clientProvider = clientProvider;
        this.disposables = [];
        this.disposables.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.createTask', (list) => this.createTask(list)));
        this.disposables.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.deleteTask', (task, tasks) => this.deleteTask(task, tasks)));
    }
    getTask(listId, taskId) {
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientProvider.getClient();
            if (!client) {
                yield vscode.window.showErrorMessage('Not logged in');
                return;
            }
            // TODO: error handling
            return yield client.api(`/me/todo/lists/${listId}/tasks/${taskId}`).get();
        });
    }
    getTasks(listId) {
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientProvider.getClient();
            if (!client) {
                yield vscode.window.showErrorMessage('Not logged in');
                return;
            }
            // TODO: error handling
            return yield this.clientProvider.getAll(client, `/me/todo/lists/${listId}/tasks`);
        });
    }
    createTask(list) {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientProvider.getClient();
            if (!client) {
                return yield vscode.window.showErrorMessage('Please log in before creating a task.');
            }
            let listId = list === null || list === void 0 ? void 0 : list.entity.id;
            if (!listId) {
                const taskLists = new Array();
                let iterUri = '/me/todo/lists';
                while (iterUri) {
                    let res = yield client.api(iterUri).get();
                    res.value.forEach(r => taskLists.push(r));
                    iterUri = res['@odata.nextLink'];
                }
                const quickPickItems = taskLists.map(l => (Object.assign({ label: l.displayName || '' }, l)));
                quickPickItems.push({
                    label: 'Create a new list...',
                    id: 'new',
                });
                const chosen = yield vscode.window.showQuickPick(quickPickItems, {
                    canPickMany: false,
                    ignoreFocusOut: true,
                    placeHolder: 'Which list would you like to add tasks to?'
                });
                listId = (_a = chosen) === null || _a === void 0 ? void 0 : _a.id;
                if (listId === 'new') {
                    const displayName = yield vscode.window.showInputBox({
                        prompt: 'Add a List',
                        placeHolder: 'Groceries',
                        ignoreFocusOut: true
                    });
                    // The user quit the prompt
                    if (!displayName) {
                        return;
                    }
                    // TODO: Error handling
                    const res = yield client.api('/me/todo/lists').post({
                        displayName
                    });
                    listId = res.id;
                    yield vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
                }
            }
            // The user quit the prompt
            if (!listId) {
                return;
            }
            const inputBoxOptions = {
                prompt: 'Add a Task',
                placeHolder: 'Eat my veggies',
                ignoreFocusOut: true
            };
            let title = yield vscode.window.showInputBox(inputBoxOptions);
            inputBoxOptions.prompt = 'Add another Task';
            while (title) {
                // TODO: error handling
                yield client.api(`/me/todo/lists/${listId}/tasks`).post({
                    title: title,
                    body: {
                        content: '',
                        contentType: 'text'
                    }
                });
                // not awaiting on purpose
                vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
                title = yield vscode.window.showInputBox(inputBoxOptions);
            }
        });
    }
    deleteTask(task, tasks) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!task) {
                return;
            }
            const client = yield this.clientProvider.getClient();
            if (!client) {
                return yield vscode.window.showErrorMessage('Not logged in');
            }
            const expected = (tasks === null || tasks === void 0 ? void 0 : tasks.length) ? 'Delete tasks' : 'Delete task';
            const tasksFormatted = tasks ? `${tasks.map(t => t.entity.title).join('", "')}` : task.entity.title;
            const choice = yield vscode.window.showWarningMessage(`"${tasksFormatted}" will be permanently deleted. You won't be able to undo this action.`, { modal: true }, expected, 'Cancel');
            if (choice !== expected) {
                return;
            }
            tasks !== null && tasks !== void 0 ? tasks : (tasks = [task]);
            const promises = tasks.map(t => client.api(`/me/todo/lists/${t.parent.entity.id}/tasks/${t.entity.id}`).delete());
            // TODO: error handling
            yield Promise.all(promises);
            yield vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList', task.parent);
        });
    }
}
exports.TaskOperations = TaskOperations;


/***/ }),
/* 51 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.ListOperations = void 0;
const vscode = __webpack_require__(1);
class ListOperations extends vscode.Disposable {
    constructor(clientProvider) {
        super(() => {
            this.disposables.forEach(d => d.dispose());
        });
        this.clientProvider = clientProvider;
        this.disposables = [];
        this.disposables.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.createList', () => this.createList()));
        this.disposables.push(vscode.commands.registerCommand('microsoft-todo-unoffcial.deleteList', (list, lists) => this.deleteList(list, lists)));
    }
    getList(listId) {
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientProvider.getClient();
            if (!client) {
                yield vscode.window.showErrorMessage('Not logged in');
                return;
            }
            // TODO: error handling
            return (yield client.api(`/me/todo/lists/${listId}`).get()).value;
        });
    }
    getLists() {
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientProvider.getClient();
            if (!client) {
                yield vscode.window.showErrorMessage('Not logged in');
                return;
            }
            // TODO: error handling
            return yield this.clientProvider.getAll(client, `/me/todo/lists`);
        });
    }
    createList(displayName) {
        return __awaiter(this, void 0, void 0, function* () {
            const client = yield this.clientProvider.getClient();
            if (!client) {
                yield vscode.window.showErrorMessage('Not logged in');
                return;
            }
            displayName !== null && displayName !== void 0 ? displayName : (displayName = yield vscode.window.showInputBox({
                prompt: 'Add a List',
                placeHolder: 'Groceries',
                ignoreFocusOut: true
            }));
            if (!displayName) {
                return;
            }
            yield client.api('/me/todo/lists').post({
                displayName
            });
            yield vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
        });
    }
    deleteList(list, lists) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!list) {
                return;
            }
            const client = yield this.clientProvider.getClient();
            if (!client) {
                return yield vscode.window.showErrorMessage('Not logged in');
            }
            const expected = (lists === null || lists === void 0 ? void 0 : lists.length) ? 'Delete lists' : 'Delete list';
            const tasksFormatted = lists ? `${lists.map(t => t.entity.displayName).join('", "')}` : list.entity.displayName;
            const choice = yield vscode.window.showWarningMessage(`"${tasksFormatted}" will be permanently deleted. You won't be able to undo this action.`, { modal: true }, expected, 'Cancel');
            if (choice !== expected) {
                return;
            }
            lists !== null && lists !== void 0 ? lists : (lists = [list]);
            const promises = lists.map(t => client.api(`/me/todo/lists/${t.entity.id}`).delete());
            // TODO: error handling
            yield Promise.all(promises);
            yield vscode.commands.executeCommand('microsoft-todo-unoffcial.refreshList');
        });
    }
}
exports.ListOperations = ListOperations;


/***/ }),
/* 52 */
/***/ ((__unused_webpack_module, __unused_webpack_exports, __webpack_require__) => {

const fetchNode = __webpack_require__(53)
const fetch = fetchNode.fetch.bind({})

fetch.polyfill = true

if (!global.fetch) {
  global.fetch = fetch
  global.Response = fetchNode.Response
  global.Headers = fetchNode.Headers
  global.Request = fetchNode.Request
}


/***/ }),
/* 53 */
/***/ ((module, exports, __webpack_require__) => {

const nodeFetch = __webpack_require__(54)
const realFetch = nodeFetch.default || nodeFetch

const fetch = function (url, options) {
  // Support schemaless URIs on the server for parity with the browser.
  // Ex: //github.com/ -> https://github.com/
  if (/^\/\//.test(url)) {
    url = 'https:' + url
  }
  return realFetch.call(this, url, options)
}

fetch.ponyfill = true

module.exports = exports = fetch
exports.fetch = fetch
exports.Headers = nodeFetch.Headers
exports.Request = nodeFetch.Request
exports.Response = nodeFetch.Response

// Needed for TypeScript consumers without esModuleInterop.
exports.default = fetch


/***/ }),
/* 54 */
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__),
/* harmony export */   "Headers": () => (/* binding */ Headers),
/* harmony export */   "Request": () => (/* binding */ Request),
/* harmony export */   "Response": () => (/* binding */ Response),
/* harmony export */   "FetchError": () => (/* binding */ FetchError)
/* harmony export */ });
/* harmony import */ var stream__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(55);
/* harmony import */ var http__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(56);
/* harmony import */ var url__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(57);
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(58);
/* harmony import */ var zlib__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(59);






// Based on https://github.com/tmpvar/jsdom/blob/aa85b2abf07766ff7bf5c1f6daafb3726f2f2db5/lib/jsdom/living/blob.js

// fix for "Readable" isn't a named export issue
const Readable = stream__WEBPACK_IMPORTED_MODULE_0__.Readable;

const BUFFER = Symbol('buffer');
const TYPE = Symbol('type');

class Blob {
	constructor() {
		this[TYPE] = '';

		const blobParts = arguments[0];
		const options = arguments[1];

		const buffers = [];
		let size = 0;

		if (blobParts) {
			const a = blobParts;
			const length = Number(a.length);
			for (let i = 0; i < length; i++) {
				const element = a[i];
				let buffer;
				if (element instanceof Buffer) {
					buffer = element;
				} else if (ArrayBuffer.isView(element)) {
					buffer = Buffer.from(element.buffer, element.byteOffset, element.byteLength);
				} else if (element instanceof ArrayBuffer) {
					buffer = Buffer.from(element);
				} else if (element instanceof Blob) {
					buffer = element[BUFFER];
				} else {
					buffer = Buffer.from(typeof element === 'string' ? element : String(element));
				}
				size += buffer.length;
				buffers.push(buffer);
			}
		}

		this[BUFFER] = Buffer.concat(buffers);

		let type = options && options.type !== undefined && String(options.type).toLowerCase();
		if (type && !/[^\u0020-\u007E]/.test(type)) {
			this[TYPE] = type;
		}
	}
	get size() {
		return this[BUFFER].length;
	}
	get type() {
		return this[TYPE];
	}
	text() {
		return Promise.resolve(this[BUFFER].toString());
	}
	arrayBuffer() {
		const buf = this[BUFFER];
		const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
		return Promise.resolve(ab);
	}
	stream() {
		const readable = new Readable();
		readable._read = function () {};
		readable.push(this[BUFFER]);
		readable.push(null);
		return readable;
	}
	toString() {
		return '[object Blob]';
	}
	slice() {
		const size = this.size;

		const start = arguments[0];
		const end = arguments[1];
		let relativeStart, relativeEnd;
		if (start === undefined) {
			relativeStart = 0;
		} else if (start < 0) {
			relativeStart = Math.max(size + start, 0);
		} else {
			relativeStart = Math.min(start, size);
		}
		if (end === undefined) {
			relativeEnd = size;
		} else if (end < 0) {
			relativeEnd = Math.max(size + end, 0);
		} else {
			relativeEnd = Math.min(end, size);
		}
		const span = Math.max(relativeEnd - relativeStart, 0);

		const buffer = this[BUFFER];
		const slicedBuffer = buffer.slice(relativeStart, relativeStart + span);
		const blob = new Blob([], { type: arguments[2] });
		blob[BUFFER] = slicedBuffer;
		return blob;
	}
}

Object.defineProperties(Blob.prototype, {
	size: { enumerable: true },
	type: { enumerable: true },
	slice: { enumerable: true }
});

Object.defineProperty(Blob.prototype, Symbol.toStringTag, {
	value: 'Blob',
	writable: false,
	enumerable: false,
	configurable: true
});

/**
 * fetch-error.js
 *
 * FetchError interface for operational errors
 */

/**
 * Create FetchError instance
 *
 * @param   String      message      Error message for human
 * @param   String      type         Error type for machine
 * @param   String      systemError  For Node.js system error
 * @return  FetchError
 */
function FetchError(message, type, systemError) {
  Error.call(this, message);

  this.message = message;
  this.type = type;

  // when err.type is `system`, err.code contains system error code
  if (systemError) {
    this.code = this.errno = systemError.code;
  }

  // hide custom error implementation details from end-users
  Error.captureStackTrace(this, this.constructor);
}

FetchError.prototype = Object.create(Error.prototype);
FetchError.prototype.constructor = FetchError;
FetchError.prototype.name = 'FetchError';

let convert;
try {
	convert = require('encoding').convert;
} catch (e) {}

const INTERNALS = Symbol('Body internals');

// fix an issue where "PassThrough" isn't a named export for node <10
const PassThrough = stream__WEBPACK_IMPORTED_MODULE_0__.PassThrough;

/**
 * Body mixin
 *
 * Ref: https://fetch.spec.whatwg.org/#body
 *
 * @param   Stream  body  Readable stream
 * @param   Object  opts  Response options
 * @return  Void
 */
function Body(body) {
	var _this = this;

	var _ref = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {},
	    _ref$size = _ref.size;

	let size = _ref$size === undefined ? 0 : _ref$size;
	var _ref$timeout = _ref.timeout;
	let timeout = _ref$timeout === undefined ? 0 : _ref$timeout;

	if (body == null) {
		// body is undefined or null
		body = null;
	} else if (isURLSearchParams(body)) {
		// body is a URLSearchParams
		body = Buffer.from(body.toString());
	} else if (isBlob(body)) ; else if (Buffer.isBuffer(body)) ; else if (Object.prototype.toString.call(body) === '[object ArrayBuffer]') {
		// body is ArrayBuffer
		body = Buffer.from(body);
	} else if (ArrayBuffer.isView(body)) {
		// body is ArrayBufferView
		body = Buffer.from(body.buffer, body.byteOffset, body.byteLength);
	} else if (body instanceof stream__WEBPACK_IMPORTED_MODULE_0__) ; else {
		// none of the above
		// coerce to string then buffer
		body = Buffer.from(String(body));
	}
	this[INTERNALS] = {
		body,
		disturbed: false,
		error: null
	};
	this.size = size;
	this.timeout = timeout;

	if (body instanceof stream__WEBPACK_IMPORTED_MODULE_0__) {
		body.on('error', function (err) {
			const error = err.name === 'AbortError' ? err : new FetchError(`Invalid response body while trying to fetch ${_this.url}: ${err.message}`, 'system', err);
			_this[INTERNALS].error = error;
		});
	}
}

Body.prototype = {
	get body() {
		return this[INTERNALS].body;
	},

	get bodyUsed() {
		return this[INTERNALS].disturbed;
	},

	/**
  * Decode response as ArrayBuffer
  *
  * @return  Promise
  */
	arrayBuffer() {
		return consumeBody.call(this).then(function (buf) {
			return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
		});
	},

	/**
  * Return raw response as Blob
  *
  * @return Promise
  */
	blob() {
		let ct = this.headers && this.headers.get('content-type') || '';
		return consumeBody.call(this).then(function (buf) {
			return Object.assign(
			// Prevent copying
			new Blob([], {
				type: ct.toLowerCase()
			}), {
				[BUFFER]: buf
			});
		});
	},

	/**
  * Decode response as json
  *
  * @return  Promise
  */
	json() {
		var _this2 = this;

		return consumeBody.call(this).then(function (buffer) {
			try {
				return JSON.parse(buffer.toString());
			} catch (err) {
				return Body.Promise.reject(new FetchError(`invalid json response body at ${_this2.url} reason: ${err.message}`, 'invalid-json'));
			}
		});
	},

	/**
  * Decode response as text
  *
  * @return  Promise
  */
	text() {
		return consumeBody.call(this).then(function (buffer) {
			return buffer.toString();
		});
	},

	/**
  * Decode response as buffer (non-spec api)
  *
  * @return  Promise
  */
	buffer() {
		return consumeBody.call(this);
	},

	/**
  * Decode response as text, while automatically detecting the encoding and
  * trying to decode to UTF-8 (non-spec api)
  *
  * @return  Promise
  */
	textConverted() {
		var _this3 = this;

		return consumeBody.call(this).then(function (buffer) {
			return convertBody(buffer, _this3.headers);
		});
	}
};

// In browsers, all properties are enumerable.
Object.defineProperties(Body.prototype, {
	body: { enumerable: true },
	bodyUsed: { enumerable: true },
	arrayBuffer: { enumerable: true },
	blob: { enumerable: true },
	json: { enumerable: true },
	text: { enumerable: true }
});

Body.mixIn = function (proto) {
	for (const name of Object.getOwnPropertyNames(Body.prototype)) {
		// istanbul ignore else: future proof
		if (!(name in proto)) {
			const desc = Object.getOwnPropertyDescriptor(Body.prototype, name);
			Object.defineProperty(proto, name, desc);
		}
	}
};

/**
 * Consume and convert an entire Body to a Buffer.
 *
 * Ref: https://fetch.spec.whatwg.org/#concept-body-consume-body
 *
 * @return  Promise
 */
function consumeBody() {
	var _this4 = this;

	if (this[INTERNALS].disturbed) {
		return Body.Promise.reject(new TypeError(`body used already for: ${this.url}`));
	}

	this[INTERNALS].disturbed = true;

	if (this[INTERNALS].error) {
		return Body.Promise.reject(this[INTERNALS].error);
	}

	let body = this.body;

	// body is null
	if (body === null) {
		return Body.Promise.resolve(Buffer.alloc(0));
	}

	// body is blob
	if (isBlob(body)) {
		body = body.stream();
	}

	// body is buffer
	if (Buffer.isBuffer(body)) {
		return Body.Promise.resolve(body);
	}

	// istanbul ignore if: should never happen
	if (!(body instanceof stream__WEBPACK_IMPORTED_MODULE_0__)) {
		return Body.Promise.resolve(Buffer.alloc(0));
	}

	// body is stream
	// get ready to actually consume the body
	let accum = [];
	let accumBytes = 0;
	let abort = false;

	return new Body.Promise(function (resolve, reject) {
		let resTimeout;

		// allow timeout on slow response body
		if (_this4.timeout) {
			resTimeout = setTimeout(function () {
				abort = true;
				reject(new FetchError(`Response timeout while trying to fetch ${_this4.url} (over ${_this4.timeout}ms)`, 'body-timeout'));
			}, _this4.timeout);
		}

		// handle stream errors
		body.on('error', function (err) {
			if (err.name === 'AbortError') {
				// if the request was aborted, reject with this Error
				abort = true;
				reject(err);
			} else {
				// other errors, such as incorrect content-encoding
				reject(new FetchError(`Invalid response body while trying to fetch ${_this4.url}: ${err.message}`, 'system', err));
			}
		});

		body.on('data', function (chunk) {
			if (abort || chunk === null) {
				return;
			}

			if (_this4.size && accumBytes + chunk.length > _this4.size) {
				abort = true;
				reject(new FetchError(`content size at ${_this4.url} over limit: ${_this4.size}`, 'max-size'));
				return;
			}

			accumBytes += chunk.length;
			accum.push(chunk);
		});

		body.on('end', function () {
			if (abort) {
				return;
			}

			clearTimeout(resTimeout);

			try {
				resolve(Buffer.concat(accum, accumBytes));
			} catch (err) {
				// handle streams that have accumulated too much data (issue #414)
				reject(new FetchError(`Could not create Buffer from response body for ${_this4.url}: ${err.message}`, 'system', err));
			}
		});
	});
}

/**
 * Detect buffer encoding and convert to target encoding
 * ref: http://www.w3.org/TR/2011/WD-html5-20110113/parsing.html#determining-the-character-encoding
 *
 * @param   Buffer  buffer    Incoming buffer
 * @param   String  encoding  Target encoding
 * @return  String
 */
function convertBody(buffer, headers) {
	if (typeof convert !== 'function') {
		throw new Error('The package `encoding` must be installed to use the textConverted() function');
	}

	const ct = headers.get('content-type');
	let charset = 'utf-8';
	let res, str;

	// header
	if (ct) {
		res = /charset=([^;]*)/i.exec(ct);
	}

	// no charset in content type, peek at response body for at most 1024 bytes
	str = buffer.slice(0, 1024).toString();

	// html5
	if (!res && str) {
		res = /<meta.+?charset=(['"])(.+?)\1/i.exec(str);
	}

	// html4
	if (!res && str) {
		res = /<meta[\s]+?http-equiv=(['"])content-type\1[\s]+?content=(['"])(.+?)\2/i.exec(str);
		if (!res) {
			res = /<meta[\s]+?content=(['"])(.+?)\1[\s]+?http-equiv=(['"])content-type\3/i.exec(str);
			if (res) {
				res.pop(); // drop last quote
			}
		}

		if (res) {
			res = /charset=(.*)/i.exec(res.pop());
		}
	}

	// xml
	if (!res && str) {
		res = /<\?xml.+?encoding=(['"])(.+?)\1/i.exec(str);
	}

	// found charset
	if (res) {
		charset = res.pop();

		// prevent decode issues when sites use incorrect encoding
		// ref: https://hsivonen.fi/encoding-menu/
		if (charset === 'gb2312' || charset === 'gbk') {
			charset = 'gb18030';
		}
	}

	// turn raw buffers into a single utf-8 buffer
	return convert(buffer, 'UTF-8', charset).toString();
}

/**
 * Detect a URLSearchParams object
 * ref: https://github.com/bitinn/node-fetch/issues/296#issuecomment-307598143
 *
 * @param   Object  obj     Object to detect by type or brand
 * @return  String
 */
function isURLSearchParams(obj) {
	// Duck-typing as a necessary condition.
	if (typeof obj !== 'object' || typeof obj.append !== 'function' || typeof obj.delete !== 'function' || typeof obj.get !== 'function' || typeof obj.getAll !== 'function' || typeof obj.has !== 'function' || typeof obj.set !== 'function') {
		return false;
	}

	// Brand-checking and more duck-typing as optional condition.
	return obj.constructor.name === 'URLSearchParams' || Object.prototype.toString.call(obj) === '[object URLSearchParams]' || typeof obj.sort === 'function';
}

/**
 * Check if `obj` is a W3C `Blob` object (which `File` inherits from)
 * @param  {*} obj
 * @return {boolean}
 */
function isBlob(obj) {
	return typeof obj === 'object' && typeof obj.arrayBuffer === 'function' && typeof obj.type === 'string' && typeof obj.stream === 'function' && typeof obj.constructor === 'function' && typeof obj.constructor.name === 'string' && /^(Blob|File)$/.test(obj.constructor.name) && /^(Blob|File)$/.test(obj[Symbol.toStringTag]);
}

/**
 * Clone body given Res/Req instance
 *
 * @param   Mixed  instance  Response or Request instance
 * @return  Mixed
 */
function clone(instance) {
	let p1, p2;
	let body = instance.body;

	// don't allow cloning a used body
	if (instance.bodyUsed) {
		throw new Error('cannot clone body after it is used');
	}

	// check that body is a stream and not form-data object
	// note: we can't clone the form-data object without having it as a dependency
	if (body instanceof stream__WEBPACK_IMPORTED_MODULE_0__ && typeof body.getBoundary !== 'function') {
		// tee instance body
		p1 = new PassThrough();
		p2 = new PassThrough();
		body.pipe(p1);
		body.pipe(p2);
		// set instance body to teed body and return the other teed body
		instance[INTERNALS].body = p1;
		body = p2;
	}

	return body;
}

/**
 * Performs the operation "extract a `Content-Type` value from |object|" as
 * specified in the specification:
 * https://fetch.spec.whatwg.org/#concept-bodyinit-extract
 *
 * This function assumes that instance.body is present.
 *
 * @param   Mixed  instance  Any options.body input
 */
function extractContentType(body) {
	if (body === null) {
		// body is null
		return null;
	} else if (typeof body === 'string') {
		// body is string
		return 'text/plain;charset=UTF-8';
	} else if (isURLSearchParams(body)) {
		// body is a URLSearchParams
		return 'application/x-www-form-urlencoded;charset=UTF-8';
	} else if (isBlob(body)) {
		// body is blob
		return body.type || null;
	} else if (Buffer.isBuffer(body)) {
		// body is buffer
		return null;
	} else if (Object.prototype.toString.call(body) === '[object ArrayBuffer]') {
		// body is ArrayBuffer
		return null;
	} else if (ArrayBuffer.isView(body)) {
		// body is ArrayBufferView
		return null;
	} else if (typeof body.getBoundary === 'function') {
		// detect form data input from form-data module
		return `multipart/form-data;boundary=${body.getBoundary()}`;
	} else if (body instanceof stream__WEBPACK_IMPORTED_MODULE_0__) {
		// body is stream
		// can't really do much about this
		return null;
	} else {
		// Body constructor defaults other things to string
		return 'text/plain;charset=UTF-8';
	}
}

/**
 * The Fetch Standard treats this as if "total bytes" is a property on the body.
 * For us, we have to explicitly get it with a function.
 *
 * ref: https://fetch.spec.whatwg.org/#concept-body-total-bytes
 *
 * @param   Body    instance   Instance of Body
 * @return  Number?            Number of bytes, or null if not possible
 */
function getTotalBytes(instance) {
	const body = instance.body;


	if (body === null) {
		// body is null
		return 0;
	} else if (isBlob(body)) {
		return body.size;
	} else if (Buffer.isBuffer(body)) {
		// body is buffer
		return body.length;
	} else if (body && typeof body.getLengthSync === 'function') {
		// detect form data input from form-data module
		if (body._lengthRetrievers && body._lengthRetrievers.length == 0 || // 1.x
		body.hasKnownLength && body.hasKnownLength()) {
			// 2.x
			return body.getLengthSync();
		}
		return null;
	} else {
		// body is stream
		return null;
	}
}

/**
 * Write a Body to a Node.js WritableStream (e.g. http.Request) object.
 *
 * @param   Body    instance   Instance of Body
 * @return  Void
 */
function writeToStream(dest, instance) {
	const body = instance.body;


	if (body === null) {
		// body is null
		dest.end();
	} else if (isBlob(body)) {
		body.stream().pipe(dest);
	} else if (Buffer.isBuffer(body)) {
		// body is buffer
		dest.write(body);
		dest.end();
	} else {
		// body is stream
		body.pipe(dest);
	}
}

// expose Promise
Body.Promise = global.Promise;

/**
 * headers.js
 *
 * Headers class offers convenient helpers
 */

const invalidTokenRegex = /[^\^_`a-zA-Z\-0-9!#$%&'*+.|~]/;
const invalidHeaderCharRegex = /[^\t\x20-\x7e\x80-\xff]/;

function validateName(name) {
	name = `${name}`;
	if (invalidTokenRegex.test(name) || name === '') {
		throw new TypeError(`${name} is not a legal HTTP header name`);
	}
}

function validateValue(value) {
	value = `${value}`;
	if (invalidHeaderCharRegex.test(value)) {
		throw new TypeError(`${value} is not a legal HTTP header value`);
	}
}

/**
 * Find the key in the map object given a header name.
 *
 * Returns undefined if not found.
 *
 * @param   String  name  Header name
 * @return  String|Undefined
 */
function find(map, name) {
	name = name.toLowerCase();
	for (const key in map) {
		if (key.toLowerCase() === name) {
			return key;
		}
	}
	return undefined;
}

const MAP = Symbol('map');
class Headers {
	/**
  * Headers class
  *
  * @param   Object  headers  Response headers
  * @return  Void
  */
	constructor() {
		let init = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : undefined;

		this[MAP] = Object.create(null);

		if (init instanceof Headers) {
			const rawHeaders = init.raw();
			const headerNames = Object.keys(rawHeaders);

			for (const headerName of headerNames) {
				for (const value of rawHeaders[headerName]) {
					this.append(headerName, value);
				}
			}

			return;
		}

		// We don't worry about converting prop to ByteString here as append()
		// will handle it.
		if (init == null) ; else if (typeof init === 'object') {
			const method = init[Symbol.iterator];
			if (method != null) {
				if (typeof method !== 'function') {
					throw new TypeError('Header pairs must be iterable');
				}

				// sequence<sequence<ByteString>>
				// Note: per spec we have to first exhaust the lists then process them
				const pairs = [];
				for (const pair of init) {
					if (typeof pair !== 'object' || typeof pair[Symbol.iterator] !== 'function') {
						throw new TypeError('Each header pair must be iterable');
					}
					pairs.push(Array.from(pair));
				}

				for (const pair of pairs) {
					if (pair.length !== 2) {
						throw new TypeError('Each header pair must be a name/value tuple');
					}
					this.append(pair[0], pair[1]);
				}
			} else {
				// record<ByteString, ByteString>
				for (const key of Object.keys(init)) {
					const value = init[key];
					this.append(key, value);
				}
			}
		} else {
			throw new TypeError('Provided initializer must be an object');
		}
	}

	/**
  * Return combined header value given name
  *
  * @param   String  name  Header name
  * @return  Mixed
  */
	get(name) {
		name = `${name}`;
		validateName(name);
		const key = find(this[MAP], name);
		if (key === undefined) {
			return null;
		}

		return this[MAP][key].join(', ');
	}

	/**
  * Iterate over all headers
  *
  * @param   Function  callback  Executed for each item with parameters (value, name, thisArg)
  * @param   Boolean   thisArg   `this` context for callback function
  * @return  Void
  */
	forEach(callback) {
		let thisArg = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : undefined;

		let pairs = getHeaders(this);
		let i = 0;
		while (i < pairs.length) {
			var _pairs$i = pairs[i];
			const name = _pairs$i[0],
			      value = _pairs$i[1];

			callback.call(thisArg, value, name, this);
			pairs = getHeaders(this);
			i++;
		}
	}

	/**
  * Overwrite header values given name
  *
  * @param   String  name   Header name
  * @param   String  value  Header value
  * @return  Void
  */
	set(name, value) {
		name = `${name}`;
		value = `${value}`;
		validateName(name);
		validateValue(value);
		const key = find(this[MAP], name);
		this[MAP][key !== undefined ? key : name] = [value];
	}

	/**
  * Append a value onto existing header
  *
  * @param   String  name   Header name
  * @param   String  value  Header value
  * @return  Void
  */
	append(name, value) {
		name = `${name}`;
		value = `${value}`;
		validateName(name);
		validateValue(value);
		const key = find(this[MAP], name);
		if (key !== undefined) {
			this[MAP][key].push(value);
		} else {
			this[MAP][name] = [value];
		}
	}

	/**
  * Check for header name existence
  *
  * @param   String   name  Header name
  * @return  Boolean
  */
	has(name) {
		name = `${name}`;
		validateName(name);
		return find(this[MAP], name) !== undefined;
	}

	/**
  * Delete all header values given name
  *
  * @param   String  name  Header name
  * @return  Void
  */
	delete(name) {
		name = `${name}`;
		validateName(name);
		const key = find(this[MAP], name);
		if (key !== undefined) {
			delete this[MAP][key];
		}
	}

	/**
  * Return raw headers (non-spec api)
  *
  * @return  Object
  */
	raw() {
		return this[MAP];
	}

	/**
  * Get an iterator on keys.
  *
  * @return  Iterator
  */
	keys() {
		return createHeadersIterator(this, 'key');
	}

	/**
  * Get an iterator on values.
  *
  * @return  Iterator
  */
	values() {
		return createHeadersIterator(this, 'value');
	}

	/**
  * Get an iterator on entries.
  *
  * This is the default iterator of the Headers object.
  *
  * @return  Iterator
  */
	[Symbol.iterator]() {
		return createHeadersIterator(this, 'key+value');
	}
}
Headers.prototype.entries = Headers.prototype[Symbol.iterator];

Object.defineProperty(Headers.prototype, Symbol.toStringTag, {
	value: 'Headers',
	writable: false,
	enumerable: false,
	configurable: true
});

Object.defineProperties(Headers.prototype, {
	get: { enumerable: true },
	forEach: { enumerable: true },
	set: { enumerable: true },
	append: { enumerable: true },
	has: { enumerable: true },
	delete: { enumerable: true },
	keys: { enumerable: true },
	values: { enumerable: true },
	entries: { enumerable: true }
});

function getHeaders(headers) {
	let kind = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 'key+value';

	const keys = Object.keys(headers[MAP]).sort();
	return keys.map(kind === 'key' ? function (k) {
		return k.toLowerCase();
	} : kind === 'value' ? function (k) {
		return headers[MAP][k].join(', ');
	} : function (k) {
		return [k.toLowerCase(), headers[MAP][k].join(', ')];
	});
}

const INTERNAL = Symbol('internal');

function createHeadersIterator(target, kind) {
	const iterator = Object.create(HeadersIteratorPrototype);
	iterator[INTERNAL] = {
		target,
		kind,
		index: 0
	};
	return iterator;
}

const HeadersIteratorPrototype = Object.setPrototypeOf({
	next() {
		// istanbul ignore if
		if (!this || Object.getPrototypeOf(this) !== HeadersIteratorPrototype) {
			throw new TypeError('Value of `this` is not a HeadersIterator');
		}

		var _INTERNAL = this[INTERNAL];
		const target = _INTERNAL.target,
		      kind = _INTERNAL.kind,
		      index = _INTERNAL.index;

		const values = getHeaders(target, kind);
		const len = values.length;
		if (index >= len) {
			return {
				value: undefined,
				done: true
			};
		}

		this[INTERNAL].index = index + 1;

		return {
			value: values[index],
			done: false
		};
	}
}, Object.getPrototypeOf(Object.getPrototypeOf([][Symbol.iterator]())));

Object.defineProperty(HeadersIteratorPrototype, Symbol.toStringTag, {
	value: 'HeadersIterator',
	writable: false,
	enumerable: false,
	configurable: true
});

/**
 * Export the Headers object in a form that Node.js can consume.
 *
 * @param   Headers  headers
 * @return  Object
 */
function exportNodeCompatibleHeaders(headers) {
	const obj = Object.assign({ __proto__: null }, headers[MAP]);

	// http.request() only supports string as Host header. This hack makes
	// specifying custom Host header possible.
	const hostHeaderKey = find(headers[MAP], 'Host');
	if (hostHeaderKey !== undefined) {
		obj[hostHeaderKey] = obj[hostHeaderKey][0];
	}

	return obj;
}

/**
 * Create a Headers object from an object of headers, ignoring those that do
 * not conform to HTTP grammar productions.
 *
 * @param   Object  obj  Object of headers
 * @return  Headers
 */
function createHeadersLenient(obj) {
	const headers = new Headers();
	for (const name of Object.keys(obj)) {
		if (invalidTokenRegex.test(name)) {
			continue;
		}
		if (Array.isArray(obj[name])) {
			for (const val of obj[name]) {
				if (invalidHeaderCharRegex.test(val)) {
					continue;
				}
				if (headers[MAP][name] === undefined) {
					headers[MAP][name] = [val];
				} else {
					headers[MAP][name].push(val);
				}
			}
		} else if (!invalidHeaderCharRegex.test(obj[name])) {
			headers[MAP][name] = [obj[name]];
		}
	}
	return headers;
}

const INTERNALS$1 = Symbol('Response internals');

// fix an issue where "STATUS_CODES" aren't a named export for node <10
const STATUS_CODES = http__WEBPACK_IMPORTED_MODULE_1__.STATUS_CODES;

/**
 * Response class
 *
 * @param   Stream  body  Readable stream
 * @param   Object  opts  Response options
 * @return  Void
 */
class Response {
	constructor() {
		let body = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : null;
		let opts = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};

		Body.call(this, body, opts);

		const status = opts.status || 200;
		const headers = new Headers(opts.headers);

		if (body != null && !headers.has('Content-Type')) {
			const contentType = extractContentType(body);
			if (contentType) {
				headers.append('Content-Type', contentType);
			}
		}

		this[INTERNALS$1] = {
			url: opts.url,
			status,
			statusText: opts.statusText || STATUS_CODES[status],
			headers,
			counter: opts.counter
		};
	}

	get url() {
		return this[INTERNALS$1].url || '';
	}

	get status() {
		return this[INTERNALS$1].status;
	}

	/**
  * Convenience property representing if the request ended normally
  */
	get ok() {
		return this[INTERNALS$1].status >= 200 && this[INTERNALS$1].status < 300;
	}

	get redirected() {
		return this[INTERNALS$1].counter > 0;
	}

	get statusText() {
		return this[INTERNALS$1].statusText;
	}

	get headers() {
		return this[INTERNALS$1].headers;
	}

	/**
  * Clone this response
  *
  * @return  Response
  */
	clone() {
		return new Response(clone(this), {
			url: this.url,
			status: this.status,
			statusText: this.statusText,
			headers: this.headers,
			ok: this.ok,
			redirected: this.redirected
		});
	}
}

Body.mixIn(Response.prototype);

Object.defineProperties(Response.prototype, {
	url: { enumerable: true },
	status: { enumerable: true },
	ok: { enumerable: true },
	redirected: { enumerable: true },
	statusText: { enumerable: true },
	headers: { enumerable: true },
	clone: { enumerable: true }
});

Object.defineProperty(Response.prototype, Symbol.toStringTag, {
	value: 'Response',
	writable: false,
	enumerable: false,
	configurable: true
});

const INTERNALS$2 = Symbol('Request internals');

// fix an issue where "format", "parse" aren't a named export for node <10
const parse_url = url__WEBPACK_IMPORTED_MODULE_2__.parse;
const format_url = url__WEBPACK_IMPORTED_MODULE_2__.format;

const streamDestructionSupported = 'destroy' in stream__WEBPACK_IMPORTED_MODULE_0__.Readable.prototype;

/**
 * Check if a value is an instance of Request.
 *
 * @param   Mixed   input
 * @return  Boolean
 */
function isRequest(input) {
	return typeof input === 'object' && typeof input[INTERNALS$2] === 'object';
}

function isAbortSignal(signal) {
	const proto = signal && typeof signal === 'object' && Object.getPrototypeOf(signal);
	return !!(proto && proto.constructor.name === 'AbortSignal');
}

/**
 * Request class
 *
 * @param   Mixed   input  Url or Request instance
 * @param   Object  init   Custom options
 * @return  Void
 */
class Request {
	constructor(input) {
		let init = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};

		let parsedURL;

		// normalize input
		if (!isRequest(input)) {
			if (input && input.href) {
				// in order to support Node.js' Url objects; though WHATWG's URL objects
				// will fall into this branch also (since their `toString()` will return
				// `href` property anyway)
				parsedURL = parse_url(input.href);
			} else {
				// coerce input to a string before attempting to parse
				parsedURL = parse_url(`${input}`);
			}
			input = {};
		} else {
			parsedURL = parse_url(input.url);
		}

		let method = init.method || input.method || 'GET';
		method = method.toUpperCase();

		if ((init.body != null || isRequest(input) && input.body !== null) && (method === 'GET' || method === 'HEAD')) {
			throw new TypeError('Request with GET/HEAD method cannot have body');
		}

		let inputBody = init.body != null ? init.body : isRequest(input) && input.body !== null ? clone(input) : null;

		Body.call(this, inputBody, {
			timeout: init.timeout || input.timeout || 0,
			size: init.size || input.size || 0
		});

		const headers = new Headers(init.headers || input.headers || {});

		if (inputBody != null && !headers.has('Content-Type')) {
			const contentType = extractContentType(inputBody);
			if (contentType) {
				headers.append('Content-Type', contentType);
			}
		}

		let signal = isRequest(input) ? input.signal : null;
		if ('signal' in init) signal = init.signal;

		if (signal != null && !isAbortSignal(signal)) {
			throw new TypeError('Expected signal to be an instanceof AbortSignal');
		}

		this[INTERNALS$2] = {
			method,
			redirect: init.redirect || input.redirect || 'follow',
			headers,
			parsedURL,
			signal
		};

		// node-fetch-only options
		this.follow = init.follow !== undefined ? init.follow : input.follow !== undefined ? input.follow : 20;
		this.compress = init.compress !== undefined ? init.compress : input.compress !== undefined ? input.compress : true;
		this.counter = init.counter || input.counter || 0;
		this.agent = init.agent || input.agent;
	}

	get method() {
		return this[INTERNALS$2].method;
	}

	get url() {
		return format_url(this[INTERNALS$2].parsedURL);
	}

	get headers() {
		return this[INTERNALS$2].headers;
	}

	get redirect() {
		return this[INTERNALS$2].redirect;
	}

	get signal() {
		return this[INTERNALS$2].signal;
	}

	/**
  * Clone this request
  *
  * @return  Request
  */
	clone() {
		return new Request(this);
	}
}

Body.mixIn(Request.prototype);

Object.defineProperty(Request.prototype, Symbol.toStringTag, {
	value: 'Request',
	writable: false,
	enumerable: false,
	configurable: true
});

Object.defineProperties(Request.prototype, {
	method: { enumerable: true },
	url: { enumerable: true },
	headers: { enumerable: true },
	redirect: { enumerable: true },
	clone: { enumerable: true },
	signal: { enumerable: true }
});

/**
 * Convert a Request to Node.js http request options.
 *
 * @param   Request  A Request instance
 * @return  Object   The options object to be passed to http.request
 */
function getNodeRequestOptions(request) {
	const parsedURL = request[INTERNALS$2].parsedURL;
	const headers = new Headers(request[INTERNALS$2].headers);

	// fetch step 1.3
	if (!headers.has('Accept')) {
		headers.set('Accept', '*/*');
	}

	// Basic fetch
	if (!parsedURL.protocol || !parsedURL.hostname) {
		throw new TypeError('Only absolute URLs are supported');
	}

	if (!/^https?:$/.test(parsedURL.protocol)) {
		throw new TypeError('Only HTTP(S) protocols are supported');
	}

	if (request.signal && request.body instanceof stream__WEBPACK_IMPORTED_MODULE_0__.Readable && !streamDestructionSupported) {
		throw new Error('Cancellation of streamed requests with AbortSignal is not supported in node < 8');
	}

	// HTTP-network-or-cache fetch steps 2.4-2.7
	let contentLengthValue = null;
	if (request.body == null && /^(POST|PUT)$/i.test(request.method)) {
		contentLengthValue = '0';
	}
	if (request.body != null) {
		const totalBytes = getTotalBytes(request);
		if (typeof totalBytes === 'number') {
			contentLengthValue = String(totalBytes);
		}
	}
	if (contentLengthValue) {
		headers.set('Content-Length', contentLengthValue);
	}

	// HTTP-network-or-cache fetch step 2.11
	if (!headers.has('User-Agent')) {
		headers.set('User-Agent', 'node-fetch/1.0 (+https://github.com/bitinn/node-fetch)');
	}

	// HTTP-network-or-cache fetch step 2.15
	if (request.compress && !headers.has('Accept-Encoding')) {
		headers.set('Accept-Encoding', 'gzip,deflate');
	}

	let agent = request.agent;
	if (typeof agent === 'function') {
		agent = agent(parsedURL);
	}

	if (!headers.has('Connection') && !agent) {
		headers.set('Connection', 'close');
	}

	// HTTP-network fetch step 4.2
	// chunked encoding is handled by Node.js

	return Object.assign({}, parsedURL, {
		method: request.method,
		headers: exportNodeCompatibleHeaders(headers),
		agent
	});
}

/**
 * abort-error.js
 *
 * AbortError interface for cancelled requests
 */

/**
 * Create AbortError instance
 *
 * @param   String      message      Error message for human
 * @return  AbortError
 */
function AbortError(message) {
  Error.call(this, message);

  this.type = 'aborted';
  this.message = message;

  // hide custom error implementation details from end-users
  Error.captureStackTrace(this, this.constructor);
}

AbortError.prototype = Object.create(Error.prototype);
AbortError.prototype.constructor = AbortError;
AbortError.prototype.name = 'AbortError';

// fix an issue where "PassThrough", "resolve" aren't a named export for node <10
const PassThrough$1 = stream__WEBPACK_IMPORTED_MODULE_0__.PassThrough;
const resolve_url = url__WEBPACK_IMPORTED_MODULE_2__.resolve;

/**
 * Fetch function
 *
 * @param   Mixed    url   Absolute url or Request instance
 * @param   Object   opts  Fetch options
 * @return  Promise
 */
function fetch(url, opts) {

	// allow custom promise
	if (!fetch.Promise) {
		throw new Error('native promise missing, set fetch.Promise to your favorite alternative');
	}

	Body.Promise = fetch.Promise;

	// wrap http.request into fetch
	return new fetch.Promise(function (resolve, reject) {
		// build request object
		const request = new Request(url, opts);
		const options = getNodeRequestOptions(request);

		const send = (options.protocol === 'https:' ? https__WEBPACK_IMPORTED_MODULE_3__ : http__WEBPACK_IMPORTED_MODULE_1__).request;
		const signal = request.signal;

		let response = null;

		const abort = function abort() {
			let error = new AbortError('The user aborted a request.');
			reject(error);
			if (request.body && request.body instanceof stream__WEBPACK_IMPORTED_MODULE_0__.Readable) {
				request.body.destroy(error);
			}
			if (!response || !response.body) return;
			response.body.emit('error', error);
		};

		if (signal && signal.aborted) {
			abort();
			return;
		}

		const abortAndFinalize = function abortAndFinalize() {
			abort();
			finalize();
		};

		// send request
		const req = send(options);
		let reqTimeout;

		if (signal) {
			signal.addEventListener('abort', abortAndFinalize);
		}

		function finalize() {
			req.abort();
			if (signal) signal.removeEventListener('abort', abortAndFinalize);
			clearTimeout(reqTimeout);
		}

		if (request.timeout) {
			req.once('socket', function (socket) {
				reqTimeout = setTimeout(function () {
					reject(new FetchError(`network timeout at: ${request.url}`, 'request-timeout'));
					finalize();
				}, request.timeout);
			});
		}

		req.on('error', function (err) {
			reject(new FetchError(`request to ${request.url} failed, reason: ${err.message}`, 'system', err));
			finalize();
		});

		req.on('response', function (res) {
			clearTimeout(reqTimeout);

			const headers = createHeadersLenient(res.headers);

			// HTTP fetch step 5
			if (fetch.isRedirect(res.statusCode)) {
				// HTTP fetch step 5.2
				const location = headers.get('Location');

				// HTTP fetch step 5.3
				const locationURL = location === null ? null : resolve_url(request.url, location);

				// HTTP fetch step 5.5
				switch (request.redirect) {
					case 'error':
						reject(new FetchError(`uri requested responds with a redirect, redirect mode is set to error: ${request.url}`, 'no-redirect'));
						finalize();
						return;
					case 'manual':
						// node-fetch-specific step: make manual redirect a bit easier to use by setting the Location header value to the resolved URL.
						if (locationURL !== null) {
							// handle corrupted header
							try {
								headers.set('Location', locationURL);
							} catch (err) {
								// istanbul ignore next: nodejs server prevent invalid response headers, we can't test this through normal request
								reject(err);
							}
						}
						break;
					case 'follow':
						// HTTP-redirect fetch step 2
						if (locationURL === null) {
							break;
						}

						// HTTP-redirect fetch step 5
						if (request.counter >= request.follow) {
							reject(new FetchError(`maximum redirect reached at: ${request.url}`, 'max-redirect'));
							finalize();
							return;
						}

						// HTTP-redirect fetch step 6 (counter increment)
						// Create a new Request object.
						const requestOpts = {
							headers: new Headers(request.headers),
							follow: request.follow,
							counter: request.counter + 1,
							agent: request.agent,
							compress: request.compress,
							method: request.method,
							body: request.body,
							signal: request.signal,
							timeout: request.timeout,
							size: request.size
						};

						// HTTP-redirect fetch step 9
						if (res.statusCode !== 303 && request.body && getTotalBytes(request) === null) {
							reject(new FetchError('Cannot follow redirect with body being a readable stream', 'unsupported-redirect'));
							finalize();
							return;
						}

						// HTTP-redirect fetch step 11
						if (res.statusCode === 303 || (res.statusCode === 301 || res.statusCode === 302) && request.method === 'POST') {
							requestOpts.method = 'GET';
							requestOpts.body = undefined;
							requestOpts.headers.delete('content-length');
						}

						// HTTP-redirect fetch step 15
						resolve(fetch(new Request(locationURL, requestOpts)));
						finalize();
						return;
				}
			}

			// prepare response
			res.once('end', function () {
				if (signal) signal.removeEventListener('abort', abortAndFinalize);
			});
			let body = res.pipe(new PassThrough$1());

			const response_options = {
				url: request.url,
				status: res.statusCode,
				statusText: res.statusMessage,
				headers: headers,
				size: request.size,
				timeout: request.timeout,
				counter: request.counter
			};

			// HTTP-network fetch step 12.1.1.3
			const codings = headers.get('Content-Encoding');

			// HTTP-network fetch step 12.1.1.4: handle content codings

			// in following scenarios we ignore compression support
			// 1. compression support is disabled
			// 2. HEAD request
			// 3. no Content-Encoding header
			// 4. no content response (204)
			// 5. content not modified response (304)
			if (!request.compress || request.method === 'HEAD' || codings === null || res.statusCode === 204 || res.statusCode === 304) {
				response = new Response(body, response_options);
				resolve(response);
				return;
			}

			// For Node v6+
			// Be less strict when decoding compressed responses, since sometimes
			// servers send slightly invalid responses that are still accepted
			// by common browsers.
			// Always using Z_SYNC_FLUSH is what cURL does.
			const zlibOptions = {
				flush: zlib__WEBPACK_IMPORTED_MODULE_4__.Z_SYNC_FLUSH,
				finishFlush: zlib__WEBPACK_IMPORTED_MODULE_4__.Z_SYNC_FLUSH
			};

			// for gzip
			if (codings == 'gzip' || codings == 'x-gzip') {
				body = body.pipe(zlib__WEBPACK_IMPORTED_MODULE_4__.createGunzip(zlibOptions));
				response = new Response(body, response_options);
				resolve(response);
				return;
			}

			// for deflate
			if (codings == 'deflate' || codings == 'x-deflate') {
				// handle the infamous raw deflate response from old servers
				// a hack for old IIS and Apache servers
				const raw = res.pipe(new PassThrough$1());
				raw.once('data', function (chunk) {
					// see http://stackoverflow.com/questions/37519828
					if ((chunk[0] & 0x0F) === 0x08) {
						body = body.pipe(zlib__WEBPACK_IMPORTED_MODULE_4__.createInflate());
					} else {
						body = body.pipe(zlib__WEBPACK_IMPORTED_MODULE_4__.createInflateRaw());
					}
					response = new Response(body, response_options);
					resolve(response);
				});
				return;
			}

			// for br
			if (codings == 'br' && typeof zlib__WEBPACK_IMPORTED_MODULE_4__.createBrotliDecompress === 'function') {
				body = body.pipe(zlib__WEBPACK_IMPORTED_MODULE_4__.createBrotliDecompress());
				response = new Response(body, response_options);
				resolve(response);
				return;
			}

			// otherwise, use response as-is
			response = new Response(body, response_options);
			resolve(response);
		});

		writeToStream(req, request);
	});
}
/**
 * Redirect code matching
 *
 * @param   Number   code  Status code
 * @return  Boolean
 */
fetch.isRedirect = function (code) {
	return code === 301 || code === 302 || code === 303 || code === 307 || code === 308;
};

// expose Promise
fetch.Promise = global.Promise;

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (fetch);



/***/ }),
/* 55 */
/***/ ((module) => {

"use strict";
module.exports = require("stream");;

/***/ }),
/* 56 */
/***/ ((module) => {

"use strict";
module.exports = require("http");;

/***/ }),
/* 57 */
/***/ ((module) => {

"use strict";
module.exports = require("url");;

/***/ }),
/* 58 */
/***/ ((module) => {

"use strict";
module.exports = require("https");;

/***/ }),
/* 59 */
/***/ ((module) => {

"use strict";
module.exports = require("zlib");;

/***/ }),
/* 60 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.MsaAuthProvider = void 0;
const MSAService_1 = __webpack_require__(61);
class MsaAuthProvider {
    constructor(context) {
        this.onDidChangeSessions = MSAService_1.onDidChangeSessions.event;
        this.msaService = new MSAService_1.MSAService(context);
    }
    initialize() {
        return this.msaService.initialize();
    }
    getSessions(scopes) {
        return this.msaService.getSessions(scopes === null || scopes === void 0 ? void 0 : scopes.sort());
    }
    createSession(scopes) {
        return __awaiter(this, void 0, void 0, function* () {
            const session = yield this.msaService.createSession(scopes.sort());
            MSAService_1.onDidChangeSessions.fire({ added: [session], removed: [], changed: [] });
            return session;
        });
    }
    removeSession(sessionId) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const session = yield this.msaService.removeSession(sessionId);
                if (session) {
                    MSAService_1.onDidChangeSessions.fire({ added: [], removed: [session], changed: [] });
                }
            }
            catch (e) {
                console.error(e);
            }
        });
    }
}
exports.MsaAuthProvider = MsaAuthProvider;
MsaAuthProvider.id = 'msa';


/***/ }),
/* 61 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.MSAService = exports.REFRESH_NETWORK_FAILURE = exports.onDidChangeSessions = void 0;
const randomBytes = __webpack_require__(62);
const querystring = __webpack_require__(64);
const buffer_1 = __webpack_require__(65);
const vscode = __webpack_require__(1);
const uuid_1 = __webpack_require__(66);
const cross_fetch_1 = __webpack_require__(53);
const keychain_1 = __webpack_require__(81);
const authServer_1 = __webpack_require__(82);
const redirectUrl = 'https://extension-auth-redirect.azurewebsites.net/';
const loginEndpointUrl = 'https://login.microsoftonline.com/';
const clientId = 'a4fd7674-4ebd-4dbc-831c-338314dd459e';
const tenant = 'common';
function toBase64UrlEncoding(base64string) {
    return base64string.replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_'); // Need to use base64url encoding
}
function parseQuery(uri) {
    return uri.query.split('&').reduce((prev, current) => {
        const queryString = current.split('=');
        prev[queryString[0]] = queryString[1];
        return prev;
    }, {});
}
function sha256(s) {
    return __awaiter(this, void 0, void 0, function* () {
        return __webpack_require__(63).createHash('sha256').update(s).digest('base64');
    });
}
exports.onDidChangeSessions = new vscode.EventEmitter();
exports.REFRESH_NETWORK_FAILURE = 'Network failure';
class UriEventHandler extends vscode.EventEmitter {
    handleUri(uri) {
        this.fire(uri);
    }
}
class MSAService {
    constructor(_context) {
        this._context = _context;
        this._tokens = [];
        this._refreshTimeouts = new Map();
        this._disposables = [];
        // Used to keep track of current requests when not using the local server approach.
        this._pendingStates = new Map();
        this._codeExchangePromises = new Map();
        this._codeVerfifiers = new Map();
        this._keychain = new keychain_1.Keychain(_context);
        this._uriHandler = new UriEventHandler();
        this._disposables.push(vscode.window.registerUriHandler(this._uriHandler));
    }
    initialize() {
        return __awaiter(this, void 0, void 0, function* () {
            const storedData = yield this._keychain.getToken();
            if (storedData) {
                try {
                    const sessions = this.parseStoredData(storedData);
                    const refreshes = sessions.map((session) => __awaiter(this, void 0, void 0, function* () {
                        var _a;
                        if (!session.refreshToken) {
                            return Promise.resolve();
                        }
                        try {
                            yield this.refreshToken(session.refreshToken, session.scope, session.id);
                        }
                        catch (e) {
                            if (e.message === exports.REFRESH_NETWORK_FAILURE) {
                                const didSucceedOnRetry = yield this.handleRefreshNetworkError(session.id, session.refreshToken, session.scope);
                                if (!didSucceedOnRetry) {
                                    this._tokens.push({
                                        accessToken: undefined,
                                        refreshToken: session.refreshToken,
                                        account: {
                                            label: (_a = session.account.label) !== null && _a !== void 0 ? _a : session.account.displayName,
                                            id: session.account.id
                                        },
                                        scope: session.scope,
                                        sessionId: session.id
                                    });
                                    this.pollForReconnect(session.id, session.refreshToken, session.scope);
                                }
                            }
                            else {
                                yield this.removeSession(session.id);
                            }
                        }
                    }));
                    yield Promise.all(refreshes);
                }
                catch (e) {
                    console.info('Failed to initialize stored data');
                    yield this.clearSessions();
                }
            }
            this._disposables.push(this._context.secrets.onDidChange(() => this.checkForUpdates));
        });
    }
    parseStoredData(data) {
        return JSON.parse(data);
    }
    storeTokenData() {
        return __awaiter(this, void 0, void 0, function* () {
            const serializedData = this._tokens.map(token => {
                return {
                    id: token.sessionId,
                    refreshToken: token.refreshToken,
                    scope: token.scope,
                    account: token.account
                };
            });
            yield this._keychain.setToken(JSON.stringify(serializedData));
        });
    }
    checkForUpdates() {
        return __awaiter(this, void 0, void 0, function* () {
            const added = [];
            let removed = [];
            const storedData = yield this._keychain.getToken();
            if (storedData) {
                try {
                    const sessions = this.parseStoredData(storedData);
                    let promises = sessions.map((session) => __awaiter(this, void 0, void 0, function* () {
                        const matchesExisting = this._tokens.some(token => token.scope === session.scope && token.sessionId === session.id);
                        if (!matchesExisting && session.refreshToken) {
                            try {
                                const token = yield this.refreshToken(session.refreshToken, session.scope, session.id);
                                added.push(this.convertToSessionSync(token));
                            }
                            catch (e) {
                                if (e.message === exports.REFRESH_NETWORK_FAILURE) {
                                    // Ignore, will automatically retry on next poll.
                                }
                                else {
                                    yield this.removeSession(session.id);
                                }
                            }
                        }
                    }));
                    promises = promises.concat(this._tokens.map((token) => __awaiter(this, void 0, void 0, function* () {
                        const matchesExisting = sessions.some(session => token.scope === session.scope && token.sessionId === session.id);
                        if (!matchesExisting) {
                            yield this.removeSession(token.sessionId);
                            removed.push(this.convertToSessionSync(token));
                        }
                    })));
                    yield Promise.all(promises);
                }
                catch (e) {
                    console.error(e.message);
                    // if data is improperly formatted, remove all of it and send change event
                    removed = this._tokens.map(this.convertToSessionSync);
                    this.clearSessions();
                }
            }
            else {
                if (this._tokens.length) {
                    // Log out all, remove all local data
                    removed = this._tokens.map(this.convertToSessionSync);
                    console.info('No stored keychain data, clearing local data');
                    this._tokens = [];
                    this._refreshTimeouts.forEach(timeout => {
                        clearTimeout(timeout);
                    });
                    this._refreshTimeouts.clear();
                }
            }
            if (added.length || removed.length) {
                exports.onDidChangeSessions.fire({ added: added, removed: removed, changed: [] });
            }
        });
    }
    /**
     * Return a session object without checking for expiry and potentially refreshing.
     * @param token The token information.
     */
    convertToSessionSync(token) {
        return {
            id: token.sessionId,
            accessToken: token.accessToken,
            idToken: token.idToken,
            account: token.account,
            scopes: token.scope.split(' ')
        };
    }
    convertToSession(token) {
        return __awaiter(this, void 0, void 0, function* () {
            const resolvedTokens = yield this.resolveAccessAndIdTokens(token);
            return {
                id: token.sessionId,
                accessToken: resolvedTokens.accessToken,
                idToken: resolvedTokens.idToken,
                account: token.account,
                scopes: token.scope.split(' ')
            };
        });
    }
    resolveAccessAndIdTokens(token) {
        return __awaiter(this, void 0, void 0, function* () {
            if (token.accessToken && (!token.expiresAt || token.expiresAt > Date.now())) {
                token.expiresAt
                    ? console.info(`Token available from cache, expires in ${token.expiresAt - Date.now()} milliseconds`)
                    : console.info('Token available from cache');
                return Promise.resolve({
                    accessToken: token.accessToken,
                    idToken: token.idToken
                });
            }
            try {
                console.info('Token expired or unavailable, trying refresh');
                const refreshedToken = yield this.refreshToken(token.refreshToken, token.scope, token.sessionId);
                if (refreshedToken.accessToken) {
                    return {
                        accessToken: refreshedToken.accessToken,
                        idToken: refreshedToken.idToken
                    };
                }
                else {
                    throw new Error();
                }
            }
            catch (e) {
                throw new Error('Unavailable due to network problems');
            }
        });
    }
    getTokenClaims(jwt) {
        try {
            return JSON.parse(buffer_1.Buffer.from(jwt.split('.')[1], 'base64').toString());
        }
        catch (e) {
            throw new Error('Unable to read token claims');
        }
    }
    get sessions() {
        return Promise.all(this._tokens.map(token => this.convertToSession(token)));
    }
    getSessions(scopes) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!scopes) {
                return this.sessions;
            }
            const orderedScopes = scopes.sort().join(' ');
            const matchingTokens = this._tokens.filter(token => token.scope === orderedScopes);
            return Promise.all(matchingTokens.map(token => this.convertToSession(token)));
        });
    }
    createSession(scopes) {
        return __awaiter(this, void 0, void 0, function* () {
            console.info('Logging in...');
            const runsRemote = vscode.env.remoteName !== undefined;
            const runsServerless = vscode.env.remoteName === undefined && vscode.env.uiKind === vscode.UIKind.Web;
            if (runsRemote || runsServerless) {
                return yield this.loginWithoutLocalServer(scopes);
            }
            const scopeStr = scopes.sort().join(' ');
            const nonce = randomBytes(16).toString('base64');
            const { server, redirectPromise, codePromise } = authServer_1.createServer(nonce);
            let token;
            try {
                const port = yield authServer_1.startServer(server);
                vscode.env.openExternal(vscode.Uri.parse(`http://localhost:${port}/signin?nonce=${encodeURIComponent(nonce)}`));
                const redirectReq = yield redirectPromise;
                if ('err' in redirectReq) {
                    const { err, res } = redirectReq;
                    res.writeHead(302, { Location: `/?error=${encodeURIComponent(err && err.message || 'Unknown error')}` });
                    res.end();
                    throw err;
                }
                const host = redirectReq.req.headers.host || '';
                const updatedPortStr = (/^[^:]+:(\d+)$/.exec(Array.isArray(host) ? host[0] : host) || [])[1];
                const updatedPort = updatedPortStr ? parseInt(updatedPortStr, 10) : port;
                const state = `${updatedPort},${encodeURIComponent(nonce)}`;
                const codeVerifier = toBase64UrlEncoding(randomBytes(32).toString('base64'));
                const codeChallenge = toBase64UrlEncoding(yield sha256(codeVerifier));
                const loginUrl = `${loginEndpointUrl}${tenant}/oauth2/v2.0/authorize?response_type=code&response_mode=query&client_id=${encodeURIComponent(clientId)}&redirect_uri=${encodeURIComponent(redirectUrl)}&state=${state}&scope=${encodeURIComponent(scopeStr)}&prompt=select_account&code_challenge_method=S256&code_challenge=${codeChallenge}`;
                yield redirectReq.res.writeHead(302, { Location: loginUrl });
                redirectReq.res.end();
                const codeRes = yield codePromise;
                const res = codeRes.res;
                try {
                    if ('err' in codeRes) {
                        throw codeRes.err;
                    }
                    token = yield this.exchangeCodeForToken(codeRes.code, codeVerifier, scopeStr);
                    this.setToken(token, scopeStr);
                    console.log('Login successful');
                    const session = yield this.convertToSession(token);
                    res.writeHead(302, { Location: '/' });
                    return session;
                }
                catch (err) {
                    res.writeHead(302, { Location: `/?error=${encodeURIComponent(err && err.message || 'Unknown error')}` });
                    throw err;
                }
                finally {
                    res.end();
                }
            }
            catch (e) {
                console.error(e.message);
                // If the error was about starting the server, try directly hitting the login endpoint instead
                if (e.message === 'Error listening to server' || e.message === 'Closed' || e.message === 'Timeout waiting for port') {
                    return yield this.loginWithoutLocalServer(scopes);
                }
                throw e;
            }
            finally {
                setTimeout(() => {
                    server.close();
                }, 5000);
            }
        });
    }
    dispose() {
        this._disposables.forEach(disposable => disposable.dispose());
        this._disposables = [];
    }
    getCallbackEnvironment(callbackUri) {
        if (callbackUri.scheme !== 'https' && callbackUri.scheme !== 'http') {
            return callbackUri.scheme;
        }
        switch (callbackUri.authority) {
            case 'online.visualstudio.com':
                return 'vso';
            case 'online-ppe.core.vsengsaas.visualstudio.com':
                return 'vsoppe';
            case 'online.dev.core.vsengsaas.visualstudio.com':
                return 'vsodev';
            default:
                return callbackUri.authority;
        }
    }
    loginWithoutLocalServer(scopes) {
        return __awaiter(this, void 0, void 0, function* () {
            const callbackUri = yield vscode.env.asExternalUri(vscode.Uri.parse(`${vscode.env.uriScheme}://tylerleonhardt.msft-todo-unofficial`));
            const nonce = randomBytes(16).toString('base64');
            const port = (callbackUri.authority.match(/:([0-9]*)$/) || [])[1] || (callbackUri.scheme === 'https' ? 443 : 80);
            const callbackEnvironment = this.getCallbackEnvironment(callbackUri);
            const state = `${callbackEnvironment}${port},${encodeURIComponent(nonce)},${encodeURIComponent(callbackUri.query)}`;
            const signInUrl = `${loginEndpointUrl}${tenant}/oauth2/v2.0/authorize`;
            let uri = vscode.Uri.parse(signInUrl);
            const codeVerifier = toBase64UrlEncoding(randomBytes(32).toString('base64'));
            const codeChallenge = toBase64UrlEncoding(yield sha256(codeVerifier));
            const scopeStr = scopes.sort().join(' ');
            uri = uri.with({
                query: `response_type=code&client_id=${encodeURIComponent(clientId)}&response_mode=query&redirect_uri=${redirectUrl}&state=${state}&scope=${scopeStr}&prompt=select_account&code_challenge_method=S256&code_challenge=${codeChallenge}`
            });
            vscode.env.openExternal(uri);
            const timeoutPromise = new Promise((_, reject) => {
                const wait = setTimeout(() => {
                    clearTimeout(wait);
                    reject('Login timed out.');
                }, 1000 * 60 * 5);
            });
            const existingStates = this._pendingStates.get(scopeStr) || [];
            this._pendingStates.set(scopeStr, [...existingStates, state]);
            // Register a single listener for the URI callback, in case the user starts the login process multiple times
            // before completing it.
            let existingPromise = this._codeExchangePromises.get(scopeStr);
            if (!existingPromise) {
                existingPromise = this.handleCodeResponse(scopeStr);
                this._codeExchangePromises.set(scopeStr, existingPromise);
            }
            this._codeVerfifiers.set(state, codeVerifier);
            return Promise.race([existingPromise, timeoutPromise])
                .finally(() => {
                this._pendingStates.delete(scopeStr);
                this._codeExchangePromises.delete(scopeStr);
                this._codeVerfifiers.delete(state);
            });
        });
    }
    handleCodeResponse(scopeStr) {
        return __awaiter(this, void 0, void 0, function* () {
            let uriEventListener;
            return new Promise((resolve, reject) => {
                uriEventListener = this._uriHandler.event((uri) => __awaiter(this, void 0, void 0, function* () {
                    var _a;
                    try {
                        const query = parseQuery(uri);
                        const code = query.code;
                        const acceptedStates = this._pendingStates.get(scopeStr) || [];
                        // Workaround double encoding issues of state in web
                        if (!acceptedStates.includes(query.state) && !acceptedStates.includes(decodeURIComponent(query.state))) {
                            throw new Error('State does not match.');
                        }
                        const verifier = (_a = this._codeVerfifiers.get(query.state)) !== null && _a !== void 0 ? _a : this._codeVerfifiers.get(decodeURIComponent(query.state));
                        if (!verifier) {
                            throw new Error('No available code verifier');
                        }
                        const token = yield this.exchangeCodeForToken(code, verifier, scopeStr);
                        this.setToken(token, scopeStr);
                        const session = yield this.convertToSession(token);
                        resolve(session);
                    }
                    catch (err) {
                        reject(err);
                    }
                }));
            }).then(result => {
                uriEventListener.dispose();
                return result;
            }).catch(err => {
                uriEventListener.dispose();
                throw err;
            });
        });
    }
    setToken(token, scope) {
        return __awaiter(this, void 0, void 0, function* () {
            const existingTokenIndex = this._tokens.findIndex(t => t.sessionId === token.sessionId);
            if (existingTokenIndex > -1) {
                this._tokens.splice(existingTokenIndex, 1, token);
            }
            else {
                this._tokens.push(token);
            }
            this.clearSessionTimeout(token.sessionId);
            if (token.expiresIn) {
                this._refreshTimeouts.set(token.sessionId, setTimeout(() => __awaiter(this, void 0, void 0, function* () {
                    try {
                        const refreshedToken = yield this.refreshToken(token.refreshToken, scope, token.sessionId);
                        exports.onDidChangeSessions.fire({ added: [], removed: [], changed: [this.convertToSessionSync(refreshedToken)] });
                    }
                    catch (e) {
                        if (e.message === exports.REFRESH_NETWORK_FAILURE) {
                            const didSucceedOnRetry = yield this.handleRefreshNetworkError(token.sessionId, token.refreshToken, scope);
                            if (!didSucceedOnRetry) {
                                this.pollForReconnect(token.sessionId, token.refreshToken, token.scope);
                            }
                        }
                        else {
                            yield this.removeSession(token.sessionId);
                            exports.onDidChangeSessions.fire({ added: [], removed: [this.convertToSessionSync(token)], changed: [] });
                        }
                    }
                }), 1000 * (token.expiresIn - 30)));
            }
            this.storeTokenData();
        });
    }
    getTokenFromResponse(json, scope, existingId) {
        let claims = undefined;
        try {
            claims = this.getTokenClaims(json.access_token);
        }
        catch (e) {
            if (json.id_token) {
                console.log('Failed to fetch token claims from access_token. Attempting to parse id_token instead');
                claims = this.getTokenClaims(json.id_token);
            }
            else {
                throw e;
            }
        }
        return {
            expiresIn: json.expires_in,
            expiresAt: json.expires_in ? Date.now() + json.expires_in * 1000 : undefined,
            accessToken: json.access_token,
            idToken: json.id_token,
            refreshToken: json.refresh_token,
            scope,
            sessionId: uuid_1.v4(),
            account: {
                label: claims.email || claims.unique_name || claims.preferred_username || 'user@example.com',
                id: `${claims.tid}/${(claims.oid || (claims.altsecid || '' + claims.ipd || ''))}`
            }
        };
    }
    exchangeCodeForToken(code, codeVerifier, scopeStr) {
        return __awaiter(this, void 0, void 0, function* () {
            console.info('Exchanging login code for token');
            try {
                const postData = querystring.stringify({
                    grant_type: 'authorization_code',
                    code,
                    client_id: clientId,
                    scope: scopeStr,
                    code_verifier: codeVerifier,
                    redirect_uri: redirectUrl
                });
                const proxyEndpoints = yield vscode.commands.executeCommand('workbench.getCodeExchangeProxyEndpoints');
                const endpoint = proxyEndpoints && proxyEndpoints['microsoft'] || `${loginEndpointUrl}${tenant}/oauth2/v2.0/token`;
                const result = yield cross_fetch_1.default(endpoint, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'Content-Length': postData.length.toString()
                    },
                    body: postData
                });
                if (result.ok) {
                    console.info('Exchanging login code for token success');
                    const json = yield result.json();
                    return this.getTokenFromResponse(json, scopeStr);
                }
                else {
                    console.error('Exchanging login code for token failed');
                    throw new Error('Unable to login.');
                }
            }
            catch (e) {
                console.error(e.message);
                throw e;
            }
        });
    }
    refreshToken(refreshToken, scope, sessionId) {
        return __awaiter(this, void 0, void 0, function* () {
            console.info('Refreshing token...');
            const postData = querystring.stringify({
                refresh_token: refreshToken,
                client_id: clientId,
                grant_type: 'refresh_token',
                scope: scope
            });
            let result;
            try {
                result = yield cross_fetch_1.default(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'Content-Length': postData.length.toString()
                    },
                    body: postData
                });
            }
            catch (e) {
                console.error('Refreshing token failed');
                throw new Error(exports.REFRESH_NETWORK_FAILURE);
            }
            try {
                if (result.ok) {
                    const json = yield result.json();
                    const token = this.getTokenFromResponse(json, scope, sessionId);
                    this.setToken(token, scope);
                    console.info('Token refresh success');
                    return token;
                }
                else {
                    throw new Error('Bad request.');
                }
            }
            catch (e) {
                vscode.window.showErrorMessage("You have been signed out because reading stored authentication information failed.");
                console.error(`Refreshing token failed: ${result.statusText}`);
                throw new Error('Refreshing token failed');
            }
        });
    }
    clearSessionTimeout(sessionId) {
        const timeout = this._refreshTimeouts.get(sessionId);
        if (timeout) {
            clearTimeout(timeout);
            this._refreshTimeouts.delete(sessionId);
        }
    }
    removeInMemorySessionData(sessionId) {
        const tokenIndex = this._tokens.findIndex(token => token.sessionId === sessionId);
        let token;
        if (tokenIndex > -1) {
            token = this._tokens[tokenIndex];
            this._tokens.splice(tokenIndex, 1);
        }
        this.clearSessionTimeout(sessionId);
        return token;
    }
    pollForReconnect(sessionId, refreshToken, scope) {
        this.clearSessionTimeout(sessionId);
        this._refreshTimeouts.set(sessionId, setTimeout(() => __awaiter(this, void 0, void 0, function* () {
            try {
                const refreshedToken = yield this.refreshToken(refreshToken, scope, sessionId);
                exports.onDidChangeSessions.fire({ added: [], removed: [], changed: [this.convertToSessionSync(refreshedToken)] });
            }
            catch (e) {
                this.pollForReconnect(sessionId, refreshToken, scope);
            }
        }), 1000 * 60 * 30));
    }
    handleRefreshNetworkError(sessionId, refreshToken, scope, attempts = 1) {
        return new Promise((resolve, _) => {
            if (attempts === 3) {
                console.error('Token refresh failed after 3 attempts');
                return resolve(false);
            }
            const delayBeforeRetry = 5 * attempts * attempts;
            this.clearSessionTimeout(sessionId);
            this._refreshTimeouts.set(sessionId, setTimeout(() => __awaiter(this, void 0, void 0, function* () {
                try {
                    const refreshedToken = yield this.refreshToken(refreshToken, scope, sessionId);
                    exports.onDidChangeSessions.fire({ added: [], removed: [], changed: [this.convertToSessionSync(refreshedToken)] });
                    return resolve(true);
                }
                catch (e) {
                    return resolve(yield this.handleRefreshNetworkError(sessionId, refreshToken, scope, attempts + 1));
                }
            }), 1000 * delayBeforeRetry));
        });
    }
    removeSession(sessionId) {
        return __awaiter(this, void 0, void 0, function* () {
            console.info(`Logging out of session '${sessionId}'`);
            const token = this.removeInMemorySessionData(sessionId);
            let session;
            if (token) {
                session = this.convertToSessionSync(token);
            }
            if (this._tokens.length === 0) {
                yield this._keychain.deleteToken();
            }
            else {
                this.storeTokenData();
            }
            return session;
        });
    }
    clearSessions() {
        return __awaiter(this, void 0, void 0, function* () {
            console.info('Logging out of all sessions');
            this._tokens = [];
            yield this._keychain.deleteToken();
            this._refreshTimeouts.forEach(timeout => {
                clearTimeout(timeout);
            });
            this._refreshTimeouts.clear();
        });
    }
}
exports.MSAService = MSAService;


/***/ }),
/* 62 */
/***/ ((module, __unused_webpack_exports, __webpack_require__) => {

module.exports = __webpack_require__(63).randomBytes


/***/ }),
/* 63 */
/***/ ((module) => {

"use strict";
module.exports = require("crypto");;

/***/ }),
/* 64 */
/***/ ((module) => {

"use strict";
module.exports = require("querystring");;

/***/ }),
/* 65 */
/***/ ((module) => {

"use strict";
module.exports = require("buffer");;

/***/ }),
/* 66 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "v1": () => (/* reexport safe */ _v1_js__WEBPACK_IMPORTED_MODULE_0__.default),
/* harmony export */   "v3": () => (/* reexport safe */ _v3_js__WEBPACK_IMPORTED_MODULE_1__.default),
/* harmony export */   "v4": () => (/* reexport safe */ _v4_js__WEBPACK_IMPORTED_MODULE_2__.default),
/* harmony export */   "v5": () => (/* reexport safe */ _v5_js__WEBPACK_IMPORTED_MODULE_3__.default),
/* harmony export */   "NIL": () => (/* reexport safe */ _nil_js__WEBPACK_IMPORTED_MODULE_4__.default),
/* harmony export */   "version": () => (/* reexport safe */ _version_js__WEBPACK_IMPORTED_MODULE_5__.default),
/* harmony export */   "validate": () => (/* reexport safe */ _validate_js__WEBPACK_IMPORTED_MODULE_6__.default),
/* harmony export */   "stringify": () => (/* reexport safe */ _stringify_js__WEBPACK_IMPORTED_MODULE_7__.default),
/* harmony export */   "parse": () => (/* reexport safe */ _parse_js__WEBPACK_IMPORTED_MODULE_8__.default)
/* harmony export */ });
/* harmony import */ var _v1_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(67);
/* harmony import */ var _v3_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(72);
/* harmony import */ var _v4_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(76);
/* harmony import */ var _v5_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(77);
/* harmony import */ var _nil_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(79);
/* harmony import */ var _version_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(80);
/* harmony import */ var _validate_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(70);
/* harmony import */ var _stringify_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(69);
/* harmony import */ var _parse_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(74);










/***/ }),
/* 67 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _rng_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(68);
/* harmony import */ var _stringify_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(69);

 // **`v1()` - Generate time-based UUID**
//
// Inspired by https://github.com/LiosK/UUID.js
// and http://docs.python.org/library/uuid.html

let _nodeId;

let _clockseq; // Previous uuid creation time


let _lastMSecs = 0;
let _lastNSecs = 0; // See https://github.com/uuidjs/uuid for API details

function v1(options, buf, offset) {
  let i = buf && offset || 0;
  const b = buf || new Array(16);
  options = options || {};
  let node = options.node || _nodeId;
  let clockseq = options.clockseq !== undefined ? options.clockseq : _clockseq; // node and clockseq need to be initialized to random values if they're not
  // specified.  We do this lazily to minimize issues related to insufficient
  // system entropy.  See #189

  if (node == null || clockseq == null) {
    const seedBytes = options.random || (options.rng || _rng_js__WEBPACK_IMPORTED_MODULE_0__.default)();

    if (node == null) {
      // Per 4.5, create and 48-bit node id, (47 random bits + multicast bit = 1)
      node = _nodeId = [seedBytes[0] | 0x01, seedBytes[1], seedBytes[2], seedBytes[3], seedBytes[4], seedBytes[5]];
    }

    if (clockseq == null) {
      // Per 4.2.2, randomize (14 bit) clockseq
      clockseq = _clockseq = (seedBytes[6] << 8 | seedBytes[7]) & 0x3fff;
    }
  } // UUID timestamps are 100 nano-second units since the Gregorian epoch,
  // (1582-10-15 00:00).  JSNumbers aren't precise enough for this, so
  // time is handled internally as 'msecs' (integer milliseconds) and 'nsecs'
  // (100-nanoseconds offset from msecs) since unix epoch, 1970-01-01 00:00.


  let msecs = options.msecs !== undefined ? options.msecs : Date.now(); // Per 4.2.1.2, use count of uuid's generated during the current clock
  // cycle to simulate higher resolution clock

  let nsecs = options.nsecs !== undefined ? options.nsecs : _lastNSecs + 1; // Time since last uuid creation (in msecs)

  const dt = msecs - _lastMSecs + (nsecs - _lastNSecs) / 10000; // Per 4.2.1.2, Bump clockseq on clock regression

  if (dt < 0 && options.clockseq === undefined) {
    clockseq = clockseq + 1 & 0x3fff;
  } // Reset nsecs if clock regresses (new clockseq) or we've moved onto a new
  // time interval


  if ((dt < 0 || msecs > _lastMSecs) && options.nsecs === undefined) {
    nsecs = 0;
  } // Per 4.2.1.2 Throw error if too many uuids are requested


  if (nsecs >= 10000) {
    throw new Error("uuid.v1(): Can't create more than 10M uuids/sec");
  }

  _lastMSecs = msecs;
  _lastNSecs = nsecs;
  _clockseq = clockseq; // Per 4.1.4 - Convert from unix epoch to Gregorian epoch

  msecs += 12219292800000; // `time_low`

  const tl = ((msecs & 0xfffffff) * 10000 + nsecs) % 0x100000000;
  b[i++] = tl >>> 24 & 0xff;
  b[i++] = tl >>> 16 & 0xff;
  b[i++] = tl >>> 8 & 0xff;
  b[i++] = tl & 0xff; // `time_mid`

  const tmh = msecs / 0x100000000 * 10000 & 0xfffffff;
  b[i++] = tmh >>> 8 & 0xff;
  b[i++] = tmh & 0xff; // `time_high_and_version`

  b[i++] = tmh >>> 24 & 0xf | 0x10; // include version

  b[i++] = tmh >>> 16 & 0xff; // `clock_seq_hi_and_reserved` (Per 4.2.2 - include variant)

  b[i++] = clockseq >>> 8 | 0x80; // `clock_seq_low`

  b[i++] = clockseq & 0xff; // `node`

  for (let n = 0; n < 6; ++n) {
    b[i + n] = node[n];
  }

  return buf || (0,_stringify_js__WEBPACK_IMPORTED_MODULE_1__.default)(b);
}

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (v1);

/***/ }),
/* 68 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (/* binding */ rng)
/* harmony export */ });
/* harmony import */ var crypto__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(63);
/* harmony import */ var crypto__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(crypto__WEBPACK_IMPORTED_MODULE_0__);

const rnds8Pool = new Uint8Array(256); // # of random values to pre-allocate

let poolPtr = rnds8Pool.length;
function rng() {
  if (poolPtr > rnds8Pool.length - 16) {
    crypto__WEBPACK_IMPORTED_MODULE_0___default().randomFillSync(rnds8Pool);
    poolPtr = 0;
  }

  return rnds8Pool.slice(poolPtr, poolPtr += 16);
}

/***/ }),
/* 69 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _validate_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(70);

/**
 * Convert array of 16 byte values to UUID string format of the form:
 * XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
 */

const byteToHex = [];

for (let i = 0; i < 256; ++i) {
  byteToHex.push((i + 0x100).toString(16).substr(1));
}

function stringify(arr, offset = 0) {
  // Note: Be careful editing this code!  It's been tuned for performance
  // and works in ways you may not expect. See https://github.com/uuidjs/uuid/pull/434
  const uuid = (byteToHex[arr[offset + 0]] + byteToHex[arr[offset + 1]] + byteToHex[arr[offset + 2]] + byteToHex[arr[offset + 3]] + '-' + byteToHex[arr[offset + 4]] + byteToHex[arr[offset + 5]] + '-' + byteToHex[arr[offset + 6]] + byteToHex[arr[offset + 7]] + '-' + byteToHex[arr[offset + 8]] + byteToHex[arr[offset + 9]] + '-' + byteToHex[arr[offset + 10]] + byteToHex[arr[offset + 11]] + byteToHex[arr[offset + 12]] + byteToHex[arr[offset + 13]] + byteToHex[arr[offset + 14]] + byteToHex[arr[offset + 15]]).toLowerCase(); // Consistency check for valid UUID.  If this throws, it's likely due to one
  // of the following:
  // - One or more input array values don't map to a hex octet (leading to
  // "undefined" in the uuid)
  // - Invalid input values for the RFC `version` or `variant` fields

  if (!(0,_validate_js__WEBPACK_IMPORTED_MODULE_0__.default)(uuid)) {
    throw TypeError('Stringified UUID is invalid');
  }

  return uuid;
}

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (stringify);

/***/ }),
/* 70 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _regex_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(71);


function validate(uuid) {
  return typeof uuid === 'string' && _regex_js__WEBPACK_IMPORTED_MODULE_0__.default.test(uuid);
}

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (validate);

/***/ }),
/* 71 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (/^(?:[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}|00000000-0000-0000-0000-000000000000)$/i);

/***/ }),
/* 72 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _v35_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(73);
/* harmony import */ var _md5_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(75);


const v3 = (0,_v35_js__WEBPACK_IMPORTED_MODULE_0__.default)('v3', 0x30, _md5_js__WEBPACK_IMPORTED_MODULE_1__.default);
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (v3);

/***/ }),
/* 73 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "DNS": () => (/* binding */ DNS),
/* harmony export */   "URL": () => (/* binding */ URL),
/* harmony export */   "default": () => (/* export default binding */ __WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _stringify_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(69);
/* harmony import */ var _parse_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(74);



function stringToBytes(str) {
  str = unescape(encodeURIComponent(str)); // UTF8 escape

  const bytes = [];

  for (let i = 0; i < str.length; ++i) {
    bytes.push(str.charCodeAt(i));
  }

  return bytes;
}

const DNS = '6ba7b810-9dad-11d1-80b4-00c04fd430c8';
const URL = '6ba7b811-9dad-11d1-80b4-00c04fd430c8';
/* harmony default export */ function __WEBPACK_DEFAULT_EXPORT__(name, version, hashfunc) {
  function generateUUID(value, namespace, buf, offset) {
    if (typeof value === 'string') {
      value = stringToBytes(value);
    }

    if (typeof namespace === 'string') {
      namespace = (0,_parse_js__WEBPACK_IMPORTED_MODULE_0__.default)(namespace);
    }

    if (namespace.length !== 16) {
      throw TypeError('Namespace must be array-like (16 iterable integer values, 0-255)');
    } // Compute hash of namespace and value, Per 4.3
    // Future: Use spread syntax when supported on all platforms, e.g. `bytes =
    // hashfunc([...namespace, ... value])`


    let bytes = new Uint8Array(16 + value.length);
    bytes.set(namespace);
    bytes.set(value, namespace.length);
    bytes = hashfunc(bytes);
    bytes[6] = bytes[6] & 0x0f | version;
    bytes[8] = bytes[8] & 0x3f | 0x80;

    if (buf) {
      offset = offset || 0;

      for (let i = 0; i < 16; ++i) {
        buf[offset + i] = bytes[i];
      }

      return buf;
    }

    return (0,_stringify_js__WEBPACK_IMPORTED_MODULE_1__.default)(bytes);
  } // Function#name is not settable on some platforms (#270)


  try {
    generateUUID.name = name; // eslint-disable-next-line no-empty
  } catch (err) {} // For CommonJS default export support


  generateUUID.DNS = DNS;
  generateUUID.URL = URL;
  return generateUUID;
}

/***/ }),
/* 74 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _validate_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(70);


function parse(uuid) {
  if (!(0,_validate_js__WEBPACK_IMPORTED_MODULE_0__.default)(uuid)) {
    throw TypeError('Invalid UUID');
  }

  let v;
  const arr = new Uint8Array(16); // Parse ########-....-....-....-............

  arr[0] = (v = parseInt(uuid.slice(0, 8), 16)) >>> 24;
  arr[1] = v >>> 16 & 0xff;
  arr[2] = v >>> 8 & 0xff;
  arr[3] = v & 0xff; // Parse ........-####-....-....-............

  arr[4] = (v = parseInt(uuid.slice(9, 13), 16)) >>> 8;
  arr[5] = v & 0xff; // Parse ........-....-####-....-............

  arr[6] = (v = parseInt(uuid.slice(14, 18), 16)) >>> 8;
  arr[7] = v & 0xff; // Parse ........-....-....-####-............

  arr[8] = (v = parseInt(uuid.slice(19, 23), 16)) >>> 8;
  arr[9] = v & 0xff; // Parse ........-....-....-....-############
  // (Use "/" to avoid 32-bit truncation when bit-shifting high-order bytes)

  arr[10] = (v = parseInt(uuid.slice(24, 36), 16)) / 0x10000000000 & 0xff;
  arr[11] = v / 0x100000000 & 0xff;
  arr[12] = v >>> 24 & 0xff;
  arr[13] = v >>> 16 & 0xff;
  arr[14] = v >>> 8 & 0xff;
  arr[15] = v & 0xff;
  return arr;
}

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (parse);

/***/ }),
/* 75 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var crypto__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(63);
/* harmony import */ var crypto__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(crypto__WEBPACK_IMPORTED_MODULE_0__);


function md5(bytes) {
  if (Array.isArray(bytes)) {
    bytes = Buffer.from(bytes);
  } else if (typeof bytes === 'string') {
    bytes = Buffer.from(bytes, 'utf8');
  }

  return crypto__WEBPACK_IMPORTED_MODULE_0___default().createHash('md5').update(bytes).digest();
}

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (md5);

/***/ }),
/* 76 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _rng_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(68);
/* harmony import */ var _stringify_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(69);



function v4(options, buf, offset) {
  options = options || {};
  const rnds = options.random || (options.rng || _rng_js__WEBPACK_IMPORTED_MODULE_0__.default)(); // Per 4.4, set bits for version and `clock_seq_hi_and_reserved`

  rnds[6] = rnds[6] & 0x0f | 0x40;
  rnds[8] = rnds[8] & 0x3f | 0x80; // Copy bytes to buffer, if provided

  if (buf) {
    offset = offset || 0;

    for (let i = 0; i < 16; ++i) {
      buf[offset + i] = rnds[i];
    }

    return buf;
  }

  return (0,_stringify_js__WEBPACK_IMPORTED_MODULE_1__.default)(rnds);
}

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (v4);

/***/ }),
/* 77 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _v35_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(73);
/* harmony import */ var _sha1_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(78);


const v5 = (0,_v35_js__WEBPACK_IMPORTED_MODULE_0__.default)('v5', 0x50, _sha1_js__WEBPACK_IMPORTED_MODULE_1__.default);
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (v5);

/***/ }),
/* 78 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var crypto__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(63);
/* harmony import */ var crypto__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(crypto__WEBPACK_IMPORTED_MODULE_0__);


function sha1(bytes) {
  if (Array.isArray(bytes)) {
    bytes = Buffer.from(bytes);
  } else if (typeof bytes === 'string') {
    bytes = Buffer.from(bytes, 'utf8');
  }

  return crypto__WEBPACK_IMPORTED_MODULE_0___default().createHash('sha1').update(bytes).digest();
}

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (sha1);

/***/ }),
/* 79 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = ('00000000-0000-0000-0000-000000000000');

/***/ }),
/* 80 */
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _validate_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(70);


function version(uuid) {
  if (!(0,_validate_js__WEBPACK_IMPORTED_MODULE_0__.default)(uuid)) {
    throw TypeError('Invalid UUID');
  }

  return parseInt(uuid.substr(14, 1), 16);
}

/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (version);

/***/ }),
/* 81 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.Keychain = void 0;
const vscode = __webpack_require__(1);
const SERVICE_ID = `microsoft-todo-unofficial.login`;
class Keychain {
    constructor(context) {
        this.context = context;
    }
    setToken(token) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                return yield this.context.secrets.store(SERVICE_ID, token);
            }
            catch (e) {
                console.error(`Setting token failed: ${e}`);
                // Temporary fix for #94005
                // This happens when processes write simulatenously to the keychain, most
                // likely when trying to refresh the token. Ignore the error since additional
                // writes after the first one do not matter. Should actually be fixed upstream.
                if (e.message === 'The specified item already exists in the keychain.') {
                    return;
                }
                const troubleshooting = "Troubleshooting Guide";
                const result = yield vscode.window.showErrorMessage(`Writing login information to the keychain failed with error '${e.message}'.`, troubleshooting);
                if (result === troubleshooting) {
                    vscode.env.openExternal(vscode.Uri.parse('https://code.visualstudio.com/docs/editor/settings-sync#_troubleshooting-keychain-issues'));
                }
            }
        });
    }
    getToken() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                return yield this.context.secrets.get(SERVICE_ID);
            }
            catch (e) {
                // Ignore
                console.error(`Getting token failed: ${e}`);
                return Promise.resolve(undefined);
            }
        });
    }
    deleteToken() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                return yield this.context.secrets.delete(SERVICE_ID);
            }
            catch (e) {
                // Ignore
                console.error(`Deleting token failed: ${e}`);
                return Promise.resolve(undefined);
            }
        });
    }
}
exports.Keychain = Keychain;


/***/ }),
/* 82 */
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.createServer = exports.startServer = void 0;
const http = __webpack_require__(56);
const url = __webpack_require__(57);
const fs = __webpack_require__(83);
const path = __webpack_require__(84);
/**
 * Asserts that the argument passed in is neither undefined nor null.
 */
function assertIsDefined(arg) {
    if (typeof (arg) === 'undefined' || arg === null) {
        throw new Error('Assertion Failed: argument is undefined or null');
    }
    return arg;
}
function startServer(server) {
    return __awaiter(this, void 0, void 0, function* () {
        let portTimer;
        function cancelPortTimer() {
            clearTimeout(portTimer);
        }
        const port = new Promise((resolve, reject) => {
            portTimer = setTimeout(() => {
                reject(new Error('Timeout waiting for port'));
            }, 5000);
            server.on('listening', () => {
                const address = server.address();
                if (typeof address === 'string') {
                    resolve(address);
                }
                else {
                    resolve(assertIsDefined(address).port.toString());
                }
            });
            server.on('error', _ => {
                reject(new Error('Error listening to server'));
            });
            server.on('close', () => {
                reject(new Error('Closed'));
            });
            server.listen(0);
        });
        port.then(cancelPortTimer, cancelPortTimer);
        return port;
    });
}
exports.startServer = startServer;
function sendFile(res, filepath, contentType) {
    fs.readFile(filepath, (err, body) => {
        if (err) {
            console.error(err);
            res.writeHead(404);
            res.end();
        }
        else {
            res.writeHead(200, {
                'Content-Length': body.length,
                'Content-Type': contentType
            });
            res.end(body);
        }
    });
}
function callback(nonce, reqUrl) {
    return __awaiter(this, void 0, void 0, function* () {
        const query = reqUrl.query;
        if (!query || typeof query === 'string') {
            throw new Error('No query received.');
        }
        let error = query.error_description || query.error;
        if (!error) {
            const state = query.state || '';
            const receivedNonce = (state.split(',')[1] || '').replace(/ /g, '+');
            if (receivedNonce !== nonce) {
                error = 'Nonce does not match.';
            }
        }
        const code = query.code;
        if (!error && code) {
            return code;
        }
        throw new Error(error || 'No code received.');
    });
}
function createServer(nonce) {
    let deferredRedirect;
    const redirectPromise = new Promise((resolve, reject) => deferredRedirect = { resolve, reject });
    let deferredCode;
    const codePromise = new Promise((resolve, reject) => deferredCode = { resolve, reject });
    const codeTimer = setTimeout(() => {
        deferredCode.reject(new Error('Timeout waiting for code'));
    }, 5 * 60 * 1000);
    function cancelCodeTimer() {
        clearTimeout(codeTimer);
    }
    const server = http.createServer(function (req, res) {
        const reqUrl = url.parse(req.url, /* parseQueryString */ true);
        switch (reqUrl.pathname) {
            case '/signin':
                const receivedNonce = (reqUrl.query.nonce || '').replace(/ /g, '+');
                if (receivedNonce === nonce) {
                    deferredRedirect.resolve({ req, res });
                }
                else {
                    const err = new Error('Nonce does not match.');
                    deferredRedirect.resolve({ err, res });
                }
                break;
            case '/':
                sendFile(res, path.join(__dirname, '../media/auth.html'), 'text/html; charset=utf-8');
                break;
            case '/auth.css':
                sendFile(res, path.join(__dirname, '../media/auth.css'), 'text/css; charset=utf-8');
                break;
            case '/callback':
                deferredCode.resolve(callback(nonce, reqUrl)
                    .then(code => ({ code, res }), err => ({ err, res })));
                break;
            default:
                res.writeHead(404);
                res.end();
                break;
        }
    });
    codePromise.then(cancelCodeTimer, cancelCodeTimer);
    return {
        server,
        redirectPromise,
        codePromise
    };
}
exports.createServer = createServer;


/***/ }),
/* 83 */
/***/ ((module) => {

"use strict";
module.exports = require("fs");;

/***/ }),
/* 84 */
/***/ ((module) => {

"use strict";
module.exports = require("path");;

/***/ })
/******/ 	]);
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	var __webpack_exports__ = __webpack_require__(0);
/******/ 	module.exports = __webpack_exports__;
/******/ 	
/******/ })()
;
//# sourceMappingURL=extension.js.map