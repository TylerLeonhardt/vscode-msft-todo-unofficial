/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as randomBytes from 'randombytes';
import * as querystring from 'querystring';
import { Buffer } from 'buffer';
import * as vscode from 'vscode';

import { v4 as uuid } from 'uuid';
import * as fetch from 'isomorphic-fetch';
import { Keychain } from './keychain';

const redirectUrl = 'https://extension-auth-redirect.azurewebsites.net/';
const loginEndpointUrl = 'https://login.microsoftonline.com/';
const clientId = 'a4fd7674-4ebd-4dbc-831c-338314dd459e';
const tenant = 'common';

/**
 * Represents a session of a currently logged in Microsoft user.
 */
export interface MicrosoftAuthenticationSession extends vscode.AuthenticationSession {
    /**
     * The id token.
     */
    idToken?: string;
}

interface IToken {
    accessToken?: string; // When unable to refresh due to network problems, the access token becomes undefined
    idToken?: string; // depending on the scopes can be either supplied or empty

    expiresIn?: number; // How long access token is valid, in seconds
    expiresAt?: number; // UNIX epoch time at which token will expire
    refreshToken: string;

    account: {
        label: string;
        id: string;
    };
    scope: string;
    sessionId: string; // The account id + the scope
}

interface ITokenClaims {
    tid: string;
    email?: string;
    unique_name?: string;
    preferred_username?: string;
    oid?: string;
    altsecid?: string;
    ipd?: string;
    scp: string;
}

interface IStoredSession {
    id: string;
    refreshToken: string;
    scope: string; // Scopes are alphabetized and joined with a space
    account: {
        label?: string;
        displayName?: string,
        id: string
    }
}

export interface ITokenResponse {
    access_token: string;
    expires_in: number;
    ext_expires_in: number;
    refresh_token: string;
    scope: string;
    token_type: string;
    id_token?: string;
}

export interface IMicrosoftTokens {
    accessToken: string;
    idToken?: string;
}

function toBase64UrlEncoding(base64string: string) {
    return base64string.replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_'); // Need to use base64url encoding
}

function parseQuery(uri: vscode.Uri) {
    return uri.query.split('&').reduce((prev: any, current) => {
        const queryString = current.split('=');
        prev[queryString[0]] = queryString[1];
        return prev;
    }, {});
}

async function sha256(s: string | Uint8Array): Promise<string> {
    return (require('crypto')).createHash('sha256').update(s).digest('base64');
}

export const onDidChangeSessions = new vscode.EventEmitter<vscode.AuthenticationProviderAuthenticationSessionsChangeEvent>();

export const REFRESH_NETWORK_FAILURE = 'Network failure';

class UriEventHandler extends vscode.EventEmitter<vscode.Uri> implements vscode.UriHandler {
    public handleUri(uri: vscode.Uri) {
        this.fire(uri);
    }
}

export class MSAService {
    private _tokens: IToken[] = [];
    private _refreshTimeouts: Map<string, NodeJS.Timeout> = new Map<string, NodeJS.Timeout>();
    private _uriHandler: UriEventHandler;
    private _disposables: vscode.Disposable[] = [];

    private _keychain: Keychain;

    // Used to keep track of current requests when not using the local server approach.
    private _pendingStates = new Map<string, string[]>();
    private _codeExchangePromises = new Map<string, Promise<vscode.AuthenticationSession>>();
    private _codeVerfifiers = new Map<string, string>();

    constructor(private _context: vscode.ExtensionContext) {
        this._keychain = new Keychain(_context);
        this._uriHandler = new UriEventHandler();
        this._disposables.push(vscode.window.registerUriHandler(this._uriHandler));
    }

    public async initialize(): Promise<void> {
        const storedData = await this._keychain.getToken();
        if (storedData) {
            try {
                const sessions = this.parseStoredData(storedData);
                const refreshes = sessions.map(async session => {
                    if (!session.refreshToken) {
                        return Promise.resolve();
                    }

                    try {
                        await this.refreshToken(session.refreshToken, session.scope, session.id);
                    } catch (e) {
                        if (e.message === REFRESH_NETWORK_FAILURE) {
                            const didSucceedOnRetry = await this.handleRefreshNetworkError(session.id, session.refreshToken, session.scope);
                            if (!didSucceedOnRetry) {
                                this._tokens.push({
                                    accessToken: undefined,
                                    refreshToken: session.refreshToken,
                                    account: {
                                        label: session.account.label ?? session.account.displayName!,
                                        id: session.account.id
                                    },
                                    scope: session.scope,
                                    sessionId: session.id
                                });
                                this.pollForReconnect(session.id, session.refreshToken, session.scope);
                            }
                        } else {
                            await this.removeSession(session.id);
                        }
                    }
                });

                await Promise.all(refreshes);
            } catch (e) {
                console.info('Failed to initialize stored data');
                await this.clearSessions();
            }
        }
        this._disposables.push(this._context.secrets.onDidChange(() => this.checkForUpdates));
    }

    private parseStoredData(data: string): IStoredSession[] {
        return JSON.parse(data);
    }

    private async storeTokenData(): Promise<void> {
        const serializedData: IStoredSession[] = this._tokens.map(token => {
            return {
                id: token.sessionId,
                refreshToken: token.refreshToken,
                scope: token.scope,
                account: token.account
            };
        });

        await this._keychain.setToken(JSON.stringify(serializedData));
    }

    private async checkForUpdates(): Promise<void> {
        const added: vscode.AuthenticationSession[] = [];
        let removed: vscode.AuthenticationSession[] = [];
        const storedData = await this._keychain.getToken();
        if (storedData) {
            try {
                const sessions = this.parseStoredData(storedData);
                let promises = sessions.map(async session => {
                    const matchesExisting = this._tokens.some(token => token.scope === session.scope && token.sessionId === session.id);
                    if (!matchesExisting && session.refreshToken) {
                        try {
                            const token = await this.refreshToken(session.refreshToken, session.scope, session.id);
                            added.push(this.convertToSessionSync(token));
                        } catch (e) {
                            if (e.message === REFRESH_NETWORK_FAILURE) {
                                // Ignore, will automatically retry on next poll.
                            } else {
                                await this.removeSession(session.id);
                            }
                        }
                    }
                });

                promises = promises.concat(this._tokens.map(async token => {
                    const matchesExisting = sessions.some(session => token.scope === session.scope && token.sessionId === session.id);
                    if (!matchesExisting) {
                        await this.removeSession(token.sessionId);
                        removed.push(this.convertToSessionSync(token));
                    }
                }));

                await Promise.all(promises);
            } catch (e) {
                console.error(e.message);
                // if data is improperly formatted, remove all of it and send change event
                removed = this._tokens.map(this.convertToSessionSync);
                this.clearSessions();
            }
        } else {
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
            onDidChangeSessions.fire({ added: added, removed: removed, changed: [] });
        }
    }

    /**
     * Return a session object without checking for expiry and potentially refreshing.
     * @param token The token information.
     */
    private convertToSessionSync(token: IToken): MicrosoftAuthenticationSession {
        return {
            id: token.sessionId,
            accessToken: token.accessToken!,
            idToken: token.idToken,
            account: token.account,
            scopes: token.scope.split(' ')
        };
    }

    private async convertToSession(token: IToken): Promise<MicrosoftAuthenticationSession> {
        const resolvedTokens = await this.resolveAccessAndIdTokens(token);
        return {
            id: token.sessionId,
            accessToken: resolvedTokens.accessToken,
            idToken: resolvedTokens.idToken,
            account: token.account,
            scopes: token.scope.split(' ')
        };
    }

    private async resolveAccessAndIdTokens(token: IToken): Promise<IMicrosoftTokens> {
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
            const refreshedToken = await this.refreshToken(token.refreshToken, token.scope, token.sessionId);
            if (refreshedToken.accessToken) {
                return {
                    accessToken: refreshedToken.accessToken,
                    idToken: refreshedToken.idToken
                };
            } else {
                throw new Error();
            }
        } catch (e) {
            throw new Error('Unavailable due to network problems');
        }
    }

    private getTokenClaims(jwt: string): ITokenClaims {
        try {
            return JSON.parse(Buffer.from(jwt.split('.')[1], 'base64').toString());
        } catch (e) {
            console.error(e.message);
            throw new Error('Unable to read token claims');
        }
    }

    get sessions(): Promise<vscode.AuthenticationSession[]> {
        return Promise.all(this._tokens.map(token => this.convertToSession(token)));
    }

    async getSessions(scopes?: string[]): Promise<vscode.AuthenticationSession[]> {
        if (!scopes) {
            return this.sessions;
        }

        const orderedScopes = scopes.sort().join(' ');
        const matchingTokens = this._tokens.filter(token => token.scope === orderedScopes);
        return Promise.all(matchingTokens.map(token => this.convertToSession(token)));
    }

    public async createSession(scopes: string[]): Promise<vscode.AuthenticationSession> {
        console.info('Logging in...');
        return await this.loginWithoutLocalServer(scopes);
    }

    public dispose(): void {
        this._disposables.forEach(disposable => disposable.dispose());
        this._disposables = [];
    }

    private getCallbackEnvironment(callbackUri: vscode.Uri): string {
        if (callbackUri.authority.endsWith('.workspaces.github.com') || callbackUri.authority.endsWith('.github.dev')) {
            return `${callbackUri.authority},`;
        }

        switch (callbackUri.authority) {
            case 'online.visualstudio.com':
                return 'vso,';
            case 'online-ppe.core.vsengsaas.visualstudio.com':
                return 'vsoppe,';
            case 'online.dev.core.vsengsaas.visualstudio.com':
                return 'vsodev,';
            default:
                return `${callbackUri.scheme},`;
        }
    }

    private async loginWithoutLocalServer(scopes: string[]): Promise<vscode.AuthenticationSession> {
        const callbackUri = await vscode.env.asExternalUri(vscode.Uri.parse(`${vscode.env.uriScheme}://tylerleonhardt.msft-todo-unofficial`));
        const nonce = randomBytes(16).toString('base64');
        const port = (callbackUri.authority.match(/:([0-9]*)$/) || [])[1] || (callbackUri.scheme === 'https' ? 443 : 80);
        const callbackEnvironment = this.getCallbackEnvironment(callbackUri);
        const state = `${callbackEnvironment}${port},${encodeURIComponent(nonce)},${encodeURIComponent(callbackUri.query)}`;
        const signInUrl = `${loginEndpointUrl}${tenant}/oauth2/v2.0/authorize`;
        let uri = vscode.Uri.parse(signInUrl);
        const codeVerifier = toBase64UrlEncoding(randomBytes(32).toString('base64'));
        const codeChallenge = toBase64UrlEncoding(await sha256(codeVerifier));
        const scopeStr = scopes.join(' ');
        uri = uri.with({
            query: `response_type=code&client_id=${encodeURIComponent(clientId)}&response_mode=query&redirect_uri=${redirectUrl}&state=${state}&scope=${scopeStr}&prompt=select_account&code_challenge_method=S256&code_challenge=${codeChallenge}`
        });
        vscode.env.openExternal(uri);

        const timeoutPromise = new Promise((_: (value: vscode.AuthenticationSession) => void, reject) => {
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
    }

    private async handleCodeResponse(scopeStr: string): Promise<vscode.AuthenticationSession> {
        let uriEventListener: vscode.Disposable;
        return new Promise((resolve: (value: vscode.AuthenticationSession) => void, reject) => {
            uriEventListener = this._uriHandler.event(async (uri: vscode.Uri) => {
                try {
                    const query = parseQuery(uri);
                    const code = query.code;

                    const acceptedStates = this._pendingStates.get(scopeStr) || [];
                    // Workaround double encoding issues of state in web
                    if (!acceptedStates.includes(query.state) && !acceptedStates.includes(decodeURIComponent(query.state))) {
                        throw new Error('State does not match.');
                    }

                    const verifier = this._codeVerfifiers.get(query.state) ?? this._codeVerfifiers.get(decodeURIComponent(query.state));
                    if (!verifier) {
                        throw new Error('No available code verifier');
                    }

                    const token = await this.exchangeCodeForToken(code, verifier, scopeStr);
                    this.setToken(token, scopeStr);

                    const session = await this.convertToSession(token);
                    resolve(session);
                } catch (err) {
                    reject(err);
                }
            });
        }).then(result => {
            uriEventListener.dispose();
            return result;
        }).catch(err => {
            uriEventListener.dispose();
            throw err;
        });
    }

    private async setToken(token: IToken, scope: string): Promise<void> {
        const existingTokenIndex = this._tokens.findIndex(t => t.sessionId === token.sessionId);
        if (existingTokenIndex > -1) {
            this._tokens.splice(existingTokenIndex, 1, token);
        } else {
            this._tokens.push(token);
        }

        this.clearSessionTimeout(token.sessionId);

        if (token.expiresIn) {
            this._refreshTimeouts.set(token.sessionId, setTimeout(async () => {
                try {
                    const refreshedToken = await this.refreshToken(token.refreshToken, scope, token.sessionId);
                    onDidChangeSessions.fire({ added: [], removed: [], changed: [this.convertToSessionSync(refreshedToken)] });
                } catch (e) {
                    if (e.message === REFRESH_NETWORK_FAILURE) {
                        const didSucceedOnRetry = await this.handleRefreshNetworkError(token.sessionId, token.refreshToken, scope);
                        if (!didSucceedOnRetry) {
                            this.pollForReconnect(token.sessionId, token.refreshToken, token.scope);
                        }
                    } else {
                        await this.removeSession(token.sessionId);
                        onDidChangeSessions.fire({ added: [], removed: [this.convertToSessionSync(token)], changed: [] });
                    }
                }
            }, 1000 * (token.expiresIn - 30)));
        }

        this.storeTokenData();
    }

    private getTokenFromResponse(json: ITokenResponse, scope: string, existingId?: string): IToken {
		let claims = undefined;

		try {
			claims = this.getTokenClaims(json.access_token);
		} catch (e) {
			if (json.id_token) {
				console.log('Failed to fetch token claims from access_token. Attempting to parse id_token instead');
				claims = this.getTokenClaims(json.id_token);
			} else {
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
			sessionId: uuid(),
			account: {
				label: claims.email || claims.unique_name || claims.preferred_username || 'user@example.com',
				id: `${claims.tid}/${(claims.oid || (claims.altsecid || '' + claims.ipd || ''))}`
			}
		};
	}

    private async exchangeCodeForToken(code: string, codeVerifier: string, scopeStr: string): Promise<IToken> {
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

            const proxyEndpoints: { [providerId: string]: string } | undefined = await vscode.commands.executeCommand('workbench.getCodeExchangeProxyEndpoints');
            const endpoint = proxyEndpoints && proxyEndpoints['microsoft'] || `${loginEndpointUrl}${tenant}/oauth2/v2.0/token`;

            const result = await fetch(endpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Content-Length': postData.length.toString()
                },
                body: postData
            });

            if (result.ok) {
                console.info('Exchanging login code for token success');
                const json = await result.json();
                return this.getTokenFromResponse(json, scopeStr);
            } else {
                console.error('Exchanging login code for token failed');
                throw new Error('Unable to login.');
            }
        } catch (e) {
            console.error(e.message);
            throw e;
        }
    }

    private async refreshToken(refreshToken: string, scope: string, sessionId: string): Promise<IToken> {
        console.info('Refreshing token...');
        const postData = querystring.stringify({
            refresh_token: refreshToken,
            client_id: clientId,
            grant_type: 'refresh_token',
            scope: scope
        });

        let result: Response;
        try {
            result = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Content-Length': postData.length.toString()
                },
                body: postData
            });
        } catch (e) {
            console.error('Refreshing token failed');
            throw new Error(REFRESH_NETWORK_FAILURE);
        }

        try {
            if (result.ok) {
                const json = await result.json();
                const token = this.getTokenFromResponse(json, scope, sessionId);
                this.setToken(token, scope);
                console.info('Token refresh success');
                return token;
            } else {
                throw new Error('Bad request.');
            }
        } catch (e) {
            vscode.window.showErrorMessage("You have been signed out because reading stored authentication information failed.");
            console.error(`Refreshing token failed: ${result.statusText}`);
            throw new Error('Refreshing token failed');
        }
    }

    private clearSessionTimeout(sessionId: string): void {
        const timeout = this._refreshTimeouts.get(sessionId);
        if (timeout) {
            clearTimeout(timeout);
            this._refreshTimeouts.delete(sessionId);
        }
    }

    private removeInMemorySessionData(sessionId: string): IToken | undefined {
        const tokenIndex = this._tokens.findIndex(token => token.sessionId === sessionId);
        let token: IToken | undefined;
        if (tokenIndex > -1) {
            token = this._tokens[tokenIndex];
            this._tokens.splice(tokenIndex, 1);
        }

        this.clearSessionTimeout(sessionId);
        return token;
    }

    private pollForReconnect(sessionId: string, refreshToken: string, scope: string): void {
        this.clearSessionTimeout(sessionId);

        this._refreshTimeouts.set(sessionId, setTimeout(async () => {
            try {
                const refreshedToken = await this.refreshToken(refreshToken, scope, sessionId);
                onDidChangeSessions.fire({ added: [], removed: [], changed: [this.convertToSessionSync(refreshedToken)] });
            } catch (e) {
                this.pollForReconnect(sessionId, refreshToken, scope);
            }
        }, 1000 * 60 * 30));
    }

    private handleRefreshNetworkError(sessionId: string, refreshToken: string, scope: string, attempts: number = 1): Promise<boolean> {
        return new Promise((resolve, _) => {
            if (attempts === 3) {
                console.error('Token refresh failed after 3 attempts');
                return resolve(false);
            }

            const delayBeforeRetry = 5 * attempts * attempts;

            this.clearSessionTimeout(sessionId);

            this._refreshTimeouts.set(sessionId, setTimeout(async () => {
                try {
                    const refreshedToken = await this.refreshToken(refreshToken, scope, sessionId);
                    onDidChangeSessions.fire({ added: [], removed: [], changed: [this.convertToSessionSync(refreshedToken)] });
                    return resolve(true);
                } catch (e) {
                    return resolve(await this.handleRefreshNetworkError(sessionId, refreshToken, scope, attempts + 1));
                }
            }, 1000 * delayBeforeRetry));
        });
    }

    public async removeSession(sessionId: string): Promise<vscode.AuthenticationSession | undefined> {
        console.info(`Logging out of session '${sessionId}'`);
        const token = this.removeInMemorySessionData(sessionId);
        let session: vscode.AuthenticationSession | undefined;
        if (token) {
            session = this.convertToSessionSync(token);
        }

        if (this._tokens.length === 0) {
            await this._keychain.deleteToken();
        } else {
            this.storeTokenData();
        }

        return session;
    }

    public async clearSessions() {
        console.info('Logging out of all sessions');
        this._tokens = [];
        await this._keychain.deleteToken();

        this._refreshTimeouts.forEach(timeout => {
            clearTimeout(timeout);
        });

        this._refreshTimeouts.clear();
    }
}
