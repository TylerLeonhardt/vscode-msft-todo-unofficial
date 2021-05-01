import * as vscode from 'vscode';

export interface IAuthService {
    initialize(): Promise<void>;
    createSession(): Promise<vscode.AuthenticationSession>;
    getSessions(scopes?: string[]): Promise<vscode.AuthenticationSession[]>
    removeSession(sessionId: string): Promise<vscode.AuthenticationSession | undefined>;
    clearSessions(): Promise<void>;
}
