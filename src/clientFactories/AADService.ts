import * as vscode from "vscode";
import { IAuthService } from "./AuthService";

export class AADService implements IAuthService {
    private static scopes = ['Tasks.ReadWrite', 'offline_access'];

    private clearedSessions = false;
    initialize(): Promise<void> {
        throw new Error("Method not implemented.");
    }
    async createSession(): Promise<vscode.AuthenticationSession> {
        const session = await vscode.authentication.getSession('microsoft', AADService.scopes, {
            createIfNone: true
        });
        this.clearedSessions = false;
        return session;
    }

    async getSessions(scopes?: string[]): Promise<vscode.AuthenticationSession[]> {
        if (this.clearedSessions) {
            return [];
        }
        const session = await vscode.authentication.getSession('microsoft', scopes || AADService.scopes);
        return session ? [session] : [];
    }

    async removeSession(sessionId: string): Promise<vscode.AuthenticationSession | undefined> {
        await this.clearSessions();
        return undefined;
    }

    clearSessions(): Promise<void> {
        this.clearedSessions = true;
        return Promise.resolve(undefined);
    }
}