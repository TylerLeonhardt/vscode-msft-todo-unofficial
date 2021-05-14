import * as vscode from 'vscode';
import { MSAService, onDidChangeSessions } from './MSAService';

export class MsaAuthProvider implements vscode.AuthenticationProvider {
    static id = 'msa';

    onDidChangeSessions: vscode.Event<vscode.AuthenticationProviderAuthenticationSessionsChangeEvent> = onDidChangeSessions.event;

    private msaService: MSAService; 

    constructor(context: vscode.ExtensionContext) {
        this.msaService = new MSAService(context);
        this.msaService.initialize();
    }

    initialize(): Thenable<void> {
        return this.msaService.initialize();
    }

    getSessions(scopes?: string[]): Thenable<readonly vscode.AuthenticationSession[]> {
        return this.msaService.getSessions(scopes?.sort());
    }

    async createSession(scopes: string[]): Promise<vscode.AuthenticationSession> {
        const session = await this.msaService.createSession(scopes.sort());
        onDidChangeSessions.fire({ added: [session], removed: [], changed: [] });
        return session;
    }

    async removeSession(sessionId: string): Promise<void> {
        try {
            const session = await this.msaService.removeSession(sessionId);
            if (session) {
                onDidChangeSessions.fire({ added: [], removed: [session], changed: [] });
            }
        } catch (e) {
            console.error(e);
        }
    }
}
