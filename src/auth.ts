export interface FastmailConfig {
  apiToken: string;
  baseUrl?: string;
}

export class FastmailAuth {
  private apiToken: string;
  private baseUrl: string;

  constructor(config: FastmailConfig) {
    this.apiToken = config.apiToken;
    this.baseUrl = config.baseUrl || 'https://api.fastmail.com';
  }

  getAuthHeaders(): Record<string, string> {
    return {
      'Authorization': `Bearer ${this.apiToken}`,
      'Content-Type': 'application/json'
    };
  }

  getSessionUrl(): string {
    return `${this.baseUrl}/jmap/session`;
  }

  getApiUrl(): string {
    return `${this.baseUrl}/jmap/api/`;
  }
}