export interface FastmailConfig {
  apiToken: string;
  baseUrl?: string;
}

function normalizeBaseUrl(input?: string): string {
  const DEFAULT = 'https://api.fastmail.com';
  if (!input) return DEFAULT;
  let url = input.trim();
  if (!url) return DEFAULT;
  if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(url)) {
    url = 'https://' + url;
  }
  url = url.replace(/\/+$/, '');
  return url;
}

export class FastmailAuth {
  private apiToken: string;
  private baseUrl: string;

  constructor(config: FastmailConfig) {
    this.apiToken = config.apiToken;
    this.baseUrl = normalizeBaseUrl(config.baseUrl);
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