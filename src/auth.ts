import { validateFastmailUrl } from './url-validation.js';

export interface FastmailConfig {
  apiToken: string;
  baseUrl?: string;
  allowUnsafeBaseUrl?: boolean;
}

function normalizeBaseUrl(input: string | undefined, allowUnsafe: boolean): string {
  const DEFAULT = 'https://api.fastmail.com';
  if (!input) return DEFAULT;
  let url = input.trim();
  if (!url) return DEFAULT;
  if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(url)) {
    url = 'https://' + url;
  }
  url = url.replace(/\/+$/, '');
  // Reject any URL that isn't HTTPS+approved-origin (or unsafe-opted-in HTTPS).
  validateFastmailUrl(url, 'FASTMAIL_BASE_URL', allowUnsafe);
  return url;
}

export class FastmailAuth {
  private apiToken: string;
  private baseUrl: string;
  private allowUnsafeBaseUrl: boolean;

  constructor(config: FastmailConfig) {
    this.apiToken = config.apiToken;
    this.allowUnsafeBaseUrl = config.allowUnsafeBaseUrl ?? false;
    this.baseUrl = normalizeBaseUrl(config.baseUrl, this.allowUnsafeBaseUrl);
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

  getAllowUnsafe(): boolean {
    return this.allowUnsafeBaseUrl;
  }
}
