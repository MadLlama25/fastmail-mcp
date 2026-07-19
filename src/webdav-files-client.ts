import { createObject, davRequest, getBasicAuthHeaders } from 'tsdav';

/**
 * Minimal WebDAV file-storage client for saving attachments to cloud storage
 * (Fastmail Files, Nextcloud, or any WebDAV server).
 *
 * Deliberately NOT built on tsdav's DAVClient account machinery — plain file
 * collections need no account discovery. Uses the low-level helpers with an
 * explicit Basic Authorization header.
 *
 * SECURITY MODEL: the base URL and credentials come exclusively from server
 * configuration (env). Tools may only ever supply a RELATIVE remote path,
 * which is validated segment-by-segment and encoded AFTER validation — a
 * prompt-injected caller cannot redirect uploads to another host or smuggle
 * traversal sequences through percent-encoding.
 */

export interface WebDAVConfig {
  baseUrl: string; // validated HTTPS collection URL, normalized to trailing slash
  username: string;
  password: string;
}

/**
 * Validate a caller-supplied relative remote path into clean segments.
 * Rejects: control chars/NUL, backslashes, absolute paths, URL smuggling,
 * empty segments (foo//bar, trailing /), '.', '..', oversized segments/paths.
 * Returns the raw (unencoded) segment list; callers must encode each segment.
 */
export function sanitizeRemotePath(remotePath: string): string[] {
  if (typeof remotePath !== 'string' || remotePath.length === 0) {
    throw new Error('remotePath is required');
  }
  if (remotePath.length > 1024) {
    throw new Error('remotePath is too long (max 1024 characters)');
  }
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001f\u007f]/.test(remotePath)) {
    throw new Error('remotePath contains control characters');
  }
  if (remotePath.includes('\\')) {
    throw new Error('remotePath must use forward slashes');
  }
  if (remotePath.startsWith('/')) {
    throw new Error('remotePath must be relative (no leading slash)');
  }
  if (remotePath.includes('://')) {
    throw new Error('remotePath must be a path, not a URL');
  }
  const segments = remotePath.split('/');
  if (segments.length > 32) {
    throw new Error('remotePath has too many segments (max 32)');
  }
  for (const segment of segments) {
    if (segment === '') {
      throw new Error('remotePath contains an empty segment (double or trailing slash)');
    }
    if (segment === '.' || segment === '..') {
      throw new Error('remotePath must not contain . or .. segments');
    }
    if (Buffer.byteLength(segment, 'utf8') > 255) {
      throw new Error('remotePath segment exceeds 255 bytes');
    }
  }
  return segments;
}

export class WebDAVFilesClient {
  private config: WebDAVConfig;
  private headers: Record<string, string>;
  // Instance-held so tests can inject mocks, mirroring the caldav-client
  // test idiom of swapping the underlying transport.
  private dav = { createObject, davRequest };

  constructor(config: WebDAVConfig) {
    this.config = {
      ...config,
      baseUrl: config.baseUrl.endsWith('/') ? config.baseUrl : `${config.baseUrl}/`,
    };
    this.headers = getBasicAuthHeaders({ username: config.username, password: config.password });
  }

  /** Build the target URL from validated segments — the only URL constructor. */
  private urlFor(segments: string[]): string {
    return this.config.baseUrl + segments.map(encodeURIComponent).join('/');
  }

  get host(): string {
    return new URL(this.config.baseUrl).hostname;
  }

  /**
   * MKCOL each ancestor collection of the target. 201 = created,
   * 405 = already exists (fine); anything else is fatal.
   */
  private async ensureParents(segments: string[]): Promise<void> {
    for (let depth = 1; depth < segments.length; depth++) {
      const url = this.urlFor(segments.slice(0, depth)) + '/';
      const [response] = await this.dav.davRequest({
        url,
        init: { method: 'MKCOL', headers: this.headers, body: undefined as any },
        convertIncoming: false,
      });
      const status = (response as any).status ?? (response as any).statusCode;
      if (status !== 201 && status !== 405) {
        throw new Error(`Could not create remote directory (${segments.slice(0, depth).join('/')}): HTTP ${status}`);
      }
    }
  }

  /**
   * Upload a buffer to the validated relative path.
   * Default refuses to replace an existing remote file (If-None-Match: *,
   * mirroring the local download path's O_EXCL semantics); overwrite=true
   * replaces unconditionally.
   */
  async uploadBuffer(
    buffer: Buffer,
    remotePath: string,
    options: { contentType?: string; overwrite?: boolean; createParents?: boolean } = {},
  ): Promise<{ savedPath: string; host: string; bytes: number; contentType: string }> {
    const segments = sanitizeRemotePath(remotePath);
    const contentType = options.contentType || 'application/octet-stream';

    // Explicit existence pre-check when not overwriting. The PUT below also
    // carries If-None-Match: * for RFC-compliant servers, but live testing
    // showed Fastmail Files ignores that precondition and overwrites anyway —
    // so the guard cannot rely on the header alone.
    if (!options.overwrite) {
      const [probe] = await this.dav.davRequest({
        url: this.urlFor(segments),
        init: { method: 'PROPFIND', headers: { ...this.headers, Depth: '0' }, body: undefined as any },
        convertIncoming: false,
      });
      const probeStatus = (probe as any).status ?? (probe as any).statusCode;
      if (probeStatus !== 404) {
        throw new Error(`Remote file already exists: ${segments.join('/')}. Pass overwrite: true to replace it.`);
      }
    }

    if (options.createParents !== false) {
      await this.ensureParents(segments);
    }

    const response = await this.dav.createObject({
      url: this.urlFor(segments),
      data: new Uint8Array(buffer) as any,
      headers: {
        ...this.headers,
        'Content-Type': contentType,
        ...(options.overwrite ? {} : { 'If-None-Match': '*' }),
      },
    });

    if (response.status === 412) {
      throw new Error(`Remote file already exists: ${segments.join('/')}. Pass overwrite: true to replace it.`);
    }
    if (!response.ok) {
      throw new Error(`WebDAV upload failed: HTTP ${response.status} ${response.statusText}`);
    }

    return {
      savedPath: segments.join('/'),
      host: this.host,
      bytes: buffer.length,
      contentType,
    };
  }
}
