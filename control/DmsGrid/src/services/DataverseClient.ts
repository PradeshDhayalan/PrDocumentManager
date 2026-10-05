import { Page } from '../types/Documents';
export type ErrorKind =
  | 'permission'
  | 'throttle'
  | 'validation'
  | 'conflict'
  | 'notFound'
  | 'network'
  | 'unknown';
export class ApiError extends Error {
  constructor(
    public kind: ErrorKind,
    public status: number,
    public correlationId = '',
    public code = '',
  ) {
    super(kind);
  }
}
export function aborted(error: unknown): boolean {
  return error instanceof DOMException && error.name === 'AbortError';
}
function delay(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new DOMException('Aborted', 'AbortError'));
      return;
    }
    const onAbort = () => {
      clearTimeout(timer);
      reject(new DOMException('Aborted', 'AbortError'));
    };
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort);
      resolve();
    }, ms);
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}
// TODO(SPIKE-S4): this boundary can be replaced by a context.webAPI adapter after real-org verification.
export class DataverseClient {
  public readonly base: string;
  constructor(
    clientUrl: string,
    public pageSize = 50,
  ) {
    this.base = clientUrl.replace(/\/$/, '') + '/api/data/v9.2/';
  }
  async request<T>(
    path: string,
    method = 'GET',
    body?: unknown,
    signal?: AbortSignal,
    extra: Record<string, string> = {},
  ): Promise<T> {
    const url = new URL(path, this.base);
    if (url.origin !== new URL(this.base).origin || !url.pathname.startsWith('/api/data/v9.2/'))
      throw new ApiError('validation', 400);
    for (let attempt = 0; attempt <= 5; attempt++) {
      let response: Response;
      try {
        response = await fetch(url.toString(), {
          method,
          credentials: 'same-origin',
          signal,
          headers: {
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            Accept: 'application/json',
            'Content-Type': 'application/json; charset=utf-8',
            Prefer: `odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=${this.pageSize}`,
            ...extra,
          },
          body: body === undefined ? undefined : JSON.stringify(body),
        });
      } catch (error) {
        if (aborted(error) || signal?.aborted) throw new DOMException('Aborted', 'AbortError');
        if (attempt === 5) throw new ApiError('network', 0);
        await delay(Math.min(8000, 500 * 2 ** attempt), signal);
        continue;
      }
      if (response.ok) {
        if (response.status === 204) return undefined as T;
        return (await response.json()) as T;
      }
      if ((response.status === 429 || response.status >= 500) && attempt < 5) {
        const retry = response.headers.get('Retry-After');
        const seconds = retry ? Number(retry) : NaN;
        const ms = Number.isFinite(seconds)
          ? seconds * 1000
          : retry
            ? Math.max(0, Date.parse(retry) - Date.now())
            : 500 * 2 ** attempt;
        await response.body?.cancel();
        await delay(Math.min(60000, Math.max(0, ms)), signal);
        continue;
      }
      const kind: ErrorKind =
        response.status === 401 || response.status === 403
          ? 'permission'
          : response.status === 429
            ? 'throttle'
            : response.status === 412
              ? 'conflict'
              : response.status === 400 || response.status === 413
                ? 'validation'
                : response.status === 404
                  ? 'notFound'
                  : 'unknown';
      let code = '';
      try {
        const error = (await response.json()) as { error?: { code?: string } };
        code = error.error?.code || '';
      } catch {
        /* Errors may have an empty body. */
      }
      throw new ApiError(
        kind,
        response.status,
        response.headers.get('x-ms-service-request-id') || response.headers.get('req_id') || '',
        code,
      );
    }
    throw new ApiError('unknown', 0);
  }
  get<T>(path: string, signal?: AbortSignal): Promise<T> {
    return this.request<T>(path, 'GET', undefined, signal);
  }
  getPage<T>(path: string, signal?: AbortSignal): Promise<Page<T>> {
    return this.get<Page<T>>(path, signal);
  }
  post<T>(path: string, body: unknown, signal?: AbortSignal): Promise<T> {
    return this.request<T>(path, 'POST', body, signal);
  }
  patch<T>(path: string, body: unknown, etag: string, signal?: AbortSignal): Promise<T> {
    return this.request<T>(path, 'PATCH', body, signal, { 'If-Match': etag });
  }
  delete(path: string, etag: string, signal?: AbortSignal): Promise<void> {
    return this.request<void>(path, 'DELETE', undefined, signal, { 'If-Match': etag });
  }
  action<T>(name: string, body: unknown, signal?: AbortSignal): Promise<T> {
    return this.post<T>(name, body, signal);
  }
}
