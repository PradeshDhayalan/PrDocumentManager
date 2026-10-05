import { DocumentRow, T } from '../types/Documents';
export function formatSize(kb: number, t: T): string {
  return kb >= 1024
    ? t('size.mb', { size: (kb / 1024).toFixed(1) })
    : t('size.kb', { size: Math.max(1, Math.ceil(kb)) });
}
export function formatDate(value: string, t: T): string {
  if (!value) return t('empty.value');
  return new Intl.DateTimeFormat(undefined, { dateStyle: 'medium' }).format(
    new Date(value.length === 10 ? value + 'T12:00:00' : value),
  );
}
export function relativeDate(value: string, t: T): string {
  const minutes = Math.max(0, Math.floor((Date.now() - Date.parse(value)) / 60000));
  if (minutes < 60) return t('date.minutes', { count: Math.max(1, minutes) });
  if (minutes < 1440) return t('date.hours', { count: Math.floor(minutes / 60) });
  return formatDate(value, t);
}
export function extension(name: string): string {
  return /\.([^.]+)$/.exec(name)?.[1].toLowerCase() || '';
}
export function fileKind(doc: DocumentRow): string {
  const ext = extension(doc.name);
  if (doc.contentType.startsWith('image/')) return 'image';
  if (ext === 'pdf') return 'pdf';
  if (['doc', 'docx'].includes(ext)) return 'word';
  if (['xls', 'xlsx', 'csv'].includes(ext)) return 'excel';
  if (['ppt', 'pptx'].includes(ext)) return 'powerpoint';
  if (doc.contentType.startsWith('text/') || ['json', 'md', 'txt'].includes(ext)) return 'text';
  if (doc.contentType.startsWith('video/')) return 'video';
  return doc.sizeKb ? 'generic' : 'link';
}
export function sanitizeFilename(name: string): string {
  const match = /^(.*?)(\.[^.]+)?$/.exec(name.trim()),
    ext = (match?.[2] || '').toLowerCase();
  let base =
    (match?.[1] || 'file')
      .split('')
      .map((char) => (char.charCodeAt(0) < 32 || char.charCodeAt(0) === 127 ? ' ' : char))
      .join('')
      .replace(/["*:<>?/\\|]/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/[. ]+$/, '')
      .trim() || 'file';
  if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(base)) base = '_' + base;
  return base.slice(0, 255 - ext.length).replace(/[. ]+$/, '') + ext;
}
export function validateSharePointUrl(input: string, hosts: string[]): string {
  const url = new URL(input.trim());
  if (
    url.protocol !== 'https:' ||
    url.username ||
    url.password ||
    url.href.length > 2000 ||
    !hosts.some((host) =>
      host.startsWith('*.')
        ? url.hostname.endsWith(host.slice(1)) && url.hostname !== host.slice(2)
        : url.hostname === host.toLowerCase(),
    )
  )
    throw new Error('link.invalid');
  return url.href;
}
export function newGuid(): string {
  return crypto.randomUUID();
}
