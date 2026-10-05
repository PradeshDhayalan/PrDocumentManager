import * as React from 'react';
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Field,
  Input,
  Text,
  Toolbar,
  ToolbarButton,
  Spinner,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { DocumentRow, T, ClientConfig } from '../types/Documents';
import { DocumentRepository, providerChoices, parallelLimit } from '../services/DocumentRepository';
import { resolveProvider, usable } from '../providers/registry';
import { DocumentThumbnail } from './FileVisuals';
import { fileKind, newGuid, sanitizeFilename, validateSharePointUrl } from '../services/format';
import { ApiError } from '../services/DataverseClient';
const useStyles = makeStyles({
  surface: { maxWidth: 'min(1000px,94vw)', width: '900px' },
  preview: {
    width: '100%',
    height: 'min(65vh,650px)',
    border: 'none',
    objectFit: 'contain',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  fallback: {
    height: '280px',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '16px',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  cover: { width: '230px', height: '160px', overflow: 'hidden' },
  text: {
    padding: '20px',
    height: 'min(60vh,600px)',
    overflow: 'auto',
    whiteSpace: 'pre-wrap',
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '13px',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  fields: { display: 'flex', flexDirection: 'column', gap: '16px', paddingTop: '12px' },
  hint: { fontSize: '12px', color: tokens.colorNeutralForeground2 },
});
export function PreviewDialog({
  row,
  repo,
  t,
  onClose,
  onDownload,
  onOpen,
}: {
  row: DocumentRow;
  repo: DocumentRepository;
  t: T;
  onClose: () => void;
  onDownload: () => void;
  onOpen: () => void;
}) {
  const s = useStyles(),
    [url, setUrl] = React.useState(''),
    [text, setText] = React.useState(''),
    [loading, setLoading] = React.useState(false),
    [failed, setFailed] = React.useState(false),
    kind = fileKind(row),
    provider = resolveProvider(row.provider, repo.client),
    inline =
      provider.capabilities.canPreviewInline &&
      ['pdf', 'image', 'text'].includes(kind) &&
      row.sizeKb <= 51200 &&
      usable(row, provider);
  React.useEffect(() => {
    if (!inline) return;
    const controller = new AbortController();
    let objectUrl = '';
    setLoading(true);
    void provider
      .getContent(row, undefined, controller.signal)
      .then(async (content) => {
        if (content.kind !== 'blob' || controller.signal.aborted) return;
        if (kind === 'text') {
          const value = await content.blob.text();
          if (!controller.signal.aborted) setText(value);
        } else {
          objectUrl = URL.createObjectURL(content.blob);
          if (!controller.signal.aborted) setUrl(objectUrl);
        }
        if (!controller.signal.aborted) setLoading(false);
      })
      .catch(() => {
        if (!controller.signal.aborted) {
          setFailed(true);
          setLoading(false);
        }
      });
    return () => {
      controller.abort();
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [row.id, row.etag, repo, inline, kind]);
  return (
    <Dialog open onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface className={s.surface}>
        <DialogBody>
          <DialogTitle>{row.name}</DialogTitle>
          <DialogContent>
            {loading ? (
              <Spinner label={t('preview.loading')} />
            ) : inline && !failed ? (
              kind === 'image' ? (
                <img
                  src={url}
                  className={s.preview}
                  alt={row.name}
                  onError={() => setFailed(true)}
                />
              ) : kind === 'pdf' ? (
                <iframe
                  src={url}
                  className={s.preview}
                  title={row.name}
                  onError={() => setFailed(true)}
                />
              ) : (
                <pre className={s.text}>{text}</pre>
              )
            ) : (
              <div className={s.fallback}>
                <div className={s.cover}>
                  <DocumentThumbnail row={row} repo={repo} t={t} />
                </div>
                <Text>
                  {t(
                    !provider.capabilities.ownsBinary && provider.capabilities.canOpen
                      ? 'preview.sharePoint'
                      : 'preview.unavailable',
                  )}
                </Text>
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <Toolbar>
              <ToolbarButton onClick={onClose}>{t('action.close')}</ToolbarButton>
              {usable(row, provider) && provider.capabilities.canDownload && (
                <ToolbarButton
                  onClick={provider.capabilities.ownsBinary ? onDownload : onOpen}
                  appearance="primary"
                >
                  {t(
                    provider.capabilities.ownsBinary ? 'action.download' : 'action.openSharePoint',
                  )}
                </ToolbarButton>
              )}
            </Toolbar>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
export function AddLinkDialog({
  repo,
  config,
  recordId,
  entityName,
  t,
  onClose,
  onAdded,
}: {
  repo: DocumentRepository;
  config: ClientConfig;
  recordId: string;
  entityName: string;
  t: T;
  onClose: () => void;
  onAdded: (row: DocumentRow) => void;
}) {
  const s = useStyles(),
    [url, setUrl] = React.useState(''),
    [name, setName] = React.useState(''),
    [saving, setSaving] = React.useState(false),
    [error, setError] = React.useState('');
  async function save() {
    setSaving(true);
    setError('');
    try {
      let valid: string;
      try {
        valid = validateSharePointUrl(url, config.sharePoint.allowedHosts);
      } catch {
        setError('link.invalid');
        setSaving(false);
        return;
      }
      const parsed = new URL(valid);
      let derived = '';
      try {
        derived = decodeURIComponent(
          parsed.pathname.split('/').filter(Boolean).pop() || parsed.hostname,
        );
      } catch {
        derived = parsed.hostname;
      }
      const row = await repo.create({
        dms_documentid: newGuid(),
        dms_name: sanitizeFilename(name.trim() || derived),
        dms_originalfilename: name.trim() || derived,
        dms_description: '',
        dms_regardingid: recordId,
        dms_regardingtype: entityName,
        dms_provider: providerChoices.SharePoint,
        dms_storageref: valid,
        dms_contenttype: 'application/x-sharepoint-link',
        dms_uploadstate: 100000001,
        dms_documenttype: 100000005,
        dms_documentstatus: 100000001,
      });
      onAdded(row);
      onClose();
    } catch (error) {
      setError(error instanceof ApiError ? 'error.' + error.kind : 'error.unknown');
    } finally {
      setSaving(false);
    }
  }
  return (
    <Dialog open onOpenChange={(_, data) => !data.open && !saving && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{t('action.addLink')}</DialogTitle>
          <DialogContent>
            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{t(error)}</MessageBarBody>
              </MessageBar>
            )}
            <div className={s.fields}>
              <Field label={t('link.url')} required>
                <Input
                  value={url}
                  placeholder={t('link.placeholder')}
                  onChange={(_, data) => setUrl(data.value)}
                />
              </Field>
              <Field label={t('link.name')}>
                <Input value={name} maxLength={255} onChange={(_, data) => setName(data.value)} />
              </Field>
              <Text className={s.hint}>{t('link.permissions')}</Text>
            </div>
          </DialogContent>
          <DialogActions>
            <Toolbar>
              <ToolbarButton disabled={saving} onClick={onClose}>
                {t('action.cancel')}
              </ToolbarButton>
              <ToolbarButton
                appearance="primary"
                disabled={saving || !url.trim()}
                onClick={() => void save()}
              >
                {t(saving ? 'saving' : 'action.addLink')}
              </ToolbarButton>
            </Toolbar>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
export function DeleteDialog({
  rows,
  repo,
  t,
  onClose,
  onDeleted,
}: {
  rows: DocumentRow[];
  repo: DocumentRepository;
  t: T;
  onClose: () => void;
  onDeleted: (rows: DocumentRow[]) => void;
}) {
  const [saving, setSaving] = React.useState(false),
    [remaining, setRemaining] = React.useState(rows),
    [error, setError] = React.useState('');
  async function remove() {
    setSaving(true);
    const results = await parallelLimit(remaining, 5, (row) => repo.delete(row));
    const deleted = remaining.filter((_, i) => results[i].status === 'fulfilled');
    onDeleted(deleted);
    const failures = remaining.filter((_, i) => results[i].status === 'rejected');
    if (failures.length) {
      setRemaining(failures);
      setError('delete.partial');
    } else onClose();
    setSaving(false);
  }
  const owns = remaining.some(
      (row) => resolveProvider(row.provider, repo.client).capabilities.ownsBinary,
    ),
    references = remaining.some(
      (row) => !resolveProvider(row.provider, repo.client).capabilities.ownsBinary,
    );
  return (
    <Dialog open onOpenChange={(_, data) => !data.open && !saving && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{t('delete.title', { count: remaining.length })}</DialogTitle>
          <DialogContent>
            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{t(error, { count: remaining.length })}</MessageBarBody>
              </MessageBar>
            )}
            <p>{owns && t('delete.owned')}</p>
            <p>{references && t('delete.reference')}</p>
          </DialogContent>
          <DialogActions>
            <Toolbar>
              <ToolbarButton disabled={saving} onClick={onClose}>
                {t('action.cancel')}
              </ToolbarButton>
              <ToolbarButton appearance="primary" disabled={saving} onClick={() => void remove()}>
                {t(saving ? 'deleting' : 'action.delete')}
              </ToolbarButton>
            </Toolbar>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
