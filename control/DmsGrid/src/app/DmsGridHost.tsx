import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  Text,
  Toolbar,
  ToolbarButton,
  Checkbox,
  Dropdown,
  Option,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  MessageBar,
  MessageBarBody,
  Skeleton,
  SkeletonItem,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components';
import {
  DocumentRegular,
  LibraryRegular,
  ChevronDownRegular,
  DismissRegular,
  ArrowSyncRegular,
  ArrowUploadRegular,
} from '@fluentui/react-icons';
import { HostContext } from '../types/HostContext';
import { DocumentRow, ClientConfig, Migration, Preset, ListQuery } from '../types/Documents';
import { createI18n } from '../services/i18n';
import { DataverseClient, ApiError, aborted } from '../services/DataverseClient';
import { DocumentRepository } from '../services/DocumentRepository';
import { UploadQueue } from '../services/UploadQueue';
import { useSelection } from '../hooks/useSelection';
import { resolveProvider, usable } from '../providers/registry';
import '../providers/NoteProvider';
import '../providers/SharePointProvider';
import { DocumentCommandBar, Action } from '../components/DocumentCommandBar';
import { DocumentList, DocumentTiles, ViewProps } from '../components/DocumentViews';
import { DetailsPane } from '../components/DetailsPane';
import { MetadataDrawer } from '../components/MetadataDrawer';
import { PreviewDialog, AddLinkDialog, DeleteDialog } from '../components/DocumentDialogs';
import { UploadTray } from '../components/UploadTray';
interface State {
  config: ClientConfig | null;
  rows: DocumentRow[];
  loading: boolean;
  total: number;
  nextCursor?: string;
  error: ApiError | null;
  search: string;
  preset: Preset;
  typeFilter: string;
  orderBy: string;
  view: 'list' | 'tiles';
  details: boolean;
  filters: boolean;
  expanded: boolean;
  compact: boolean;
  columns: string[];
  notice: string;
  dialog: 'preview' | 'edit' | 'delete' | 'link' | 'columns' | 'shortcuts' | null;
  dialogRows: DocumentRow[];
  migration: Migration | null;
  dragging: boolean;
}
type StateAction =
  | { type: 'set'; patch: Partial<State> }
  | { type: 'rows'; rows: DocumentRow[]; append?: boolean; total: number; nextCursor?: string }
  | { type: 'patch'; rows: DocumentRow[] }
  | { type: 'remove'; ids: string[] };
function reduce(state: State, action: StateAction): State {
  if (action.type === 'set') return { ...state, ...action.patch };
  if (action.type === 'rows') {
    const seen = new Set(state.rows.map((row) => row.id));
    return {
      ...state,
      rows: action.append
        ? [...state.rows, ...action.rows.filter((row) => !seen.has(row.id))]
        : action.rows,
      total: action.total,
      nextCursor: action.nextCursor,
      loading: false,
      error: null,
    };
  }
  if (action.type === 'remove')
    return {
      ...state,
      rows: state.rows.filter((row) => !action.ids.includes(row.id)),
      total: Math.max(0, state.total - action.ids.length),
    };
  const updates = new Map(action.rows.map((row) => [row.id, row])),
    existing = new Set(state.rows.map((row) => row.id)),
    added = action.rows.filter((row) => !existing.has(row.id));
  return {
    ...state,
    rows: [...added, ...state.rows.map((row) => updates.get(row.id) || row)],
    total: state.total + added.length,
  };
}
const useStyles = makeStyles({
  root: {
    minHeight: '420px',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    position: 'relative',
  },
  expanded: { position: 'fixed', inset: 0, zIndex: 30, overflow: 'hidden' },
  heading: { display: 'flex', alignItems: 'center', gap: '12px', padding: '24px 24px 20px' },
  title: { fontSize: '24px', fontWeight: tokens.fontWeightSemibold, lineHeight: '32px', margin: 0 },
  count: { fontSize: '14px', color: tokens.colorNeutralForeground2 },
  chip: { backgroundColor: tokens.colorBrandBackground2, color: tokens.colorBrandForeground1 },
  workspace: { display: 'flex', minHeight: 0, flex: 1 },
  main: { flex: 1, minWidth: 0, display: 'flex', flexDirection: 'column' },
  scroll: { overflow: 'auto', minHeight: 0, flex: 1 },
  footer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '12px',
    padding: '8px 20px',
    fontSize: '12px',
    color: tokens.colorNeutralForeground2,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  empty: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '16px',
    padding: '60px 24px',
    minHeight: '260px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground2,
  },
  icon: { fontSize: '48px', color: tokens.colorNeutralForeground3 },
  filter: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '12px 20px',
    backgroundColor: tokens.colorNeutralBackground2,
    flexWrap: 'wrap',
  },
  filterInput: { minWidth: '160px' },
  message: { borderRadius: 0 },
  messageBody: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '8px',
  },
  skeleton: { padding: '10px 24px', display: 'flex', flexDirection: 'column', gap: '20px' },
  live: {
    position: 'absolute',
    width: '1px',
    height: '1px',
    padding: 0,
    margin: '-1px',
    overflow: 'hidden',
    clip: 'rect(0,0,0,0)',
    whiteSpace: 'nowrap',
    border: 0,
  },
  drop: {
    position: 'absolute',
    inset: 0,
    zIndex: 25,
    backgroundColor: tokens.colorBrandBackground2,
    border: `2px dashed ${tokens.colorBrandStroke1}`,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '12px',
    pointerEvents: 'none',
  },
  hidden: { display: 'none' },
  columns: { display: 'flex', flexDirection: 'column', gap: '8px', paddingTop: '12px' },
  errorDetails: { fontSize: '12px', padding: '8px 0' },
});
function Canvas({ host }: { host: HostContext }) {
  const s = useStyles(),
    t = React.useMemo(() => createI18n(host.getString), [host.getString]),
    repo = React.useMemo(
      () => new DocumentRepository(new DataverseClient(host.clientUrl, host.pageSize || 50)),
      [host.clientUrl, host.pageSize],
    );
  const [state, dispatch] = React.useReducer(reduce, {
    config: null,
    rows: [],
    loading: !!host.recordId,
    total: 0,
    error: null,
    search: '',
    preset: 'all',
    typeFilter: '',
    orderBy: 'modifiedon desc',
    view: host.defaultView || 'list',
    details: host.showDetailsPane === true,
    filters: false,
    expanded: false,
    compact: false,
    columns: [],
    notice: '',
    dialog: null,
    dialogRows: [],
    migration: null,
    dragging: false,
  });
  const set = (patch: Partial<State>) => dispatch({ type: 'set', patch }),
    [refresh, setRefresh] = React.useState(0),
    [debounced, setDebounced] = React.useState(''),
    [actionBusy, setActionBusy] = React.useState(false),
    [offline, setOffline] = React.useState(!navigator.onLine),
    root = React.useRef<HTMLElement>(null),
    uploadInput = React.useRef<HTMLInputElement>(null),
    urls = React.useRef(new Set<string>()),
    nextAbort = React.useRef<AbortController>();
  const query: ListQuery = {
      recordId: host.recordId,
      entityName: host.entityName,
      search: debounced,
      orderBy: state.orderBy,
      preset: state.preset,
      typeFilter: state.typeFilter,
      pageSize: host.pageSize || 50,
    },
    queryKey = JSON.stringify(query),
    queryRef = React.useRef(queryKey);
  queryRef.current = queryKey;
  const selection = useSelection(state.rows.map((row) => row.id)),
    picked = state.rows.filter((row) => selection.selected.has(row.id));
  const queue = React.useMemo(
    () =>
      state.config?.entity.enabled && host.recordId
        ? new UploadQueue(repo, state.config, host.recordId, host.entityName, (row) =>
            dispatch({ type: 'patch', rows: [row] }),
          )
        : null,
    [repo, state.config, host.recordId, host.entityName],
  );
  React.useEffect(() => () => queue?.dispose(), [queue]);
  React.useEffect(
    () => () => {
      urls.current.forEach((url) => URL.revokeObjectURL(url));
    },
    [],
  );
  React.useEffect(() => {
    const timer = setTimeout(() => setDebounced(state.search.trim()), 300);
    return () => clearTimeout(timer);
  }, [state.search]);
  React.useEffect(() => {
    const online = () => {
      setOffline(!navigator.onLine);
      if (navigator.onLine) setRefresh((value) => value + 1);
    };
    window.addEventListener('online', online);
    window.addEventListener('offline', online);
    return () => {
      window.removeEventListener('online', online);
      window.removeEventListener('offline', online);
    };
  }, []);
  React.useEffect(() => {
    if (!host.recordId) return;
    const controller = new AbortController();
    const load = () => {
      void repo
        .config(host.entityName, host.recordId, controller.signal)
        .then((config) => {
          if (controller.signal.aborted) return;
          set({
            config,
            columns: config.entity.visibleColumns,
            migration: config.migration,
            error: null,
            ...(!config.entity.enabled ? { loading: false } : {}),
          });
        })
        .catch((error) => {
          if (!aborted(error))
            set({
              error: error instanceof ApiError ? error : new ApiError('unknown', 0),
              loading: false,
            });
        });
    };
    load();
    window.addEventListener('dms-config-changed', load);
    return () => {
      controller.abort();
      window.removeEventListener('dms-config-changed', load);
    };
  }, [repo, host.entityName, host.recordId]);
  React.useEffect(() => {
    if (!state.config?.entity.enabled || !host.recordId || offline) return;
    const controller = new AbortController();
    nextAbort.current?.abort();
    set({ loading: true });
    void repo
      .list(query, undefined, controller.signal)
      .then((page) => {
        if (!controller.signal.aborted)
          dispatch({
            type: 'rows',
            rows: page.rows,
            total: page.totalCount,
            nextCursor: page.nextCursor,
          });
      })
      .catch((error) => {
        if (!aborted(error))
          set({
            error: error instanceof ApiError ? error : new ApiError('unknown', 0),
            loading: false,
          });
      });
    return () => {
      controller.abort();
      nextAbort.current?.abort();
    };
  }, [repo, queryKey, refresh, state.config?.entity.enabled, offline]);
  React.useEffect(() => {
    if (
      !state.config ||
      !host.recordId ||
      !['NotStarted', 'Running'].includes(state.config.migration.state)
    )
      return;
    const controller = new AbortController(),
      started = Date.now();
    let timer: ReturnType<typeof setTimeout> | undefined,
      last = -1;
    const poll = async (ensure = false) => {
      try {
        const migration = await repo.migration(
          host.entityName,
          host.recordId,
          ensure,
          controller.signal,
        );
        if (controller.signal.aborted) return;
        set({ migration });
        if (migration.processed !== last) {
          last = migration.processed;
          setRefresh((value) => value + 1);
        }
        if (migration.state === 'Running' && Date.now() - started < 1800000)
          timer = setTimeout(() => void poll(), 10000);
      } catch (error) {
        if (!aborted(error))
          set({ error: error instanceof ApiError ? error : new ApiError('unknown', 0) });
      }
    };
    void poll(state.config.migration.state === 'NotStarted');
    return () => {
      controller.abort();
      if (timer) clearTimeout(timer);
    };
  }, [repo, state.config, host.entityName, host.recordId]);
  const active = resolveProvider(state.config?.activeProvider || 'Note', repo.client),
    selectedProviders = picked.map((row) => resolveProvider(row.provider, repo.client));
  const available = {
    open: picked.length === 1 && usable(picked[0], selectedProviders[0]),
    preview: picked.length === 1 && usable(picked[0], selectedProviders[0]),
    download: picked.some(
      (row, index) =>
        usable(row, selectedProviders[index]) && selectedProviders[index].capabilities.canDownload,
    ),
    copy: picked.length > 0 && selectedProviders.every((provider) => provider.capabilities.canOpen),
    edit: picked.length > 0 && selectedProviders.every((provider) => provider.capabilities.canOpen),
    delete:
      picked.length > 0 && selectedProviders.every((provider) => provider.capabilities.canDelete),
  };
  async function nextPage() {
    if (!state.nextCursor || state.loading) return;
    const controller = new AbortController();
    nextAbort.current = controller;
    set({ loading: true });
    try {
      const page = await repo.list(query, state.nextCursor, controller.signal);
      if (queryRef.current === queryKey && !controller.signal.aborted)
        dispatch({
          type: 'rows',
          rows: page.rows,
          total: page.totalCount,
          nextCursor: page.nextCursor,
          append: true,
        });
    } catch (error) {
      if (!aborted(error))
        set({
          loading: false,
          error: error instanceof ApiError ? error : new ApiError('unknown', 0),
        });
    }
  }
  function retainUrl(blob: Blob): string {
    const url = URL.createObjectURL(blob);
    urls.current.add(url);
    setTimeout(() => {
      URL.revokeObjectURL(url);
      urls.current.delete(url);
    }, 60000);
    return url;
  }
  function openReference(row: DocumentRow) {
    try {
      const url = new URL(row.storageRef);
      if (url.protocol !== 'https:') throw new Error();
      window.open(url.href, '_blank', 'noopener,noreferrer');
    } catch {
      set({ notice: t('link.invalid') });
    }
  }
  async function download(rows: DocumentRow[]) {
    setActionBusy(true);
    let count = 0,
      skipped = 0,
      failed = 0;
    const references = rows.filter((row) => {
      const provider = resolveProvider(row.provider, repo.client);
      return (
        usable(row, provider) &&
        provider.capabilities.canDownload &&
        !provider.capabilities.ownsBinary
      );
    });
    if (references.length <= 3)
      references.forEach((row) => {
        openReference(row);
        count++;
      });
    else skipped += references.length;
    for (const row of rows) {
      const provider = resolveProvider(row.provider, repo.client);
      if (!provider.capabilities.ownsBinary) continue;
      if (!usable(row, provider)) {
        skipped++;
        continue;
      }
      try {
        const content = await provider.getContent(row, (percent) =>
          set({
            notice: t('download.progress', { count: count + 1, total: rows.length, percent }),
          }),
        );
        if (content.kind === 'blob') {
          const anchor = document.createElement('a');
          anchor.href = retainUrl(content.blob);
          anchor.download = row.name;
          anchor.style.display = 'none';
          document.body.appendChild(anchor);
          anchor.click();
          anchor.remove();
          count++;
        }
      } catch {
        failed++;
      }
    }
    set({ notice: t('download.result', { count, skipped, failed }) });
    setActionBusy(false);
  }
  async function copy(rows: DocumentRow[]) {
    try {
      const links = rows.map((row) => {
        const provider = resolveProvider(row.provider, repo.client);
        if (!provider.capabilities.ownsBinary && provider.capabilities.canOpen)
          return row.storageRef;
        const url = new URL('/main.aspx', host.clientUrl);
        url.searchParams.set('pagetype', 'entityrecord');
        url.searchParams.set('etn', 'dms_document');
        url.searchParams.set('id', row.id);
        return url.toString();
      });
      await navigator.clipboard.writeText(links.join('\n'));
      set({ notice: t('copy.result', { count: rows.length }) });
    } catch {
      set({ notice: t('copy.failed') });
    }
  }
  function action(kind: Action, row?: DocumentRow) {
    const target = row ? (selection.selected.has(row.id) ? picked : [row]) : picked;
    if (!target.length) return;
    if (row) selection.context(row.id);
    if (kind === 'copy') {
      void copy(target);
      return;
    }
    if (kind === 'download') {
      void download(target);
      return;
    }
    if (kind === 'preview' || kind === 'open') {
      const current = row || target[0],
        provider = resolveProvider(current.provider, repo.client);
      if (!usable(current, provider)) {
        set({ notice: t('provider.unavailable') });
        return;
      }
      if (kind === 'open' && !provider.capabilities.ownsBinary) {
        openReference(current);
        return;
      }
      set({ dialog: 'preview', dialogRows: [current] });
      return;
    }
    set({ dialog: kind, dialogRows: target });
  }
  const upload = (files: File[]) => {
    if (!queue || !active.capabilities.canUpload) return;
    const errors = queue.add(files);
    if (errors.length) {
      const max = state.config
        ? Math.min(
            state.config.entity.maxFileSizeMb,
            state.config.org.maxUploadFileSizeBytes / 1048576,
          )
        : 0;
      set({ notice: t(errors[0], { limit: max }) });
    }
    if (files.length) set({ search: '', preset: 'all', typeFilter: '' });
  };
  const focus = (event: React.KeyboardEvent<HTMLElement>) => {
    const target = event.target as Element;
    if (target.closest('input,textarea,[role="combobox"],[contenteditable="true"]')) return;
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'a') {
      event.preventDefault();
      selection.setSelected(new Set(selection.ids));
      return;
    }
    if (event.key === 'Escape') {
      if (state.expanded) set({ expanded: false });
      else selection.clear();
      return;
    }
    const node = target.closest('[data-document-id]') as HTMLElement | null,
      id = node?.dataset.documentId;
    if (!id) return;
    if (event.key === 'Enter') {
      event.preventDefault();
      action(
        'preview',
        state.rows.find((row) => row.id === id),
      );
      return;
    }
    if (event.key === 'F2') {
      event.preventDefault();
      action(
        'edit',
        state.rows.find((row) => row.id === id),
      );
      return;
    }
    if (event.key === 'Delete') {
      event.preventDefault();
      action(
        'delete',
        state.rows.find((row) => row.id === id),
      );
      return;
    }
    if (state.view === 'tiles' && event.key === ' ') {
      event.preventDefault();
      selection.pick(id, event, true);
      return;
    }
    if (!['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key))
      return;
    const nodes = Array.from(
        root.current?.querySelectorAll<HTMLElement>('[data-document-id]') || [],
      ),
      index = nodes.findIndex((item) => item.dataset.documentId === id);
    let columns = 1;
    if (state.view === 'tiles' && nodes.length > 1) {
      const top = nodes[0].getBoundingClientRect().top;
      columns =
        nodes.filter((item) => Math.abs(item.getBoundingClientRect().top - top) < 2).length || 1;
    }
    let next = index;
    if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = nodes.length - 1;
    else if (event.key === 'ArrowDown') next += columns;
    else if (event.key === 'ArrowUp') next -= columns;
    else if (state.view === 'tiles') next += event.key === 'ArrowRight' ? 1 : -1;
    else return;
    next = Math.min(nodes.length - 1, Math.max(0, next));
    event.preventDefault();
    event.stopPropagation();
    nodes[next]?.focus();
    const nextId = nodes[next]?.dataset.documentId;
    if (nextId && !(event.ctrlKey || event.metaKey) && next !== index)
      selection.pick(nextId, event);
  };
  const props: ViewProps = {
    rows: state.rows,
    selection,
    repo,
    t,
    columns: state.columns,
    orderBy: state.orderBy,
    onSort: (orderBy) => set({ orderBy }),
    onHide: (column) => set({ columns: state.columns.filter((value) => value !== column) }),
    onMove: (column, direction) => {
      const columns = [...state.columns],
        index = columns.indexOf(column),
        target = index + direction;
      if (target >= 0 && target < columns.length) {
        [columns[index], columns[target]] = [columns[target], columns[index]];
        set({ columns });
      }
    },
    onColumns: () => set({ dialog: 'columns' }),
    onFilter: () => set({ filters: !state.filters }),
    onAction: action,
    onDetails: (row) => {
      selection.setSelected(new Set([row.id]));
      set({ details: true });
    },
    compact: state.compact,
    details: state.details,
  };
  const emptyKey = !host.recordId
    ? 'unsaved.title'
    : state.config && !state.config.entity.enabled
      ? 'notConfigured'
      : state.search || state.preset !== 'all' || state.typeFilter
        ? 'noResults'
        : 'empty.title';
  const ready = !!host.recordId && !!state.config?.entity.enabled;
  return (
    <>
      <section
        ref={root}
        className={mergeClasses(s.root, state.expanded && s.expanded)}
        aria-label={t('documents')}
        onKeyDown={focus}
        style={
          host.allocatedHeight > 0 && !state.expanded
            ? { height: host.allocatedHeight }
            : state.expanded
              ? undefined
              : { maxHeight: 'calc(100vh - 200px)' }
        }
        onDragOver={(event) => {
          if (
            ready &&
            active.capabilities.canUpload &&
            host.enableDragDrop !== false &&
            event.dataTransfer.types.includes('Files')
          ) {
            event.preventDefault();
            set({ dragging: true });
          }
        }}
        onDragLeave={(event) => {
          if (!event.currentTarget.contains(event.relatedTarget as Node)) set({ dragging: false });
        }}
        onDrop={(event) => {
          event.preventDefault();
          set({ dragging: false });
          if (!active.capabilities.canUpload) return;
          const folder = Array.from(event.dataTransfer.items).some(
            (item) => item.webkitGetAsEntry?.()?.isDirectory,
          );
          if (folder) {
            set({ notice: t('upload.folders') });
            return;
          }
          upload(Array.from(event.dataTransfer.files));
        }}
      >
        {host.showTitle && (
          <div className={s.heading}>
            <h1 className={s.title}>{t('documents')}</h1>
            {ready && <span className={s.count}>{state.total}</span>}
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton
                  className={s.chip}
                  appearance="subtle"
                  icon={<LibraryRegular />}
                  aria-label={t('views.label')}
                >
                  <ChevronDownRegular />
                </ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem onClick={() => set({ view: 'list' })}>{t('view.list')}</MenuItem>
                  <MenuItem onClick={() => set({ view: 'tiles' })}>{t('view.tiles')}</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>
        )}
        {ready && (
          <DocumentCommandBar
            t={t}
            active={active.capabilities}
            selectionCount={selection.selected.size}
            available={available}
            busy={actionBusy || offline}
            onAction={action}
            onUpload={() => uploadInput.current?.click()}
            onAddLink={() => set({ dialog: 'link' })}
            onRefresh={() => {
              setRefresh((value) => value + 1);
              set({ notice: '' });
            }}
            onClear={selection.clear}
            search={state.search}
            onSearch={(search) => set({ search })}
            preset={state.preset}
            onPreset={(preset) => set({ preset })}
            view={state.view}
            onView={(view) => set({ view })}
            filters={state.filters}
            onFilters={() => set({ filters: !state.filters })}
            details={state.details}
            onDetails={() => set({ details: !state.details })}
            expanded={state.expanded}
            onExpand={() => set({ expanded: !state.expanded })}
            onDensity={() => set({ compact: !state.compact })}
            onColumns={() => set({ dialog: 'columns' })}
            onShortcuts={() => set({ dialog: 'shortcuts' })}
          />
        )}
        {offline && (
          <MessageBar className={s.message} intent="warning">
            <MessageBarBody>{t('offline')}</MessageBarBody>
          </MessageBar>
        )}
        {state.error && (
          <MessageBar className={s.message} intent="error">
            <MessageBarBody>
              <div className={s.messageBody}>
                <span>{t('error.' + state.error.kind)}</span>
                <Toolbar>
                  <ToolbarButton
                    icon={<ArrowSyncRegular />}
                    onClick={() => {
                      if (!state.config) window.dispatchEvent(new Event('dms-config-changed'));
                      setRefresh((value) => value + 1);
                    }}
                  >
                    {t('action.retry')}
                  </ToolbarButton>
                </Toolbar>
              </div>
              {state.error.correlationId && (
                <div className={s.errorDetails}>
                  {t('error.correlation', { id: state.error.correlationId })}
                </div>
              )}
            </MessageBarBody>
          </MessageBar>
        )}
        {state.notice && (
          <MessageBar className={s.message} intent="info">
            <MessageBarBody>
              <div className={s.messageBody}>
                {state.notice}
                <Toolbar>
                  <ToolbarButton
                    appearance="subtle"
                    icon={<DismissRegular />}
                    aria-label={t('message.dismiss')}
                    onClick={() => set({ notice: '' })}
                  />
                </Toolbar>
              </div>
            </MessageBarBody>
          </MessageBar>
        )}
        {state.migration?.state === 'Running' && (
          <MessageBar className={s.message} intent="info">
            <MessageBarBody>
              {t('migration.running', {
                processed: state.migration.processed,
                total: state.migration.total,
              })}
            </MessageBarBody>
          </MessageBar>
        )}
        {state.migration?.state === 'Failed' && (
          <MessageBar intent="warning">
            <MessageBarBody>{t('migration.failed')}</MessageBarBody>
          </MessageBar>
        )}
        {queue && <UploadTray queue={queue} t={t} />}
        {ready && state.filters && (
          <div className={s.filter}>
            <Text>{t('column.dms_documenttype')}</Text>
            <Dropdown
              className={s.filterInput}
              aria-label={t('column.dms_documenttype')}
              value={state.typeFilter ? t('type.' + state.typeFilter) : t('filters.allTypes')}
              selectedOptions={[state.typeFilter]}
              onOptionSelect={(_, data) => set({ typeFilter: data.optionValue || '' })}
            >
              <Option value="">{t('filters.allTypes')}</Option>
              {[0, 1, 2, 3, 4, 5].map((value) => (
                <Option key={value} value={String(100000000 + value)}>
                  {t('type.' + (100000000 + value))}
                </Option>
              ))}
            </Dropdown>
            <Toolbar>
              <ToolbarButton onClick={() => set({ search: '', typeFilter: '', preset: 'all' })}>
                {t('filters.clear')}
              </ToolbarButton>
            </Toolbar>
          </div>
        )}
        <div className={s.workspace}>
          <div className={s.main}>
            <div className={s.scroll}>
              {state.loading && !state.rows.length ? (
                <Skeleton className={s.skeleton} aria-label={t('loading')}>
                  {Array.from({ length: 8 }, (_, index) => (
                    <SkeletonItem key={index} size={24} />
                  ))}
                </Skeleton>
              ) : ready && state.rows.length ? (
                state.view === 'list' ? (
                  <DocumentList {...props} />
                ) : (
                  <DocumentTiles {...props} />
                )
              ) : (
                !state.error && (
                  <div className={s.empty}>
                    <DocumentRegular className={s.icon} />
                    <Text weight="semibold">{t(emptyKey)}</Text>
                    {emptyKey === 'noResults' && (
                      <Toolbar>
                        <ToolbarButton
                          onClick={() => set({ search: '', preset: 'all', typeFilter: '' })}
                        >
                          {t('filters.clear')}
                        </ToolbarButton>
                      </Toolbar>
                    )}
                    {emptyKey === 'empty.title' && <Text>{t('empty.description')}</Text>}
                  </div>
                )
              )}
            </div>
            {ready && (
              <div className={s.footer}>
                <span>{t('items.count', { loaded: state.rows.length, total: state.total })}</span>
                {state.nextCursor ? (
                  <Toolbar>
                    <ToolbarButton disabled={state.loading} onClick={() => void nextPage()}>
                      {t(state.loading ? 'loading' : 'items.loadMore')}
                    </ToolbarButton>
                  </Toolbar>
                ) : (
                  <span>{t('items.loaded')}</span>
                )}
              </div>
            )}
          </div>
          {ready && state.details && (
            <DetailsPane
              rows={picked}
              repo={repo}
              t={t}
              onClose={() => set({ details: false })}
              onAction={action}
            />
          )}
        </div>
        <input
          className={s.hidden}
          type="file"
          multiple
          ref={uploadInput}
          aria-label={t('action.upload')}
          onChange={(event) => {
            upload(Array.from(event.target.files || []));
            event.target.value = '';
          }}
        />
        {state.dragging && (
          <div className={s.drop}>
            <ArrowUploadRegular fontSize={40} />
            <Text weight="semibold">{t('upload.drop')}</Text>
          </div>
        )}
        <div className={s.live} aria-live="polite">
          {t('selection.count', { count: selection.selected.size })}
        </div>
      </section>
      {state.dialog === 'preview' && state.dialogRows[0] && (
        <PreviewDialog
          row={state.dialogRows[0]}
          repo={repo}
          t={t}
          onClose={() => set({ dialog: null })}
          onDownload={() => void download(state.dialogRows)}
          onOpen={() => openReference(state.dialogRows[0])}
        />
      )}
      {state.dialog === 'edit' && state.config && (
        <MetadataDrawer
          rows={state.dialogRows}
          repo={repo}
          config={state.config}
          t={t}
          onClose={() => set({ dialog: null })}
          onSaved={(rows) => {
            dispatch({ type: 'patch', rows });
            set({ notice: t('edit.saved', { count: rows.length }) });
          }}
        />
      )}
      {state.dialog === 'delete' && (
        <DeleteDialog
          rows={state.dialogRows}
          repo={repo}
          t={t}
          onClose={() => set({ dialog: null })}
          onDeleted={(rows) => {
            dispatch({ type: 'remove', ids: rows.map((row) => row.id) });
            set({ notice: t('delete.result', { count: rows.length }) });
          }}
        />
      )}
      {state.dialog === 'link' && state.config && (
        <AddLinkDialog
          repo={repo}
          config={state.config}
          recordId={host.recordId}
          entityName={host.entityName}
          t={t}
          onClose={() => set({ dialog: null })}
          onAdded={(row) => {
            dispatch({ type: 'patch', rows: [row] });
            set({ notice: t('link.added') });
          }}
        />
      )}
      {['columns', 'shortcuts'].includes(state.dialog || '') && (
        <Dialog open onOpenChange={(_, data) => !data.open && set({ dialog: null })}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>
                {t(state.dialog === 'columns' ? 'columns.edit' : 'shortcuts.title')}
              </DialogTitle>
              <DialogContent>
                {state.dialog === 'columns' ? (
                  <div className={s.columns}>
                    {state.config?.entity.visibleColumns.map((column) => (
                      <Checkbox
                        key={column}
                        label={t('column.' + column)}
                        checked={state.columns.includes(column)}
                        disabled={column === 'dms_name'}
                        onChange={(_, data) =>
                          set({
                            columns: data.checked
                              ? [...state.columns, column]
                              : state.columns.filter((value) => value !== column),
                          })
                        }
                      />
                    ))}
                  </div>
                ) : (
                  <Text>{t('shortcuts.body')}</Text>
                )}
              </DialogContent>
              <DialogActions>
                <Toolbar>
                  <ToolbarButton appearance="primary" onClick={() => set({ dialog: null })}>
                    {t('action.done')}
                  </ToolbarButton>
                </Toolbar>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </>
  );
}
class ControlBoundary extends React.Component<
  { host: HostContext; children: React.ReactNode },
  { failed: boolean }
> {
  state = { failed: false };
  static getDerivedStateFromError() {
    return { failed: true };
  }
  render() {
    return this.state.failed ? (
      <MessageBar intent="error">
        <MessageBarBody>
          {this.props.host.getString('error.control')}
          <Toolbar>
            <ToolbarButton onClick={() => this.setState({ failed: false })}>
              {this.props.host.getString('action.reload')}
            </ToolbarButton>
          </Toolbar>
        </MessageBarBody>
      </MessageBar>
    ) : (
      this.props.children
    );
  }
}
export function DmsGridHost({ host }: { host: HostContext }) {
  return (
    <FluentProvider theme={host.theme || webLightTheme}>
      <ControlBoundary host={host}>
        <Canvas key={host.recordId + '|' + host.entityName} host={host} />
      </ControlBoundary>
    </FluentProvider>
  );
}
