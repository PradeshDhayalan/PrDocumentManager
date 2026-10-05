import {
  availableColumns,
  configuredColumns,
  defaultColumns,
  moveColumn,
} from '../services/columns';
import { CustomAction, parseCustomActions, invokeCustomAction } from '../services/customActions';
import { microsoftFontFamily } from '../services/fonts';
import { useDragSelection } from '../components/useDragSelection';
import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Text,
  Badge,
  Checkbox,
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
  Field,
  Input,
  ProgressBar,
  InlineDrawer,
  OverlayDrawer,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerFooter,
  Spinner,
  Dropdown,
  Option,
} from '@fluentui/react-components';
import {
  LibraryRegular,
  ChevronDownRegular,
  TextBulletListRegular,
  GridRegular,
  DismissRegular,
  OpenRegular,
  ArrowDownloadRegular,
  DocumentRegular,
  EditRegular,
  ArrowUploadRegular,
  WarningRegular,
  ArrowUpRegular,
  ArrowDownRegular,
} from '@fluentui/react-icons';
import { HostContext } from '../types/HostContext';
import { createI18n } from '../services/i18n';
import {
  DocumentClient,
  DocumentRow,
  ClientConfig,
  Attribute,
  Entity,
  Query,
  label,
  safeExternalUrl,
} from '../services/documents';
import DocumentGrid from '../components/DocumentGrid';
import DocumentCommandBar from '../components/DocumentCommandBar';
import DocumentThumbnail from '../components/DocumentThumbnail';
import { FileIcon } from '../components/FileIcon';
import { DocumentPerson } from '../components/DocumentPerson';
import { MetadataFields } from '../components/MetadataFields';
const useStyles = makeStyles({
  secondary: { color: tokens.colorNeutralForeground2, fontSize: '12px' },
  library: {
    fontFamily: tokens.fontFamilyBase,
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusNone,
    overflow: 'hidden',
    minHeight: '730px',
  },
  expanded: {
    position: 'fixed',
    inset: '12px',
    zIndex: 10,
    boxShadow: tokens.shadow64,
    overflowY: 'auto',
  },
  heading: { display: 'flex', alignItems: 'center', gap: '10px', padding: '24px 24px 20px' },
  title: { fontSize: '24px', fontWeight: tokens.fontWeightSemibold, lineHeight: '32px' },
  count: { fontSize: '14px', color: tokens.colorNeutralForeground3 },
  chip: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    minWidth: '40px',
  },
  workspace: { display: 'flex', minHeight: '560px' },
  main: { flex: 1, minWidth: 0 },
  gridScroll: { overflowX: 'auto', width: '100%' },
  pane: {
    width: '300px',
    flexShrink: 0,
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
    padding: '18px 22px',
    boxSizing: 'border-box',
    '@media(max-width:1000px)': {
      position: 'fixed',
      top: 0,
      right: 0,
      bottom: 0,
      width: '320px',
      zIndex: 20,
      backgroundColor: tokens.colorNeutralBackground1,
      boxShadow: tokens.shadow64,
      overflowY: 'auto',
    },
  },
  paneHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '24px',
  },
  paneThumbnail: {
    height: '172px',
    marginBottom: '16px',
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
  },
  paneHero: { paddingBottom: '22px', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  paneName: {
    fontSize: '17px',
    fontWeight: tokens.fontWeightSemibold,
    margin: '14px 0 8px',
    overflowWrap: 'anywhere',
  },
  paneActions: { display: 'flex', gap: '6px', marginTop: '16px' },
  paneSection: { padding: '20px 0', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: '14px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '16px',
  },
  props: { display: 'grid', gridTemplateColumns: '100px 1fr', gap: '14px 10px', fontSize: '12px' },
  hint: {
    padding: '12px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    fontSize: '12px',
    lineHeight: '18px',
    marginTop: '16px',
  },
  footer: {
    height: '42px',
    padding: '0 22px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    fontSize: '12px',
    color: tokens.colorNeutralForeground2,
  },
  tileStrip: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '10px 16px',
  },
  tiles: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill,minmax(208px,1fr))',
    gap: '24px',
    padding: '12px 24px 24px',
    '@media(max-width:480px)': {
      gridTemplateColumns: 'repeat(2,minmax(0,1fr))',
      gap: '16px',
      padding: '12px 16px 20px',
    },
  },
  tile: {
    padding: 0,
    minWidth: 0,
    overflow: 'hidden',
    boxShadow: tokens.shadow2,
    borderRadius: tokens.borderRadiusMedium,
    transitionProperty: 'box-shadow',
    transitionDuration: tokens.durationNormal,
    transitionTimingFunction: tokens.curveEasyEase,
    '&:hover': { boxShadow: tokens.shadow16 },
    '&:focus-within': { boxShadow: tokens.shadow16 },
    '@media(prefers-reduced-motion:reduce)': { transitionDuration: '0ms' },
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: 'pointer',
    position: 'relative',
    '&:hover .selection': { opacity: 1 },
  },
  tileSelected: {
    boxShadow: tokens.shadow8,
    outline: `2px solid ${tokens.colorBrandStroke1}`,
    outlineOffset: '-2px',
    backgroundColor: tokens.colorBrandBackground2,
  },
  tilePreview: {
    height: '196px',
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'grid',
    placeItems: 'center',
  },
  tileBody: { padding: '14px 16px', display: 'flex', flexDirection: 'column', gap: '8px' },
  tileName: {
    fontWeight: tokens.fontWeightRegular,
    fontSize: '14px',
    height: '36px',
    minHeight: '36px',
    maxHeight: '36px',
    lineHeight: '18px',
    overflow: 'hidden',
    display: '-webkit-box',
    WebkitBoxOrient: 'vertical',
    WebkitLineClamp: 2,
    overflowWrap: 'anywhere',
  },
  tileCheck: { position: 'absolute', top: '6px', left: '6px' },
  tileMore: { position: 'absolute', top: '6px', right: '6px' },
  filter: {
    padding: '12px 20px',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    backgroundColor: tokens.colorNeutralBackground2,
    flexWrap: 'wrap',
  },
  empty: { textAlign: 'center', padding: '90px 24px', color: tokens.colorNeutralForeground2 },
  notice: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 20px',
    backgroundColor: tokens.colorBrandBackground2,
    fontSize: '12px',
  },
  editDrawer: { width: '420px', maxWidth: '100vw' },
  dialogFields: { display: 'flex', flexDirection: 'column', gap: '16px', marginTop: '14px' },
  preview: {
    height: '300px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexDirection: 'column',
    gap: '20px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
});
const initialQuery: Query = {
  search: '',
  type: '',
  status: '',
  preset: 'All documents',
  sort: 'modifiedon',
  descending: true,
};
function Library({ host }: { host: HostContext }) {
  const s = useStyles(),
    t = createI18n(host.getString);
  const client = React.useMemo(
    () => new DocumentClient(host),
    [host.clientUrl, host.recordId, host.entityName, host.pageSize, host.userId],
  );
  const [config, setConfig] = React.useState<ClientConfig>(),
    [attributes, setAttributes] = React.useState<Attribute[]>([]);
  const [rows, setRows] = React.useState<DocumentRow[]>([]),
    [count, setCount] = React.useState(0),
    [next, setNext] = React.useState<string>();
  const [query, setQuery] = React.useState<Query>(initialQuery),
    [search, setSearch] = React.useState(''),
    [revision, setRevision] = React.useState(0);
  const [selected, setSelected] = React.useState<string[]>([]),
    [details, setDetails] = React.useState(false),
    [view, setView] = React.useState<string>(host.defaultView || 'List');
  const [filters, setFilters] = React.useState(false),
    [expanded, setExpanded] = React.useState(false),
    [compact, setCompact] = React.useState(false),
    [columns, setColumns] = React.useState<string[]>(defaultColumns);
  const [columnWidths, setColumnWidths] = React.useState<Record<string, number>>({});
  const [loading, setLoading] = React.useState(false),
    [busy, setBusy] = React.useState(false),
    [notice, setNotice] = React.useState(''),
    [error, setError] = React.useState('');
  const [dialog, setDialog] = React.useState(''),
    [editValues, setEditValues] = React.useState<Entity>({}),
    [dirty, setDirty] = React.useState<Entity>({});
  const [linkUrl, setLinkUrl] = React.useState(''),
    [linkName, setLinkName] = React.useState('');
  const [files, setFiles] = React.useState<File[]>([]),
    [progress, setProgress] = React.useState(0),
    [uploadName, setUploadName] = React.useState(''),
    [uploadResults, setUploadResults] = React.useState<string[]>([]);
  const [previewUrl, setPreviewUrl] = React.useState(''),
    [previewText, setPreviewText] = React.useState('');
  const actionController = React.useRef<AbortController>();
  const pageController = React.useRef<AbortController>();
  const listGeneration = React.useRef(0);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const region = React.useRef<HTMLElement>(null),
    expandButtonFocus = React.useRef<HTMLElement | null>(null);
  const [pendingAction, setPendingAction] = React.useState<{
    action: CustomAction;
    rows: DocumentRow[];
  }>();
  const customConfiguration = React.useMemo(() => {
    try {
      return { actions: parseCustomActions(host.customActionsJson), error: '' };
    } catch (error) {
      return {
        actions: [],
        error: error instanceof Error ? error.message : 'Invalid custom action configuration.',
      };
    }
  }, [host.customActionsJson]);
  const dragSelection = useDragSelection(selected, setSelected);
  const picked = rows.filter((row) => selected.includes(row.id)),
    current = picked[0];
  const editAttributes = attributes.filter(
    (attribute) =>
      attribute.IsValidForUpdate && config?.entity.editableColumns.includes(attribute.LogicalName),
  );
  const refresh = () => {
    setRevision((value) => value + 1);
    setError('');
  };
  React.useEffect(() => {
    const timer = setTimeout(
      () => setQuery((value) => (value.search === search ? value : { ...value, search })),
      250,
    );
    return () => clearTimeout(timer);
  }, [search]);
  React.useEffect(() => {
    const controller = new AbortController();
    setConfig(undefined);
    setAttributes([]);
    setRows([]);
    setCount(0);
    setSelected([]);
    setDetails(false);
    setDialog('');
    setError('');
    setBusy(false);
    actionController.current?.abort();
    actionController.current = undefined;
    if (host.recordId)
      void Promise.all([client.config(controller.signal), client.attributes(controller.signal)])
        .then(([settings, metadata]) => {
          setConfig(settings);
          setAttributes(metadata);
          setColumns(configuredColumns(settings.entity.visibleColumns, metadata));
          setColumnWidths({});
        })
        .catch((e) => {
          if (!controller.signal.aborted)
            setError(e instanceof Error ? e.message : 'Could not load configuration.');
        });
    const reload = () => {
      void client
        .config(controller.signal)
        .then(setConfig)
        .catch((e) => {
          if (!controller.signal.aborted) setError(String(e));
        });
    };
    window.addEventListener('dms-config-changed', reload);
    return () => {
      controller.abort();
      actionController.current?.abort();
      actionController.current = undefined;
      window.removeEventListener('dms-config-changed', reload);
    };
  }, [client, host.recordId]);
  React.useEffect(() => {
    const controller = new AbortController();
    const generation = ++listGeneration.current;
    pageController.current?.abort();
    setNext(undefined);
    if (!host.recordId || !config?.entity.enabled) return;
    setLoading(true);
    void client
      .list(query, controller.signal)
      .then((result) => {
        if (generation === listGeneration.current) {
          setRows(result.rows);
          setSelected((value) => value.filter((id) => result.rows.some((row) => row.id === id)));
          setCount(result.count);
          setNext(result.next);
        }
      })
      .catch((e) => {
        if (!controller.signal.aborted)
          setError(e instanceof Error ? e.message : 'Could not load documents.');
      })
      .finally(() => {
        if (!controller.signal.aborted) setLoading(false);
      });
    return () => controller.abort();
  }, [client, host.recordId, config?.entity.enabled, query, revision]);
  React.useEffect(() => () => pageController.current?.abort(), []);
  React.useEffect(() => {
    if (!expanded) return;
    expandButtonFocus.current = document.activeElement as HTMLElement;
    const previous = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    region.current?.focus();
    return () => {
      document.body.style.overflow = previous;
      expandButtonFocus.current?.focus();
    };
  }, [expanded]);
  React.useEffect(() => {
    setPreviewUrl('');
    setPreviewText('');
    if (
      dialog !== 'Preview' ||
      !current ||
      current.source !== 'Note' ||
      current.raw.dms_uploadstate !== 100000001
    )
      return;
    const controller = new AbortController();
    let objectUrl = '';
    void client
      .download(current, controller.signal)
      .then(async (blob) => {
        if (controller.signal.aborted) return;
        if (current.fileType === 'Text') {
          const text = await blob.text();
          if (!controller.signal.aborted) setPreviewText(text.slice(0, 100000));
        } else if (['PDF', 'Photo', 'Video'].includes(current.fileType)) {
          objectUrl = URL.createObjectURL(blob);
          setPreviewUrl(objectUrl);
        }
      })
      .catch((e) => {
        if (!controller.signal.aborted)
          setError(e instanceof Error ? e.message : 'Preview unavailable.');
      });
    return () => {
      controller.abort();
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [dialog, current?.id, client]);
  const run = async (task: (signal: AbortSignal) => Promise<void>) => {
    if (busy) return;
    const controller = new AbortController();
    actionController.current = controller;
    setBusy(true);
    setError('');
    try {
      await task(controller.signal);
    } catch (e) {
      if (!controller.signal.aborted)
        setError(e instanceof Error ? e.message : 'The action failed.');
    } finally {
      if (actionController.current === controller) {
        setBusy(false);
        if (controller.signal.aborted) {
          setNotice('Upload cancelled.');
          refresh();
        }
      }
    }
  };
  const executeCustomAction = (action: CustomAction, documents: DocumentRow[]) =>
    void run(async (signal) => {
      const result = await invokeCustomAction(
        action,
        documents,
        { id: host.recordId, entityName: host.entityName },
        host.raiseCustomAction,
        signal,
      );
      if (signal.aborted) return;
      if (dialog === 'Custom action') setDialog('');
      setPendingAction(undefined);
      setNotice(result.message || `${action.label} completed.`);
      if (result.refresh ?? action.refreshOnSuccess) refresh();
    });
  const customAction = (action: CustomAction) => {
    if (busy) return;
    if (action.confirmMessage) {
      setPendingAction({ action, rows: [...picked] });
      setError('');
      setDialog('Custom action');
    } else executeCustomAction(action, picked);
  };
  const changeQuery = (patch: Partial<Query>) => {
    setQuery((value) => ({ ...value, ...patch }));
    setSelected([]);
  };
  const select = (
    id: string,
    event: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean },
    toggle = false,
  ) => {
    if (event.shiftKey && selected.length) {
      const a = rows.findIndex((row) => row.id === selected[selected.length - 1]),
        b = rows.findIndex((row) => row.id === id);
      if (a >= 0 && b >= 0) {
        setSelected(rows.slice(Math.min(a, b), Math.max(a, b) + 1).map((row) => row.id));
        return;
      }
    }
    setSelected(
      toggle || event.ctrlKey || event.metaKey
        ? selected.includes(id)
          ? selected.filter((value) => value !== id)
          : [...selected, id]
        : [id],
    );
  };
  const copy = async (documents: DocumentRow[]) => {
    const links = documents.map((row) =>
      row.source === 'SharePoint'
        ? safeExternalUrl(row)
        : `${host.clientUrl.replace(/\/$/, '')}/main.aspx?pagetype=entityrecord&etn=dms_document&id=${encodeURIComponent(row.id)}`,
    );
    await navigator.clipboard.writeText(links.join('\n'));
    setNotice(`${links.length === 1 ? 'Link' : 'Links'} copied.`);
  };
  const download = async (documents: DocumentRow[], signal: AbortSignal) => {
    for (const row of documents) {
      if (signal.aborted) return;
      if (row.source === 'SharePoint') {
        window.open(safeExternalUrl(row), '_blank', 'noopener,noreferrer');
        continue;
      }
      const blob = await client.download(row, signal);
      if (signal.aborted) return;
      const url = URL.createObjectURL(blob),
        anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = String(row.raw.dms_originalfilename || row.name);
      anchor.click();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    }
  };
  const action = (kind: string) => {
    if (kind === 'Copy link') {
      void run(async () => copy(picked));
      return;
    }
    if (kind === 'Download') {
      void run(async (signal) => {
        await download(picked, signal);
        if (!signal.aborted) setNotice('Download started.');
      });
      return;
    }
    if (kind === 'Open' && current) {
      if (current.source === 'SharePoint') {
        void run(async () => {
          window.open(safeExternalUrl(current), '_blank', 'noopener,noreferrer');
        });
      } else setDialog('Preview');
      return;
    }
    if (kind === 'Edit details') {
      setEditValues(picked.length === 1 ? { ...current.raw } : {});
      setDirty({});
    }
    setError('');
    setDialog(kind);
  };
  const loadMore = () => {
    if (!next || loading) return;
    const controller = new AbortController();
    pageController.current = controller;
    const generation = listGeneration.current;
    setLoading(true);
    void client
      .list(query, controller.signal, next)
      .then((result) => {
        if (generation !== listGeneration.current) return;
        setRows((value) => [
          ...value,
          ...result.rows.filter((row) => !value.some((existing) => existing.id === row.id)),
        ]);
        setNext(result.next);
        setCount(result.count);
      })
      .catch((e) => {
        if (!controller.signal.aborted) setError(String(e));
      })
      .finally(() => {
        if (!controller.signal.aborted) setLoading(false);
      });
  };
  const queueFiles = (incoming: File[]) => {
    if (busy) return;
    setFiles(incoming);
    setUploadResults([]);
    setProgress(0);
    setDialog('Upload');
    setError('');
  };
  const upload = () => {
    if (!config || !files.length) return;
    void run(async (signal) => {
      const results: string[] = [];
      let completed = 0;
      for (let index = 0; index < files.length; index++) {
        if (signal.aborted) break;
        const file = files[index];
        setUploadName(file.name);
        try {
          await client.upload(
            file,
            config,
            (value) => setProgress((index + value) / files.length),
            signal,
          );
          completed++;
          results.push(`${file.name}: uploaded`);
        } catch (e) {
          results.push(
            `${file.name}: ${signal.aborted ? 'cancelled' : e instanceof Error ? e.message : 'failed'}`,
          );
        }
        if (!signal.aborted) setUploadResults([...results]);
      }
      if (!signal.aborted) {
        setFiles([]);
        setUploadName('');
        setNotice(`${completed} file${completed === 1 ? '' : 's'} uploaded.`);
        refresh();
      }
    });
  };
  const close = () => {
    if (busy) return;
    setDialog('');
    setFiles([]);
    setError('');
  };
  const status = (row: DocumentRow) => (
    <Badge
      appearance="tint"
      color={
        row.status === 'Failed'
          ? 'danger'
          : row.status === 'Draft' || row.status === 'Pending'
            ? 'subtle'
            : row.status === 'Archived'
              ? 'informative'
              : 'success'
      }
    >
      {row.status}
    </Badge>
  );
  const canManage = Boolean(host.recordId && config?.entity.enabled);
  if (!host.recordId)
    return (
      <section className={s.library} aria-label={t('documents')}>
        <div className={s.heading}>
          <h1 className={s.title}>{t('documents')}</h1>
        </div>
        <div className={s.empty}>
          <DocumentRegular fontSize={48} />
          <p>{t('unsaved.title')}</p>
        </div>
      </section>
    );
  return (
    <section
      ref={region}
      tabIndex={-1}
      className={mergeClasses(s.library, expanded && s.expanded)}
      aria-label={t('documents')}
      onDragOver={(event) => {
        if (host.enableDragDrop !== false && canManage) {
          event.preventDefault();
          event.dataTransfer.dropEffect = busy ? 'none' : 'copy';
        }
      }}
      onDrop={(event) => {
        if (host.enableDragDrop !== false && canManage) {
          event.preventDefault();
          queueFiles(Array.from(event.dataTransfer.files));
        }
      }}
      onKeyDown={(event) => {
        const target = event.target as HTMLElement;
        if (target.closest('input,textarea,[role="combobox"],[contenteditable="true"]')) return;
        if (event.key === 'Escape') {
          if (dialog && !busy) close();
          else if (expanded) setExpanded(false);
          else setSelected([]);
        }
        if ((event.ctrlKey || event.metaKey) && event.key === 'a') {
          event.preventDefault();
          setSelected(rows.map((row) => row.id));
        }
      }}
    >
      {host.showTitle && (
        <div className={s.heading}>
          <h1 className={s.title} style={{ margin: 0 }}>
            {t('documents')}
          </h1>
          <span className={s.count} aria-label="Total documents">
            {count}
          </span>
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button
                className={s.chip}
                appearance="subtle"
                icon={<LibraryRegular />}
                aria-label="Switch view"
              >
                <ChevronDownRegular />
              </Button>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                {['List', 'Tiles'].map((value) => (
                  <MenuItem
                    key={value}
                    icon={value === 'List' ? <TextBulletListRegular /> : <GridRegular />}
                    onClick={() => setView(value)}
                  >
                    {value} view{view === value ? ' ✓' : ''}
                  </MenuItem>
                ))}
              </MenuList>
            </MenuPopover>
          </Menu>
        </div>
      )}
      {canManage && (
        <DocumentCommandBar
          customActions={customConfiguration.actions}
          onCustomAction={customAction}
          busy={busy}
          provider={config!.activeProvider}
          selectionCount={selected.length}
          onAction={action}
          onUpload={() => queueFiles([])}
          onAddLink={() => {
            setLinkUrl('');
            setLinkName('');
            setDialog('Add link');
            setError('');
          }}
          onRefresh={() => {
            refresh();
            setNotice('Documents refreshed.');
          }}
          onClear={() => setSelected([])}
          query={search}
          onSearch={(value) => {
            setSearch(value);
            setSelected([]);
          }}
          preset={query.preset}
          onPreset={(value) => changeQuery({ preset: value })}
          view={view}
          onView={setView}
          filters={filters}
          onFilters={() => setFilters(!filters)}
          details={details}
          onDetails={() => setDetails(!details)}
          expanded={expanded}
          onExpand={() => setExpanded(!expanded)}
          feedback={setNotice}
          onColumns={() => setDialog('Edit columns')}
          onShortcuts={() => setDialog('Keyboard shortcuts')}
          onDensity={() => setCompact(!compact)}
          compact={compact}
        />
      )}
      {customConfiguration.error && (
        <div role="alert" className={s.notice}>
          Custom buttons: {customConfiguration.error}
        </div>
      )}
      {error && (
        <div role="alert" className={s.notice}>
          <WarningRegular />
          {error}
          <Button appearance="transparent" size="small" onClick={() => setError('')}>
            Dismiss
          </Button>
        </div>
      )}
      {notice && (
        <div role="status" className={s.notice}>
          {notice}
          <Button
            appearance="transparent"
            size="small"
            aria-label="Dismiss message"
            icon={<DismissRegular />}
            onClick={() => setNotice('')}
          />
        </div>
      )}
      {config && !config.entity.enabled && (
        <div className={s.empty}>Documents are disabled for this record type.</div>
      )}
      {config && config.migration.total > 0 && config.migration.state !== 'Completed' && (
        <div className={s.filter}>
          <Text>
            Legacy documents: {config.migration.state} · {config.migration.processed} of{' '}
            {config.migration.total}
          </Text>
          <Button
            disabled={busy}
            onClick={() =>
              void run(async (signal) => {
                let migration = await client.migration(true, signal);
                while (migration.state === 'Running' && !signal.aborted) {
                  setConfig((value) => (value ? { ...value, migration } : value));
                  await new Promise((resolve) => setTimeout(resolve, 1500));
                  if (signal.aborted) return;
                  migration = await client.migration(false, signal);
                }
                if (!signal.aborted) {
                  setConfig((value) => (value ? { ...value, migration } : value));
                  refresh();
                  setNotice(
                    migration.state === 'Completed'
                      ? 'Legacy documents migrated.'
                      : 'Migration failed. Try again.',
                  );
                }
              })
            }
          >
            Migrate documents
          </Button>
        </div>
      )}
      {filters && (
        <div className={s.filter}>
          <Field label="Document type">
            <Dropdown
              aria-label="Filter document type"
              value={
                attributes
                  .find((a) => a.LogicalName === 'dms_documenttype')
                  ?.OptionSet?.Options?.find((o) => String(o.Value) === query.type)?.Label
                  .UserLocalizedLabel.Label || 'All types'
              }
              selectedOptions={[query.type]}
              onOptionSelect={(_, data) => changeQuery({ type: data.optionValue || '' })}
            >
              <Option value="">All types</Option>
              {attributes
                .find((a) => a.LogicalName === 'dms_documenttype')
                ?.OptionSet?.Options?.map((option) => (
                  <Option key={option.Value} value={String(option.Value)}>
                    {option.Label.UserLocalizedLabel.Label}
                  </Option>
                ))}
            </Dropdown>
          </Field>
          <Field label="Status">
            <Dropdown
              aria-label="Filter status"
              value={
                attributes
                  .find((a) => a.LogicalName === 'dms_documentstatus')
                  ?.OptionSet?.Options?.find((o) => String(o.Value) === query.status)?.Label
                  .UserLocalizedLabel.Label || 'All statuses'
              }
              selectedOptions={[query.status]}
              onOptionSelect={(_, data) => changeQuery({ status: data.optionValue || '' })}
            >
              <Option value="">All statuses</Option>
              {attributes
                .find((a) => a.LogicalName === 'dms_documentstatus')
                ?.OptionSet?.Options?.map((option) => (
                  <Option key={option.Value} value={String(option.Value)}>
                    {option.Label.UserLocalizedLabel.Label}
                  </Option>
                ))}
            </Dropdown>
          </Field>
          <Button
            onClick={() => {
              setSearch('');
              changeQuery(initialQuery);
            }}
          >
            Clear filters
          </Button>
        </div>
      )}
      <div className={s.workspace}>
        <div
          ref={dragSelection.container}
          {...dragSelection.handlers}
          className={s.main}
          aria-busy={loading}
          style={{ userSelect: dragSelection.rectangle ? 'none' : undefined }}
        >
          {!config && !error && (
            <div className={s.empty}>
              <Spinner label="Loading documents…" />
            </div>
          )}
          {canManage && view === 'List' && rows.length > 0 && (
            <div className={s.gridScroll}>
              <DocumentGrid
                rows={rows}
                selected={selected}
                setSelected={setSelected}
                select={select}
                host={host}
                onPreview={(id) => {
                  setSelected([id]);
                  setDialog('Preview');
                }}
                onDetails={(id) => {
                  setSelected([id]);
                  setDetails(true);
                }}
                onCopy={(row) => void run(async () => copy([row]))}
                sort={query.sort}
                descending={query.descending}
                onSort={(key) =>
                  changeQuery({
                    sort: key,
                    descending: query.sort === key ? !query.descending : false,
                  })
                }
                visibleColumns={columns}
                widths={columnWidths}
                onResize={(id, width) => setColumnWidths((value) => ({ ...value, [id]: width }))}
                onReorder={(source, target) =>
                  setColumns((value) => moveColumn(value, source, target))
                }
                extraColumns={attributes}
                compact={compact}
              />
            </div>
          )}
          {canManage && view === 'Tiles' && rows.length > 0 && (
            <>
              <div className={s.tileStrip}>
                <Checkbox
                  shape="circular"
                  label="Select all documents on this page"
                  checked={
                    selected.length === rows.length ? true : selected.length ? 'mixed' : false
                  }
                  onChange={() =>
                    setSelected(selected.length === rows.length ? [] : rows.map((row) => row.id))
                  }
                />
                <Button
                  appearance="subtle"
                  onClick={() =>
                    changeQuery({
                      sort: 'dms_name',
                      descending: query.sort === 'dms_name' ? !query.descending : false,
                    })
                  }
                >
                  Sort: Name <ChevronDownRegular />
                </Button>
              </div>
              <div
                className={s.tiles}
                role="listbox"
                aria-label="Documents"
                aria-multiselectable="true"
              >
                {rows.map((row) => (
                  <div
                    key={row.id}
                    data-document-id={row.id}
                    role="option"
                    aria-selected={selected.includes(row.id)}
                    tabIndex={0}
                    className={mergeClasses(s.tile, selected.includes(row.id) && s.tileSelected)}
                    onClick={(event) => select(row.id, event)}
                    onDoubleClick={() => {
                      setSelected([row.id]);
                      setDialog('Preview');
                    }}
                    onKeyDown={(event) => {
                      if (event.key === ' ') {
                        event.preventDefault();
                        select(row.id, event, true);
                      }
                      if (event.key === 'Enter') {
                        setSelected([row.id]);
                        setDialog('Preview');
                      }
                    }}
                  >
                    <div className={s.tilePreview}>
                      <DocumentThumbnail
                        doc={row}
                        host={host}
                        icon={<FileIcon doc={row} large />}
                      />
                    </div>
                    <Checkbox
                      className={s.tileCheck}
                      shape="circular"
                      checked={selected.includes(row.id)}
                      aria-label={`Select ${row.name}`}
                      onClick={(event) => event.stopPropagation()}
                      onChange={() =>
                        setSelected(
                          selected.includes(row.id)
                            ? selected.filter((id) => id !== row.id)
                            : [...selected, row.id],
                        )
                      }
                    />
                    <Button
                      className={s.tileMore}
                      appearance="subtle"
                      aria-label={`Details for ${row.name}`}
                      icon={<EditRegular />}
                      onClick={(event) => {
                        event.stopPropagation();
                        setSelected([row.id]);
                        setDetails(true);
                      }}
                    />
                    <div className={s.tileBody}>
                      <div className={s.tileName} title={row.name}>
                        {row.name}
                      </div>
                      <DocumentPerson name={row.by} />
                      <Text className={s.secondary}>{row.modified}</Text>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
          {canManage && !rows.length && !loading && (
            <div className={s.empty}>
              <DocumentRegular fontSize={48} />
              <p>
                {count === 0 &&
                (query.search || query.type || query.status || query.preset !== 'All documents')
                  ? 'No matching documents'
                  : t('empty.title')}
              </p>
              <Text>{t('empty.description')}</Text>
            </div>
          )}
          {canManage && loading && (
            <div style={{ padding: 24 }}>
              <Spinner label="Loading documents…" />
            </div>
          )}
          {next && !loading && (
            <div style={{ padding: 20, textAlign: 'center' }}>
              <Button onClick={loadMore}>Load more documents</Button>
            </div>
          )}
        </div>
        {details && host.showDetailsPane !== false && (
          <InlineDrawer open position="end" className={s.pane}>
            <div className={s.paneHeader}>
              <Text weight="semibold">Details</Text>
              <Button
                appearance="subtle"
                aria-label="Close details"
                icon={<DismissRegular />}
                onClick={() => setDetails(false)}
              />
            </div>
            {current && picked.length === 1 ? (
              <>
                <div className={s.paneHero}>
                  <div className={s.paneThumbnail}>
                    <DocumentThumbnail
                      doc={current}
                      host={host}
                      icon={<FileIcon doc={current} large />}
                    />
                  </div>
                  <div className={s.paneName}>{current.name}</div>
                  <Text className={s.secondary}>
                    {current.fileType} document · {current.size}
                  </Text>
                  <div className={s.paneActions}>
                    <Button size="small" icon={<OpenRegular />} onClick={() => action('Open')}>
                      Open
                    </Button>
                    <Button
                      size="small"
                      icon={<ArrowDownloadRegular />}
                      onClick={() => action('Download')}
                    >
                      Download
                    </Button>
                  </div>
                </div>
                <div className={s.paneSection}>
                  <div className={s.sectionTitle}>
                    Properties
                    <Button
                      appearance="subtle"
                      aria-label="Edit properties"
                      icon={<EditRegular />}
                      onClick={() => action('Edit details')}
                    />
                  </div>
                  <div className={s.props}>
                    <Text>Document type</Text>
                    <span>{current.type}</span>
                    <Text>Status</Text>
                    <span>{status(current)}</span>
                    <Text>Expiry date</Text>
                    <span>{current.expiry || '—'}</span>
                    <Text>Description</Text>
                    <span style={{ overflowWrap: 'anywhere' }}>
                      {String(current.raw.dms_description || '—')}
                    </span>
                  </div>
                </div>
                <div className={s.paneSection}>
                  <div className={s.sectionTitle}>Activity</div>
                  <DocumentPerson name={current.by} secondaryText="Modified by" />
                  <div style={{ marginTop: 14 }}>
                    <DocumentPerson
                      name={label(current.raw, '_createdby_value')}
                      secondaryText="Created by"
                    />
                  </div>
                  <div className={s.props} style={{ marginTop: 14 }}>
                    <Text>Modified</Text>
                    <span>{current.modified}</span>
                    <Text>Created</Text>
                    <span>{label(current.raw, 'createdon')}</span>
                  </div>
                </div>
                <div className={s.paneSection}>
                  <div className={s.sectionTitle}>Storage</div>
                  <div className={s.props}>
                    <Text>Source</Text>
                    <span>{current.source}</span>
                    <Text>File size</Text>
                    <span>{current.size}</span>
                  </div>
                  {current.source === 'SharePoint' && (
                    <div className={s.hint}>
                      Access to this file is controlled by SharePoint permissions.
                    </div>
                  )}
                </div>
              </>
            ) : (
              <div className={s.empty}>
                <DocumentRegular fontSize={44} />
                <p>
                  {picked.length
                    ? `${picked.length} documents selected`
                    : 'Select a document to see details'}
                </p>
                {picked.length > 1 && (
                  <Button onClick={() => action('Edit details')}>Edit details</Button>
                )}
              </div>
            )}
          </InlineDrawer>
        )}
      </div>
      {dragSelection.rectangle && (
        <div
          aria-hidden="true"
          data-testid="selection-marquee"
          style={{
            position: 'fixed',
            ...dragSelection.rectangle,
            border: `1px solid ${tokens.colorBrandStroke1}`,
            backgroundColor: tokens.colorBrandBackground2,
            opacity: 0.5,
            pointerEvents: 'none',
            zIndex: 30,
          }}
        />
      )}
      <div className={s.footer}>
        <span>
          {rows.length} of {count} items{selected.length ? ` · ${selected.length} selected` : ''}
        </span>
        <span>{next ? 'More documents available' : 'All documents loaded'}</span>
      </div>
      <input
        ref={inputRef}
        type="file"
        multiple
        hidden
        aria-label="Choose files to upload"
        onChange={(event) => {
          queueFiles(Array.from(event.target.files || []));
          event.target.value = '';
        }}
      />
      <OverlayDrawer
        open={dialog === 'Edit details'}
        position="end"
        className={s.editDrawer}
        onOpenChange={(_, data) => {
          if (!data.open) close();
        }}
      >
        <DrawerHeader>
          <DrawerHeaderTitle
            action={
              <Button
                disabled={busy}
                appearance="subtle"
                aria-label="Close edit details"
                icon={<DismissRegular />}
                onClick={close}
              />
            }
          >
            Edit details
          </DrawerHeaderTitle>
          <Text>{picked.length === 1 ? current?.name : `${picked.length} documents selected`}</Text>
        </DrawerHeader>
        <DrawerBody>
          {error && <p role="alert">{error}</p>}
          <div className={s.dialogFields}>
            <MetadataFields
              attributes={editAttributes}
              values={editValues}
              bulk={picked.length > 1}
              onChange={(key, value) => {
                setEditValues((previous) => ({ ...previous, [key]: value }));
                setDirty((previous) => ({ ...previous, [key]: value }));
              }}
            />
          </div>
        </DrawerBody>
        <DrawerFooter>
          <Button
            appearance="primary"
            disabled={busy || !Object.keys(dirty).length}
            onClick={() =>
              void run(async (signal) => {
                for (const attribute of editAttributes) {
                  if (
                    attribute.RequiredLevel.Value === 'ApplicationRequired' &&
                    attribute.LogicalName in dirty &&
                    !String(dirty[attribute.LogicalName] ?? '').trim()
                  )
                    throw new Error(
                      `${attribute.DisplayName.UserLocalizedLabel.Label} is required.`,
                    );
                  if (
                    attribute.AttributeType === 'Lookup' &&
                    dirty[attribute.LogicalName] &&
                    !/^[0-9a-f-]{36}$/i.test(String(dirty[attribute.LogicalName]))
                  )
                    throw new Error('Enter a valid user GUID.');
                }
                let saved = 0;
                try {
                  for (const row of picked) {
                    await client.save(row, dirty, signal);
                    saved++;
                  }
                  if (!signal.aborted) {
                    setDialog('');
                    setNotice(`${saved} document${saved === 1 ? '' : 's'} updated.`);
                    refresh();
                  }
                } catch (e) {
                  if (saved) {
                    setDialog('');
                    refresh();
                    throw new Error(
                      `${saved} document(s) saved. ${e instanceof Error ? e.message : 'Remaining updates failed.'}`,
                    );
                  }
                  throw e;
                }
              })
            }
          >
            {busy ? 'Saving…' : 'Save'}
          </Button>
          <Button disabled={busy} onClick={close}>
            Cancel
          </Button>
        </DrawerFooter>
      </OverlayDrawer>
      <Dialog
        open={Boolean(dialog) && dialog !== 'Edit details'}
        onOpenChange={(_, data) => {
          if (!data.open) close();
        }}
      >
        <DialogSurface style={dialog === 'Preview' ? { maxWidth: 900, width: '90vw' } : undefined}>
          <DialogBody>
            <DialogTitle>
              {dialog === 'Preview'
                ? current?.name
                : dialog === 'Custom action'
                  ? pendingAction?.action.label
                  : dialog}
            </DialogTitle>
            <DialogContent>
              {error && <p role="alert">{error}</p>}
              {dialog === 'Custom action' && pendingAction && (
                <p>
                  {pendingAction.action.confirmMessage}
                  <br />
                  {pendingAction.rows.length} document(s) selected.
                </p>
              )}
              {dialog === 'Delete' && (
                <p>
                  {picked.length} document{picked.length === 1 ? '' : 's'} selected.
                  <br />
                  {picked.some((row) => row.source === 'Note') && (
                    <>
                      This permanently deletes the local document and its owned file.
                      <br />
                    </>
                  )}
                  {picked.some((row) => row.source === 'SharePoint') && (
                    <>This removes the reference. The SharePoint file remains in place.</>
                  )}
                </p>
              )}
              {dialog === 'Add link' && (
                <div className={s.dialogFields}>
                  <Field label="SharePoint URL" required>
                    <Input
                      value={linkUrl}
                      onChange={(_, data) => setLinkUrl(data.value)}
                      placeholder="https://contoso.sharepoint.com/…"
                    />
                  </Field>
                  <Field label="Display name">
                    <Input
                      value={linkName}
                      maxLength={255}
                      onChange={(_, data) => setLinkName(data.value)}
                    />
                  </Field>
                  <Text className={s.secondary}>
                    The file stays in SharePoint. Its existing permissions apply.
                  </Text>
                </div>
              )}
              {dialog === 'Upload' && (
                <div className={s.dialogFields}>
                  <div className={s.preview} style={{ height: 180 }}>
                    <ArrowUploadRegular fontSize={48} />
                    <Text weight="semibold">Drop files here to upload</Text>
                    <Text className={s.secondary}>Or choose files from your computer.</Text>
                    <Button disabled={busy} onClick={() => inputRef.current?.click()}>
                      Choose files
                    </Button>
                  </div>
                  {files.map((file, index) => (
                    <Text key={`${file.name}-${index}`}>
                      {file.name} · {(file.size / 1024).toFixed(1)} KB
                    </Text>
                  ))}
                  {busy && (
                    <>
                      <Text>{uploadName}</Text>
                      <ProgressBar value={progress} />
                    </>
                  )}
                  {uploadResults.map((result, index) => (
                    <Text key={index}>{result}</Text>
                  ))}
                </div>
              )}
              {dialog === 'Edit columns' && (
                <div className={s.dialogFields}>
                  <p>
                    Choose any columns to show. Drag table headers to reorder and their right edges
                    to resize. Keep at least one column visible.
                  </p>
                  {[
                    ...columns,
                    ...availableColumns(attributes)
                      .map((column) => column.id)
                      .filter((id) => !columns.includes(id)),
                  ].map((id) => {
                    const column = availableColumns(attributes).find((entry) => entry.id === id);
                    if (!column) return null;
                    const index = columns.indexOf(id);
                    return (
                      <div key={id} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <Checkbox
                          style={{ flex: 1 }}
                          label={column.label}
                          checked={index >= 0}
                          disabled={index >= 0 && columns.length === 1}
                          onChange={(_, data) =>
                            setColumns((value) =>
                              data.checked ? [...value, id] : value.filter((key) => key !== id),
                            )
                          }
                        />
                        <Button
                          appearance="subtle"
                          icon={<ArrowUpRegular />}
                          aria-label={`Move ${column.label} earlier`}
                          disabled={index <= 0}
                          onClick={() =>
                            setColumns((value) => moveColumn(value, id, value[index - 1]))
                          }
                        />
                        <Button
                          appearance="subtle"
                          icon={<ArrowDownRegular />}
                          aria-label={`Move ${column.label} later`}
                          disabled={index < 0 || index === columns.length - 1}
                          onClick={() =>
                            setColumns((value) => moveColumn(value, id, value[index + 1]))
                          }
                        />
                      </div>
                    );
                  })}
                  <Button
                    onClick={() => {
                      setColumns(
                        configuredColumns(config?.entity.visibleColumns || [], attributes),
                      );
                      setColumnWidths({});
                    }}
                  >
                    Reset columns
                  </Button>
                </div>
              )}
              {dialog === 'Keyboard shortcuts' && (
                <p>
                  Click to select · Ctrl/Cmd + click to add
                  <br />
                  Shift + click to select a range
                  <br />
                  Space to toggle · Ctrl/Cmd + A to select loaded documents
                  <br />
                  Enter to preview · Esc to clear or close
                </p>
              )}
              {dialog === 'Preview' && current && (
                <div style={{ marginTop: 16 }}>
                  {current.source === 'SharePoint' ? (
                    <div className={s.preview}>
                      <FileIcon doc={current} large />
                      <Text>Open this document in SharePoint to view it.</Text>
                      <Button onClick={() => action('Open')}>Open in SharePoint</Button>
                    </div>
                  ) : previewUrl && current.fileType === 'Photo' ? (
                    <img
                      src={previewUrl}
                      alt={current.name}
                      style={{ width: '100%', maxHeight: 500, objectFit: 'contain' }}
                    />
                  ) : previewUrl && current.fileType === 'PDF' ? (
                    <iframe
                      title={current.name}
                      src={previewUrl}
                      style={{ width: '100%', height: 500, border: 0 }}
                    />
                  ) : previewUrl && current.fileType === 'Video' ? (
                    <video src={previewUrl} controls style={{ width: '100%', maxHeight: 500 }} />
                  ) : current.fileType === 'Text' && previewText ? (
                    <pre style={{ whiteSpace: 'pre-wrap', maxHeight: 500, overflow: 'auto' }}>
                      {previewText}
                    </pre>
                  ) : (
                    <div className={s.preview}>
                      <FileIcon doc={current} large />
                      <Text>
                        {['Word', 'Excel', 'PowerPoint', 'File'].includes(current.fileType)
                          ? 'Download this file to view it in its application.'
                          : current.status === 'Failed' || current.status === 'Pending'
                            ? 'This file is not available.'
                            : 'Loading preview…'}
                      </Text>
                    </div>
                  )}
                </div>
              )}
            </DialogContent>
            <DialogActions>
              <Button disabled={busy} onClick={close}>
                Close
              </Button>
              {dialog === 'Custom action' && pendingAction && (
                <Button
                  appearance="primary"
                  disabled={busy}
                  onClick={() => executeCustomAction(pendingAction.action, pendingAction.rows)}
                >
                  {busy ? 'Running…' : pendingAction.action.label}
                </Button>
              )}
              {busy && dialog === 'Upload' && (
                <Button
                  onClick={() => {
                    actionController.current?.abort();
                    setFiles([]);
                    setUploadName('');
                    setNotice('Cancelling upload…');
                  }}
                >
                  Cancel upload
                </Button>
              )}
              {dialog === 'Upload' && (
                <Button appearance="primary" disabled={busy || !files.length} onClick={upload}>
                  {busy ? 'Uploading…' : 'Upload files'}
                </Button>
              )}
              {dialog === 'Preview' && current && (
                <Button disabled={busy} appearance="primary" onClick={() => action('Download')}>
                  {current.source === 'SharePoint' ? 'Open in SharePoint' : 'Download'}
                </Button>
              )}
              {dialog === 'Add link' && (
                <Button
                  appearance="primary"
                  disabled={busy || !linkUrl.trim()}
                  onClick={() =>
                    void run(async (signal) => {
                      if (!config) return;
                      await client.addLink(linkUrl, linkName, config, signal);
                      if (!signal.aborted) {
                        setDialog('');
                        setNotice('SharePoint link added.');
                        refresh();
                      }
                    })
                  }
                >
                  Add link
                </Button>
              )}
              {dialog === 'Delete' && (
                <Button
                  appearance="primary"
                  disabled={busy || !picked.length}
                  onClick={() =>
                    void run(async (signal) => {
                      let removed = 0;
                      try {
                        for (const row of picked) {
                          await client.remove(row, signal);
                          removed++;
                        }
                        if (!signal.aborted) {
                          setDialog('');
                          setNotice(`${removed} document${removed === 1 ? '' : 's'} deleted.`);
                          refresh();
                        }
                      } catch (e) {
                        if (removed) {
                          setDialog('');
                          refresh();
                          throw new Error(
                            `${removed} document(s) deleted. ${e instanceof Error ? e.message : 'Remaining deletions failed.'}`,
                          );
                        }
                        throw e;
                      }
                    })
                  }
                >
                  {busy ? 'Deleting…' : 'Delete'}
                </Button>
              )}
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </section>
  );
}
export function DmsGridHost({ host }: { host: HostContext }) {
  return (
    <FluentProvider
      theme={{ ...webLightTheme, fontFamilyBase: microsoftFontFamily }}
      style={{ fontFamily: microsoftFontFamily }}
    >
      <Library key={`${host.entityName}:${host.recordId}`} host={host} />
    </FluentProvider>
  );
}
