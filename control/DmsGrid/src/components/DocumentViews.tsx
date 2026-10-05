import * as React from 'react';
import {
  DataGrid,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridBody,
  DataGridRow,
  DataGridCell,
  createTableColumn,
  TableColumnDefinition,
  TableColumnId,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  MenuDivider,
  Toolbar,
  ToolbarButton,
  Checkbox,
  Card,
  Text,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components';
import {
  ChevronDownRegular,
  AddRegular,
  MoreHorizontalRegular,
  EditRegular,
  ArrowDownloadRegular,
  EyeRegular,
  DeleteRegular,
  LinkRegular,
  WarningRegular,
} from '@fluentui/react-icons';
import { DocumentRow, T } from '../types/Documents';
import { DocumentRepository } from '../services/DocumentRepository';
import { Selection } from '../hooks/useSelection';
import { resolveProvider, usable } from '../providers/registry';
import {
  DocumentThumbnail,
  FileTypeIcon,
  DocumentPerson,
  StatusBadge,
  ExpiryBadge,
} from './FileVisuals';
import { formatDate, relativeDate, formatSize } from '../services/format';
import { Action } from './DocumentCommandBar';
export interface ViewProps {
  rows: DocumentRow[];
  selection: Selection;
  repo: DocumentRepository;
  t: T;
  columns: string[];
  orderBy: string;
  onSort: (order: string) => void;
  onHide: (column: string) => void;
  onMove: (column: string, direction: number) => void;
  onColumns: () => void;
  onFilter: () => void;
  onAction: (action: Action, row?: DocumentRow) => void;
  onDetails: (row: DocumentRow) => void;
  compact: boolean;
  details: boolean;
}
const useStyles = makeStyles({
  grid: { minWidth: '1060px', padding: '0 12px', boxSizing: 'border-box' },
  header: { height: '48px', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  headerCell: { fontSize: '14px', fontWeight: tokens.fontWeightSemibold, padding: '0 12px' },
  headerCommand: {
    fontSize: '14px',
    fontWeight: tokens.fontWeightSemibold,
    padding: '0px',
    justifyContent: 'flex-start',
  },
  row: {
    height: '56px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    '&:hover': { backgroundColor: tokens.colorSubtleBackgroundHover },
    '&:hover .row-actions': { opacity: 1 },
    '&:focus-within .row-actions': { opacity: 1 },
  },
  compact: { height: '44px' },
  selected: { backgroundColor: tokens.colorBrandBackground2 },
  cell: { padding: '0 12px', fontSize: '14px', overflow: 'hidden' },
  check: { width: '36px', minWidth: '36px', maxWidth: '36px', padding: 0 },
  name: {
    width: '100%',
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    position: 'relative',
    minWidth: 0,
  },
  title: {
    fontWeight: tokens.fontWeightRegular,
    fontSize: '14px',
    lineHeight: '18px',
    minWidth: 0,
    overflow: 'hidden',
    display: '-webkit-box',
    WebkitBoxOrient: 'vertical',
    WebkitLineClamp: 2,
  },
  secondary: { fontSize: '12px', color: tokens.colorNeutralForeground2 },
  failed: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: '11px',
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  expiry: { display: 'flex', flexDirection: 'column', gap: '3px', alignItems: 'flex-start' },
  actions: {
    position: 'absolute',
    right: 0,
    opacity: 0,
    backgroundColor: tokens.colorSubtleBackgroundHover,
    display: 'flex',
    '& button': { minWidth: '28px', padding: '4px' },
  },
  tileStrip: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '10px 16px',
  },
  tiles: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill,minmax(192px,1fr))',
    gap: '16px',
    padding: '0 20px 24px',
    '@media(max-width:480px)': {
      gridTemplateColumns: 'repeat(2,minmax(0,1fr))',
      gap: '10px',
      padding: '0 10px 16px',
    },
  },
  tile: {
    padding: 0,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    boxShadow: tokens.shadow4,
    position: 'relative',
    cursor: 'pointer',
    transitionProperty: 'box-shadow',
    transitionDuration: tokens.durationNormal,
    '&:hover': { boxShadow: tokens.shadow16 },
    '&:focus-within': { boxShadow: tokens.shadow16 },
    '@media(prefers-reduced-motion:reduce)': { transitionDuration: '0ms' },
    '&:hover .tile-check': { opacity: 1 },
    '&:focus-within .tile-check': { opacity: 1 },
  },
  tileSelected: {
    boxShadow: tokens.shadow8,
    border: `2px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
  },
  hero: { height: '136px', backgroundColor: tokens.colorNeutralBackground2 },
  tileBody: { padding: '16px', display: 'flex', flexDirection: 'column', gap: '8px' },
  tileName: {
    fontSize: '14px',
    lineHeight: '18px',
    height: '36px',
    overflow: 'hidden',
    display: '-webkit-box',
    WebkitBoxOrient: 'vertical',
    WebkitLineClamp: 2,
    overflowWrap: 'anywhere',
  },
  tileCheck: { position: 'absolute', top: '5px', left: '5px', opacity: 0 },
  showCheck: { opacity: 1 },
  tileMore: {
    position: 'absolute',
    right: '5px',
    top: '5px',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  badges: { display: 'flex', gap: '5px', flexWrap: 'wrap' },
});
export function RowMenu({
  row,
  p,
  children,
  context = false,
}: {
  row: DocumentRow;
  p: ViewProps;
  children: React.ReactElement;
  context?: boolean;
}) {
  const provider = resolveProvider(row.provider, p.repo.client),
    can = usable(row, provider);
  return (
    <Menu openOnContext={context}>
      <MenuTrigger disableButtonEnhancement>{children}</MenuTrigger>
      <MenuPopover>
        <MenuList>
          {can && (
            <MenuItem icon={<EyeRegular />} onClick={() => p.onAction('preview', row)}>
              {p.t('action.preview')}
            </MenuItem>
          )}
          {can && provider.capabilities.canDownload && (
            <MenuItem icon={<ArrowDownloadRegular />} onClick={() => p.onAction('download', row)}>
              {p.t('action.download')}
            </MenuItem>
          )}
          {provider.capabilities.canOpen && (
            <MenuItem icon={<LinkRegular />} onClick={() => p.onAction('copy', row)}>
              {p.t('action.copy')}
            </MenuItem>
          )}
          {provider.capabilities.canOpen && (
            <MenuItem icon={<EditRegular />} onClick={() => p.onAction('edit', row)}>
              {p.t('action.edit')}
            </MenuItem>
          )}
          <MenuDivider />
          <MenuItem onClick={() => p.onDetails(row)}>{p.t('details.title')}</MenuItem>
          {provider.capabilities.canDelete && (
            <MenuItem icon={<DeleteRegular />} onClick={() => p.onAction('delete', row)}>
              {p.t('action.delete')}
            </MenuItem>
          )}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
export function DocumentList(p: ViewProps) {
  const s = useStyles();
  const header = (field: string) => (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <ToolbarButton
          className={s.headerCommand}
          appearance="transparent"
          icon={<ChevronDownRegular />}
          aria-label={p.t('column.menu', { name: p.t('column.' + field) })}
        >
          {p.t('column.' + field)}
        </ToolbarButton>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuItem onClick={() => p.onSort(field + ' asc')}>{p.t('sort.ascending')}</MenuItem>
          <MenuItem onClick={() => p.onSort(field + ' desc')}>{p.t('sort.descending')}</MenuItem>
          <MenuDivider />
          <MenuItem onClick={p.onFilter}>{p.t('filters.title')}</MenuItem>
          {field !== 'dms_name' && (
            <MenuItem onClick={() => p.onHide(field)}>{p.t('columns.hide')}</MenuItem>
          )}
          <MenuItem onClick={() => p.onMove(field, -1)}>{p.t('columns.left')}</MenuItem>
          <MenuItem onClick={() => p.onMove(field, 1)}>{p.t('columns.right')}</MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
  const cell = (field: string, row: DocumentRow): React.ReactNode => {
    if (field === 'dms_name')
      return (
        <div className={s.name}>
          <DocumentThumbnail row={row} repo={p.repo} t={p.t} mini />
          <div className={s.title}>
            {row.name}
            {row.uploadState === 100000002 && (
              <div className={s.failed}>
                <WarningRegular />
                {p.t('upload.failedHint')}
              </div>
            )}
            {!resolveProvider(row.provider, p.repo.client).capabilities.canOpen && (
              <div className={s.secondary}>
                <WarningRegular />
                {p.t('provider.unavailable')}
              </div>
            )}
          </div>
          <Toolbar
            className={mergeClasses('row-actions', s.actions)}
            aria-label={p.t('row.actions', { name: row.name })}
          >
            <ToolbarButton
              appearance="subtle"
              icon={<LinkRegular />}
              aria-label={p.t('action.copy')}
              onClick={(e) => {
                e.stopPropagation();
                p.onAction('copy', row);
              }}
            />
            <RowMenu row={row} p={p}>
              <ToolbarButton
                appearance="subtle"
                icon={<MoreHorizontalRegular />}
                aria-label={p.t('row.more', { name: row.name })}
                onClick={(e) => e.stopPropagation()}
              />
            </RowMenu>
          </Toolbar>
        </div>
      );
    if (field === 'modifiedon')
      return <Text className={s.secondary}>{relativeDate(row.modified, p.t)}</Text>;
    if (field === '_modifiedby_value') return <DocumentPerson name={row.modifiedBy} />;
    if (field === 'dms_filesizekb')
      return (
        <Text className={s.secondary}>
          {row.sizeKb ? formatSize(row.sizeKb, p.t) : p.t('empty.value')}
        </Text>
      );
    if (field === 'dms_documenttype') return p.t('type.' + (row.documentType ?? 100000005));
    if (field === 'dms_documentstatus') return <StatusBadge row={row} t={p.t} />;
    if (field === 'dms_expirydate')
      return (
        <div className={s.expiry}>
          <Text className={s.secondary}>{formatDate(row.expiry, p.t)}</Text>
          <ExpiryBadge row={row} t={p.t} />
        </div>
      );
    return String(row.raw[field] ?? p.t('empty.value'));
  };
  const columns: TableColumnDefinition<DocumentRow>[] = React.useMemo(
    () => [
      createTableColumn({
        columnId: 'icon',
        renderHeaderCell: () => null,
        renderCell: (row) => <FileTypeIcon row={row} />,
      }),
      ...p.columns.map((field) =>
        createTableColumn<DocumentRow>({
          columnId: field,
          renderHeaderCell: () => header(field),
          renderCell: (row) => cell(field, row),
        }),
      ),
      createTableColumn({
        columnId: 'add',
        renderHeaderCell: () => (
          <ToolbarButton
            appearance="transparent"
            icon={<AddRegular />}
            onClick={p.onColumns}
            aria-label={p.t('columns.add')}
          ></ToolbarButton>
        ),
        renderCell: () => null,
      }),
    ],
    [p.columns, p.rows, p.t, p.selection.selected, p.orderBy, p.compact, p.details],
  );
  const sizing: Record<
    TableColumnId,
    { minWidth: number; defaultWidth: number; idealWidth: number }
  > = {
    icon: { minWidth: 36, defaultWidth: 36, idealWidth: 36 },
    dms_name: { minWidth: 230, defaultWidth: 290, idealWidth: 300 },
    modifiedon: { minWidth: 140, defaultWidth: 165, idealWidth: 170 },
    _modifiedby_value: { minWidth: 155, defaultWidth: 165, idealWidth: 180 },
    dms_filesizekb: { minWidth: 85, defaultWidth: 95, idealWidth: 100 },
    dms_documenttype: { minWidth: 125, defaultWidth: 125, idealWidth: 135 },
    dms_documentstatus: { minWidth: 90, defaultWidth: 95, idealWidth: 100 },
    dms_expirydate: { minWidth: 150, defaultWidth: 170, idealWidth: 170 },
    add: { minWidth: 40, defaultWidth: 40, idealWidth: 40 },
  };
  return (
    <DataGrid
      items={p.rows}
      columns={columns}
      getRowId={(row) => row.id}
      className={s.grid}
      aria-label={p.t('documents')}
      selectionMode="multiselect"
      selectedItems={p.selection.selected}
      onSelectionChange={(event, data) => {
        const target = event.target as Element,
          id = target.closest('[data-document-id]')?.getAttribute('data-document-id');
        if (id) p.selection.pick(id, event, !!target.closest('.document-select'));
        else p.selection.setSelected(new Set([...data.selectedItems].map(String)));
      }}
      subtleSelection={!p.selection.selected.size}
      selectionAppearance="brand"
      focusMode="row_unstable"
      resizableColumns
      columnSizingOptions={sizing}
    >
      <DataGridHeader>
        <DataGridRow
          className={s.header}
          selectionCell={{
            className: s.check,
            checkboxIndicator: { shape: 'circular', 'aria-label': p.t('selection.all') },
          }}
        >
          {({ renderHeaderCell }) => (
            <DataGridHeaderCell className={s.headerCell}>{renderHeaderCell()}</DataGridHeaderCell>
          )}
        </DataGridRow>
      </DataGridHeader>
      <DataGridBody<DocumentRow>>
        {({ item, rowId }) => (
          <RowMenu key={rowId} row={item} p={p} context>
            <DataGridRow
              key={rowId}
              data-document-id={item.id}
              className={mergeClasses(
                s.row,
                p.compact && s.compact,
                p.selection.selected.has(item.id) && s.selected,
              )}
              selectionCell={{
                className: mergeClasses(s.check, 'document-select'),
                checkboxIndicator: {
                  shape: 'circular',
                  'aria-label': p.t('selection.item', { name: item.name }),
                },
              }}
              onDoubleClick={() => p.onAction('preview', item)}
              onContextMenu={() => p.selection.context(item.id)}
            >
              {({ renderCell }) => (
                <DataGridCell focusMode="group" className={s.cell}>
                  {renderCell(item)}
                </DataGridCell>
              )}
            </DataGridRow>
          </RowMenu>
        )}
      </DataGridBody>
    </DataGrid>
  );
}
export function DocumentTiles(p: ViewProps) {
  const s = useStyles();
  return (
    <>
      <div className={s.tileStrip}>
        <Checkbox
          label={p.t('selection.all')}
          shape="circular"
          checked={
            p.rows.length && p.selection.selected.size === p.rows.length
              ? true
              : p.selection.selected.size
                ? 'mixed'
                : false
          }
          onChange={p.selection.all}
        />
        <Toolbar>
          <ToolbarButton
            onClick={() => p.onSort(p.orderBy.endsWith('asc') ? 'dms_name desc' : 'dms_name asc')}
            icon={<ChevronDownRegular />}
          >
            {p.t('sort.name')}
          </ToolbarButton>
        </Toolbar>
      </div>
      <div
        className={s.tiles}
        data-tiles-grid
        role="listbox"
        aria-label={p.t('documents')}
        aria-multiselectable="true"
      >
        {p.rows.map((row) => (
          <RowMenu key={row.id} row={row} p={p} context>
            <Card
              key={row.id}
              role="option"
              data-document-id={row.id}
              aria-selected={p.selection.selected.has(row.id)}
              tabIndex={0}
              className={mergeClasses(s.tile, p.selection.selected.has(row.id) && s.tileSelected)}
              onClick={(event) => p.selection.pick(row.id, event)}
              onContextMenu={() => p.selection.context(row.id)}
              onDoubleClick={() => p.onAction('preview', row)}
            >
              <div className={s.hero}>
                <DocumentThumbnail row={row} repo={p.repo} t={p.t} />
              </div>
              <span
                className={mergeClasses(
                  'tile-check',
                  s.tileCheck,
                  !!p.selection.selected.size && s.showCheck,
                )}
                onClick={(event) => event.stopPropagation()}
              >
                <Checkbox
                  shape="circular"
                  checked={p.selection.selected.has(row.id)}
                  aria-label={p.t('selection.item', { name: row.name })}
                  onChange={(event) =>
                    p.selection.pick(row.id, event.nativeEvent as MouseEvent, true)
                  }
                />
              </span>
              <Toolbar className={s.tileMore}>
                <RowMenu row={row} p={p}>
                  <ToolbarButton
                    appearance="subtle"
                    icon={<MoreHorizontalRegular />}
                    aria-label={p.t('row.more', { name: row.name })}
                    onClick={(e) => e.stopPropagation()}
                  />
                </RowMenu>
              </Toolbar>
              <div className={s.tileBody}>
                <div className={s.tileName}>{row.name}</div>
                <DocumentPerson name={row.modifiedBy} />
                <Text className={s.secondary}>{relativeDate(row.modified, p.t)}</Text>
                <div className={s.badges}>
                  <StatusBadge row={row} t={p.t} />
                  <ExpiryBadge row={row} t={p.t} />
                </div>
              </div>
            </Card>
          </RowMenu>
        ))}
      </div>
    </>
  );
}
