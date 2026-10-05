import * as React from 'react';
import {
  DataGrid,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridBody,
  DataGridRow,
  DataGridCell,
  createTableColumn,
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Tooltip,
  Badge,
} from '@fluentui/react-components';
import {
  DocumentRegular,
  MoreHorizontalRegular,
  LinkRegular,
  ChevronDownRegular,
  ChevronUpRegular,
  WarningRegular,
  ReOrderDotsVerticalRegular,
} from '@fluentui/react-icons';
import { DocumentRow, Attribute, label } from '../services/documents';
import { HostContext } from '../types/HostContext';
import DocumentThumbnail from './DocumentThumbnail';
import { DocumentPerson } from './DocumentPerson';
import { FileIcon } from './FileIcon';
const useStyles = makeStyles({
  grid: { padding: '0 12px', boxSizing: 'border-box' },
  header: {
    height: '48px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  headerCell: {
    fontSize: '14px',
    fontWeight: tokens.fontWeightSemibold,
    padding: '0 12px',
    '&:hover': { backgroundColor: tokens.colorNeutralBackground1 },
  },
  compact: { minHeight: '40px' },
  row: {
    minHeight: '56px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    '&:hover': { backgroundColor: tokens.colorSubtleBackgroundHover },
    '&:hover .row-actions': { opacity: 1 },
    '&:focus-within .row-actions': { opacity: 1 },
  },
  selected: {
    backgroundColor: tokens.colorBrandBackground2,
    '&:hover': { backgroundColor: tokens.colorBrandBackground2Hover },
  },
  cell: { padding: '0 12px', fontSize: '14px', overflow: 'hidden' },
  selection: { flexBasis: '36px', width: '36px', minWidth: '36px', maxWidth: '36px', padding: 0 },
  nameCell: {
    position: 'relative',
    width: '100%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '12px',
    minWidth: 0,
  },
  fileName: {
    flex: 1,
    fontWeight: tokens.fontWeightRegular,
    fontSize: '14px',
    overflow: 'hidden',
    whiteSpace: 'normal',
    lineHeight: '18px',
    display: '-webkit-box',
    WebkitBoxOrient: 'vertical',
    WebkitLineClamp: 2,
    textOverflow: 'ellipsis',
    minWidth: 0,
  },
  actions: {
    position: 'absolute',
    right: 0,
    backgroundColor: tokens.colorSubtleBackgroundHover,
    display: 'flex',
    opacity: 0,
    flexShrink: 0,
    gap: '2px',
    '& button': { minWidth: '28px', padding: '4px', color: tokens.colorNeutralForeground2 },
  },
  secondary: {
    fontSize: '13px',
    color: tokens.colorNeutralForeground2,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  badge: {
    fontSize: '12px',
    height: '22px',
    padding: '0 8px',
    fontWeight: tokens.fontWeightRegular,
  },
  expiryCell: {
    display: 'flex',
    flexDirection: 'column',
    gap: '3px',
    alignItems: 'flex-start',
    padding: '5px 0',
  },
  failed: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    fontSize: '11px',
    color: tokens.colorPaletteRedForeground1,
  },
});
interface Props {
  rows: DocumentRow[];
  selected: string[];
  setSelected: (ids: string[]) => void;
  select: (
    id: string,
    event: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean },
    toggle?: boolean,
  ) => void;
  host: HostContext;
  onPreview: (id: string) => void;
  onDetails: (id: string) => void;
  onCopy: (row: DocumentRow) => void;
  sort: string;
  descending: boolean;
  onSort: (key: string) => void;
  extraColumns: Attribute[];
  compact: boolean;
  visibleColumns: string[];
  widths: Record<string, number>;
  onResize: (id: string, width: number) => void;
  onReorder: (source: string, target: string) => void;
}
export default function DocumentGrid({
  rows,
  selected,
  setSelected,
  select,
  host,
  onPreview,
  onDetails,
  onCopy,
  sort,
  descending,
  onSort,
  extraColumns,
  compact,
  visibleColumns,
  widths,
  onResize,
  onReorder,
}: Props) {
  const s = useStyles();
  const [draggingColumn, setDraggingColumn] = React.useState<string>();
  const head = (text: string, key: string) => (
    <Button
      appearance="transparent"
      onClick={() => onSort(key)}
      icon={sort === key && !descending ? <ChevronUpRegular /> : <ChevronDownRegular />}
      iconPosition="after"
      style={{ paddingLeft: 0, fontWeight: 600 }}
    >
      {text}
    </Button>
  );
  const expiry = (d: DocumentRow) => {
    return (
      <div className={s.expiryCell}>
        <span className={s.secondary}>
          {d.expiry
            ? new Date(d.expiry + 'T00:00:00').toLocaleDateString(undefined, {
                dateStyle: 'medium',
              })
            : '—'}
        </span>
      </div>
    );
  };
  const columns = [
    createTableColumn<DocumentRow>({
      columnId: 'icon',
      renderHeaderCell: () => <DocumentRegular aria-label="File type" />,
      renderCell: (d) => <FileIcon doc={d} />,
    }),
    createTableColumn<DocumentRow>({
      columnId: 'name',
      renderHeaderCell: () => head('Name', 'dms_name'),
      renderCell: (d) => (
        <div className={s.nameCell}>
          <DocumentThumbnail doc={d} host={host} miniature icon={<FileIcon doc={d} />} />
          <div className={s.fileName} title={d.name}>
            {d.name}
            {d.status === 'Failed' && (
              <div className={s.failed}>
                <WarningRegular /> Upload failed
              </div>
            )}
          </div>
          <div className={mergeClasses('row-actions', s.actions)}>
            <Tooltip content="Copy link" relationship="label">
              <Button
                appearance="subtle"
                icon={<LinkRegular />}
                aria-label={`Copy link to ${d.name}`}
                onClick={(e) => {
                  e.stopPropagation();
                  onCopy(d);
                }}
              />
            </Tooltip>
            <Tooltip content="More actions" relationship="label">
              <Button
                appearance="subtle"
                icon={<MoreHorizontalRegular />}
                aria-label={`Details for ${d.name}`}
                onClick={(e) => {
                  e.stopPropagation();
                  onDetails(d.id);
                }}
              />
            </Tooltip>
          </div>
        </div>
      ),
    }),
    createTableColumn<DocumentRow>({
      columnId: 'modified',
      renderHeaderCell: () => head('Modified', 'modifiedon'),
      renderCell: (d) => <span className={s.secondary}>{d.modified}</span>,
    }),
    createTableColumn<DocumentRow>({
      columnId: 'by',
      renderHeaderCell: () => head('Modified By', '_modifiedby_value'),
      renderCell: (d) => <DocumentPerson name={d.by} />,
    }),
    ...[
      createTableColumn<DocumentRow>({
        columnId: 'type',
        renderHeaderCell: () => head('Document type', 'dms_documenttype'),
        renderCell: (d) => <span>{d.type}</span>,
      }),
      createTableColumn<DocumentRow>({
        columnId: 'status',
        renderHeaderCell: () => head('Status', 'dms_documentstatus'),
        renderCell: (d) => (
          <Badge
            className={s.badge}
            appearance="tint"
            color={
              d.status === 'Failed'
                ? 'danger'
                : d.status === 'Draft' || d.status === 'Pending'
                  ? 'subtle'
                  : d.status === 'Archived'
                    ? 'informative'
                    : 'success'
            }
          >
            {d.status}
          </Badge>
        ),
      }),
      createTableColumn<DocumentRow>({
        columnId: 'expiry',
        renderHeaderCell: () => head('Expiry date', 'dms_expirydate'),
        renderCell: expiry,
      }),
      ...extraColumns.map((attribute) =>
        createTableColumn<DocumentRow>({
          columnId: attribute.LogicalName,
          renderHeaderCell: () =>
            head(attribute.DisplayName.UserLocalizedLabel.Label, attribute.LogicalName),
          renderCell: (d) => <span>{label(d.raw, attribute.LogicalName)}</span>,
        }),
      ),
    ],
  ];
  const orderedColumns = visibleColumns.flatMap((id) => {
    const column = columns.find((entry) => entry.columnId === id);
    return column ? [column] : [];
  });
  const defaults: Record<string, number> = {
    icon: 72,
    name: 300,
    modified: 170,
    by: 176,
    type: 150,
    status: 110,
    expiry: 160,
  };
  const sizing = Object.fromEntries(
    visibleColumns.map((id) => [
      id,
      {
        minWidth: id === 'icon' ? 64 : id === 'name' ? 160 : 100,
        defaultWidth: widths[id] ?? defaults[id] ?? 170,
        idealWidth: widths[id] ?? defaults[id] ?? 170,
      },
    ]),
  );
  return (
    <DataGrid
      items={rows}
      columns={orderedColumns}
      resizableColumns
      resizableColumnsOptions={{ autoFitColumns: false }}
      columnSizingOptions={sizing}
      onColumnResize={(_, data) => onResize(String(data.columnId), data.width)}
      getRowId={(d) => d.id}
      selectionMode="multiselect"
      selectedItems={new Set(selected)}
      onSelectionChange={(_, data) => setSelected([...data.selectedItems].map(String))}
      subtleSelection={!selected.length}
      selectionAppearance="neutral"
      focusMode="composite"
      aria-label="Documents"
      className={s.grid}
    >
      <DataGridHeader>
        <DataGridRow
          className={s.header}
          selectionCell={{
            subtle: false,
            className: s.selection,
            checkboxIndicator: {
              shape: 'circular',
              'aria-label': 'Select all documents on this page',
            },
          }}
        >
          {({ renderHeaderCell, columnId }) => (
            <DataGridHeaderCell
              className={s.headerCell}
              data-column-id={columnId}
              title="Drag to reorder; drag the right edge to resize"
              style={{ opacity: draggingColumn === String(columnId) ? 0.55 : 1 }}
              onPointerDown={(event: React.PointerEvent<HTMLElement>) => event.stopPropagation()}
            >
              <Button
                appearance="subtle"
                size="small"
                icon={<ReOrderDotsVerticalRegular />}
                aria-label={`Drag ${String(columnId)} column`}
                title="Drag to reorder column"
                onPointerDown={(event: React.PointerEvent<HTMLElement>) => {
                  if (event.button !== 0) return;
                  event.preventDefault();
                  event.stopPropagation();
                  event.currentTarget.setPointerCapture(event.pointerId);
                  setDraggingColumn(String(columnId));
                }}
                onPointerUp={(event: React.PointerEvent<HTMLElement>) => {
                  event.stopPropagation();
                  if (!draggingColumn) return;
                  const target = event.currentTarget.ownerDocument
                    .elementFromPoint(event.clientX, event.clientY)
                    ?.closest('[role="columnheader"][data-column-id]');
                  const targetId = target?.getAttribute('data-column-id');
                  if (targetId) onReorder(draggingColumn, targetId);
                  if (event.currentTarget.hasPointerCapture(event.pointerId))
                    event.currentTarget.releasePointerCapture(event.pointerId);
                  setDraggingColumn(undefined);
                }}
                onPointerCancel={() => setDraggingColumn(undefined)}
                style={{
                  minWidth: 20,
                  padding: 0,
                  flexShrink: 0,
                  cursor: draggingColumn ? 'grabbing' : 'grab',
                  touchAction: 'none',
                }}
              />
              {renderHeaderCell()}
            </DataGridHeaderCell>
          )}
        </DataGridRow>
      </DataGridHeader>
      <DataGridBody<DocumentRow>>
        {({ item, rowId }) => (
          <DataGridRow<DocumentRow>
            key={rowId}
            data-document-id={rowId}
            className={mergeClasses(
              s.row,
              compact && s.compact,
              selected.includes(String(rowId)) && s.selected,
            )}
            selectionCell={{
              className: mergeClasses(s.selection, 'select-cell'),
              checkboxIndicator: { shape: 'circular', 'aria-label': `Select ${item.name}` },
            }}
            onClick={(e: React.MouseEvent<HTMLTableRowElement>) => {
              if (!(e.target as HTMLElement).closest('button,input,.select-cell'))
                select(String(rowId), e);
            }}
            onDoubleClick={() => onPreview(String(rowId))}
            onKeyDown={(e: React.KeyboardEvent<HTMLTableRowElement>) => {
              if (e.key === 'Enter') onPreview(String(rowId));
            }}
          >
            {({ renderCell, columnId }) => (
              <DataGridCell focusMode="group" className={s.cell} data-column-id={columnId}>
                {renderCell(item)}
              </DataGridCell>
            )}
          </DataGridRow>
        )}
      </DataGridBody>
    </DataGrid>
  );
}
