import { CustomAction, actionEnabled } from '../services/customActions';
import * as React from 'react';
import { useState } from 'react';
import {
  Toolbar,
  ToolbarButton,
  Overflow,
  OverflowItem,
  useOverflowMenu,
  useIsOverflowItemVisible,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  MenuItemRadio,
  MenuDivider,
  Tooltip,
  Input,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  AddRegular,
  SendRegular,
  CheckmarkRegular,
  DocumentRegular,
  ArrowUploadRegular,
  ArrowSyncRegular,
  MoreHorizontalRegular,
  ChevronDownRegular,
  SearchRegular,
  TextBulletListRegular,
  GridRegular,
  GlobeRegular,
  FilterRegular,
  InfoRegular,
  ArrowMaximizeRegular,
  ArrowMinimizeRegular,
  DismissRegular,
  OpenRegular,
  EyeRegular,
  ArrowDownloadRegular,
  LinkRegular,
  DeleteRegular,
  EditRegular,
} from '@fluentui/react-icons';
const useStyles = makeStyles({
  root: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    height: '56px',
    padding: '0 20px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    boxSizing: 'border-box',
    width: '100%',
    minWidth: 0,
    overflow: 'hidden',
    flexWrap: 'nowrap',
    '@media(max-width:720px)': { padding: '0 12px' },
  },
  item: { display: 'flex', flexShrink: 0, alignItems: 'center' },
  spacer: { flexGrow: 1, minWidth: '8px' },
  primary: { '& svg': { color: tokens.colorNeutralForegroundOnBrand } },
  command: { whiteSpace: 'nowrap', '& svg': { color: tokens.colorBrandForeground1 } },
  label: { '@media(max-width:720px)': { display: 'none' } },
  search: {
    width: '190px',
    '@media(max-width:1000px)': { width: '150px' },
    '@media(max-width:720px)': { width: '140px' },
  },
  toggle: { backgroundColor: tokens.colorBrandBackground2 },
  overflow: { flexShrink: 0 },
});
interface Entry {
  id: string;
  label: string;
  icon: React.ReactElement;
  onClick?: () => void;
  priority: number;
  node?: React.ReactNode;
  overflowNode?: React.ReactNode;
  appearance?: 'primary';
  iconOnly?: boolean;
  pressed?: boolean;
  showLabel?: boolean;
  disabled?: boolean;
}
interface Props {
  customActions?: CustomAction[];
  onCustomAction?: (action: CustomAction) => void;
  busy?: boolean;
  provider: string;
  selectionCount: number;
  onAction: (kind: string) => void;
  onUpload: () => void;
  onAddLink: () => void;
  onRefresh: () => void;
  onClear: () => void;
  query: string;
  onSearch: (value: string) => void;
  preset: string;
  onPreset: (value: string) => void;
  view: string;
  onView: (value: string) => void;
  filters: boolean;
  onFilters: () => void;
  details: boolean;
  onDetails: () => void;
  expanded: boolean;
  onExpand: () => void;
  feedback: (value: string) => void;
  onColumns: () => void;
  onShortcuts: () => void;
  onDensity: () => void;
  compact: boolean;
}
function OverflowEntry({ entry }: { entry: Entry }) {
  const visible = useIsOverflowItemVisible(entry.id);
  if (visible) return null;
  if (entry.overflowNode) return <>{entry.overflowNode}</>;
  return (
    <MenuItem icon={entry.icon} onClick={entry.onClick} disabled={entry.disabled}>
      {entry.label}
    </MenuItem>
  );
}
function OverflowCommands({
  entries,
  onColumns,
  onShortcuts,
  onDensity,
  compact,
}: {
  entries: Entry[];
  onColumns: () => void;
  onShortcuts: () => void;
  onDensity: () => void;
  compact: boolean;
}) {
  const { ref } = useOverflowMenu();
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <ToolbarButton
          ref={ref as React.Ref<HTMLButtonElement>}
          icon={<MoreHorizontalRegular />}
          aria-label="More commands"
          title="More commands"
        />
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {entries.map((entry) => (
            <OverflowEntry key={entry.id} entry={entry} />
          ))}
          <MenuDivider />
          <MenuItem onClick={onDensity}>Density: {compact ? 'Compact' : 'Comfortable'}</MenuItem>
          <MenuItem onClick={onColumns}>Edit columns</MenuItem>
          <MenuDivider />
          <MenuItem onClick={onShortcuts}>Keyboard shortcuts</MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
export default function DocumentCommandBar({
  customActions = [],
  onCustomAction,
  busy = false,
  provider,
  selectionCount,
  onAction,
  onUpload,
  onAddLink,
  onRefresh,
  onClear,
  query,
  onSearch,
  preset,
  onPreset,
  view,
  onView,
  filters,
  onFilters,
  details,
  onDetails,
  expanded,
  onExpand,
  feedback: _feedback,
  onColumns,
  onShortcuts,
  onDensity,
  compact,
}: Props) {
  const s = useStyles(),
    [searchOpen, setSearchOpen] = useState(false);
  const viewItems = (
    <MenuList>
      {['All documents', 'Recently modified', 'Expiring soon', 'Expired', 'Added by me'].map(
        (p) => (
          <MenuItemRadio key={p} name="preset" value={p} onClick={() => onPreset(p)}>
            {p}
          </MenuItemRadio>
        ),
      )}
      <MenuDivider />
      {['List', 'Tiles'].map((v) => (
        <MenuItem
          key={v}
          icon={v === 'List' ? <TextBulletListRegular /> : <GridRegular />}
          onClick={() => onView(v)}
        >
          {v}
          {view === v ? ' ✓' : ''}
        </MenuItem>
      ))}
    </MenuList>
  );
  const make = (
    id: string,
    label: string,
    icon: React.ReactElement,
    onClick: (() => void) | undefined,
    priority = 10,
    extra: Partial<Entry> = {},
  ): Entry => ({ id, label, icon, onClick, priority, ...extra });
  const entries = selectionCount
    ? [
        ...(selectionCount === 1
          ? [
              make('open', 'Open', <OpenRegular />, () => onAction('Open'), 90),
              make('preview', 'Preview', <EyeRegular />, () => onAction('Preview'), 40),
            ]
          : []),
        make('download', 'Download', <ArrowDownloadRegular />, () => onAction('Download'), 70),
        make('copy', 'Copy link', <LinkRegular />, () => onAction('Copy link'), 30),
        make('delete', 'Delete', <DeleteRegular />, () => onAction('Delete'), 20),
        make('edit', 'Edit details', <EditRegular />, () => onAction('Edit details'), 50),
      ]
    : [
        provider === 'Note'
          ? make('add', 'Upload', <ArrowUploadRegular />, onUpload, 100, { appearance: 'primary' })
          : make('add', 'New', <AddRegular />, onAddLink, 100, {
              node: (
                <Menu>
                  <MenuTrigger disableButtonEnhancement>
                    <ToolbarButton appearance="primary" icon={<AddRegular />}>
                      New
                      <ChevronDownRegular />
                    </ToolbarButton>
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      <MenuItem icon={<GlobeRegular />} onClick={onAddLink}>
                        Link
                      </MenuItem>
                    </MenuList>
                  </MenuPopover>
                </Menu>
              ),
              overflowNode: (
                <MenuItem icon={<GlobeRegular />} onClick={onAddLink}>
                  Add link
                </MenuItem>
              ),
            }),
        make('refresh', 'Refresh', <ArrowSyncRegular />, onRefresh, 35),
      ];
  const customIcons: Record<string, React.ReactElement> = {
    Send: <SendRegular />,
    Checkmark: <CheckmarkRegular />,
    Document: <DocumentRegular />,
    Edit: <EditRegular />,
    Link: <LinkRegular />,
    Open: <OpenRegular />,
    Download: <ArrowDownloadRegular />,
  };
  entries.push(
    ...customActions
      .filter((action) => selectionCount > 0 || action.selection.min === 0)
      .map((action) =>
        make(
          `custom:${action.id}`,
          action.label,
          customIcons[action.icon],
          () => onCustomAction?.(action),
          65,
          { disabled: busy || !actionEnabled(action, selectionCount), showLabel: true },
        ),
      ),
  );
  const leftCount = entries.length;
  entries.push(
    selectionCount
      ? make('clear', `${selectionCount} selected`, <DismissRegular />, onClear, 80, {
          showLabel: true,
        })
      : make('search', 'Search documents', <SearchRegular />, () => setSearchOpen(true), 5, {
          node: (
            <Input
              aria-label="Search documents"
              className={s.search}
              contentBefore={<SearchRegular />}
              placeholder="Search documents"
              value={query}
              onChange={(_, data) => onSearch(data.value)}
            />
          ),
        }),
  );
  entries.push(
    make('view', preset, <TextBulletListRegular />, undefined, 25, {
      node: (
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <ToolbarButton
              className={s.command}
              icon={view === 'List' ? <TextBulletListRegular /> : <GridRegular />}
            >
              {preset}
              <ChevronDownRegular />
            </ToolbarButton>
          </MenuTrigger>
          <MenuPopover>{viewItems}</MenuPopover>
        </Menu>
      ),
      overflowNode: (
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <MenuItem icon={<TextBulletListRegular />}>Views and layout</MenuItem>
          </MenuTrigger>
          <MenuPopover>{viewItems}</MenuPopover>
        </Menu>
      ),
    }),
  );
  entries.push(
    make('filters', 'Filters', <FilterRegular />, onFilters, 15, {
      iconOnly: true,
      pressed: filters,
    }),
  );
  entries.push(
    make('details', 'Details', <InfoRegular />, onDetails, 60, {
      iconOnly: true,
      pressed: details,
    }),
  );
  entries.push(
    make(
      'expand',
      expanded ? 'Exit expanded view' : 'Expand',
      expanded ? <ArrowMinimizeRegular /> : <ArrowMaximizeRegular />,
      onExpand,
      0,
      { iconOnly: true },
    ),
  );
  return (
    <>
      <Overflow minimumVisible={1} padding={4} key={selectionCount ? 'selection' : 'default'}>
        <Toolbar className={s.root} aria-label="Document commands">
          {entries.map((entry, i) => (
            <React.Fragment key={entry.id}>
              {i === leftCount && <div className={s.spacer} />}
              <OverflowItem id={entry.id} priority={entry.priority}>
                <div className={s.item}>
                  {entry.node || (
                    <Tooltip content={entry.label} relationship="label">
                      <ToolbarButton
                        appearance={entry.appearance || 'subtle'}
                        disabled={entry.disabled || busy}
                        aria-label={entry.label}
                        aria-pressed={entry.pressed}
                        icon={entry.icon}
                        onClick={entry.onClick}
                        className={
                          entry.appearance ? s.primary : entry.pressed ? s.toggle : s.command
                        }
                      >
                        {!entry.iconOnly && (
                          <span className={entry.showLabel ? undefined : s.label}>
                            {entry.label}
                          </span>
                        )}
                      </ToolbarButton>
                    </Tooltip>
                  )}
                </div>
              </OverflowItem>
            </React.Fragment>
          ))}
          <OverflowCommands
            entries={entries}
            onColumns={onColumns}
            onShortcuts={onShortcuts}
            onDensity={onDensity}
            compact={compact}
          />
        </Toolbar>
      </Overflow>
      <Dialog open={searchOpen} onOpenChange={(_, data) => setSearchOpen(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Search documents</DialogTitle>
            <DialogContent>
              <Input
                aria-label="Search documents"
                value={query}
                onChange={(_, data) => onSearch(data.value)}
                contentBefore={<SearchRegular />}
              />
            </DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={() => setSearchOpen(false)}>
                Done
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
}
