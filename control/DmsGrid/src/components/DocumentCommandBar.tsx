import * as React from 'react';
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
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowUploadRegular,
  AddRegular,
  ArrowSyncRegular,
  SearchRegular,
  ChevronDownRegular,
  MoreHorizontalRegular,
  TextBulletListRegular,
  GridRegular,
  InfoRegular,
  FilterRegular,
  ArrowMaximizeRegular,
  ArrowMinimizeRegular,
  DismissRegular,
  OpenRegular,
  EyeRegular,
  ArrowDownloadRegular,
  DeleteRegular,
  EditRegular,
  LinkRegular,
} from '@fluentui/react-icons';
import { T, Preset, Capabilities } from '../types/Documents';
export type Action = 'open' | 'preview' | 'download' | 'copy' | 'delete' | 'edit';
export interface CommandProps {
  t: T;
  active: Capabilities;
  selectionCount: number;
  available: {
    open: boolean;
    preview: boolean;
    download: boolean;
    delete: boolean;
    edit: boolean;
    copy: boolean;
  };
  busy: boolean;
  onAction: (action: Action) => void;
  onUpload: () => void;
  onAddLink: () => void;
  onRefresh: () => void;
  onClear: () => void;
  search: string;
  onSearch: (value: string) => void;
  preset: Preset;
  onPreset: (value: Preset) => void;
  view: 'list' | 'tiles';
  onView: (value: 'list' | 'tiles') => void;
  filters: boolean;
  onFilters: () => void;
  details: boolean;
  onDetails: () => void;
  expanded: boolean;
  onExpand: () => void;
  onDensity: () => void;
  onColumns: () => void;
  onShortcuts: () => void;
}
interface Entry {
  id: string;
  label: string;
  icon: React.ReactElement;
  onClick?: () => void;
  priority: number;
  node?: React.ReactElement;
  overflowNode?: React.ReactElement;
  iconOnly?: boolean;
  primary?: boolean;
  pressed?: boolean;
  disabled?: boolean;
}
const useStyles = makeStyles({
  root: {
    height: '56px',
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    padding: '0 20px',
    width: '100%',
    minWidth: 0,
    boxSizing: 'border-box',
    overflow: 'hidden',
    flexWrap: 'nowrap',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    '@media(max-width:720px)': { padding: '0 10px' },
  },
  item: { display: 'flex', flexShrink: 0, whiteSpace: 'nowrap' },
  command: { whiteSpace: 'nowrap', '& svg': { color: tokens.colorBrandForeground1 } },
  primary: { whiteSpace: 'nowrap', '& svg': { color: tokens.colorNeutralForegroundOnBrand } },
  spacer: { flexGrow: 1, minWidth: '8px' },
  search: { width: '190px', '@media(max-width:1000px)': { width: '150px' } },
  pressed: { backgroundColor: tokens.colorBrandBackground2 },
  fields: { marginTop: '12px' },
});
function OverflowEntry({ entry }: { entry: Entry }) {
  const visible = useIsOverflowItemVisible(entry.id);
  if (visible) return null;
  return (
    entry.overflowNode || (
      <MenuItem icon={entry.icon} onClick={entry.onClick} disabled={entry.disabled}>
        {entry.label}
      </MenuItem>
    )
  );
}
function OverflowMenu({ entries, p }: { entries: Entry[]; p: CommandProps }) {
  const { ref } = useOverflowMenu();
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <ToolbarButton
          ref={ref as React.Ref<HTMLButtonElement>}
          aria-label={p.t('commands.more')}
          icon={<MoreHorizontalRegular />}
        />
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {entries.map((entry) => (
            <OverflowEntry key={entry.id} entry={entry} />
          ))}
          <MenuDivider />
          <MenuItem onClick={p.onDensity}>{p.t('density.toggle')}</MenuItem>
          <MenuItem onClick={p.onColumns}>{p.t('columns.edit')}</MenuItem>
          <MenuItem onClick={p.onShortcuts}>{p.t('shortcuts.title')}</MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
export function DocumentCommandBar(p: CommandProps) {
  const s = useStyles(),
    [searchOpen, setSearchOpen] = React.useState(false),
    t = p.t;
  const viewItems = (
    <MenuList>
      {(['all', 'recent', 'expiring', 'expired', 'mine'] as Preset[]).map((value) => (
        <MenuItemRadio key={value} name="preset" value={value} onClick={() => p.onPreset(value)}>
          {t('preset.' + value)}
        </MenuItemRadio>
      ))}
      <MenuDivider />
      {(['list', 'tiles'] as const).map((value) => (
        <MenuItem
          key={value}
          icon={value === 'list' ? <TextBulletListRegular /> : <GridRegular />}
          onClick={() => p.onView(value)}
        >
          {t('view.' + value)}
        </MenuItem>
      ))}
    </MenuList>
  );
  const entries: Entry[] = [];
  const add = (
    id: string,
    label: string,
    icon: React.ReactElement,
    onClick: () => void,
    priority: number,
    extra: Partial<Entry> = {},
  ) => entries.push({ id, label, icon, onClick, priority, ...extra });
  if (p.selectionCount) {
    if (p.selectionCount === 1 && p.available.open)
      add('open', t('action.open'), <OpenRegular />, () => p.onAction('open'), 100);
    if (p.selectionCount === 1 && p.available.preview)
      add('preview', t('action.preview'), <EyeRegular />, () => p.onAction('preview'), 40);
    if (p.available.download)
      add(
        'download',
        t('action.download'),
        <ArrowDownloadRegular />,
        () => p.onAction('download'),
        70,
      );
    if (p.available.copy)
      add('copy', t('action.copy'), <LinkRegular />, () => p.onAction('copy'), 20);
    if (p.available.delete)
      add('delete', t('action.delete'), <DeleteRegular />, () => p.onAction('delete'), 30);
    if (p.available.edit)
      add('edit', t('action.edit'), <EditRegular />, () => p.onAction('edit'), 60);
  } else {
    if (p.active.canUpload)
      add('upload', t('action.upload'), <ArrowUploadRegular />, p.onUpload, 100, {
        primary: true,
        disabled: p.busy,
      });
    if (p.active.canAddLink)
      add('link', t('action.new'), <AddRegular />, p.onAddLink, 100, {
        primary: true,
        disabled: p.busy,
        node: (
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <ToolbarButton appearance="primary" icon={<AddRegular />}>
                {t('action.new')}
                <ChevronDownRegular />
              </ToolbarButton>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem onClick={p.onAddLink} icon={<LinkRegular />}>
                  {t('action.addLink')}
                </MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        ),
        overflowNode: (
          <MenuItem icon={<LinkRegular />} onClick={p.onAddLink}>
            {t('action.addLink')}
          </MenuItem>
        ),
      });
    add('refresh', t('action.refresh'), <ArrowSyncRegular />, p.onRefresh, 40, {
      disabled: p.busy,
    });
  }
  const left = entries.length;
  if (p.selectionCount)
    add(
      'clear',
      t('selection.count', { count: p.selectionCount }),
      <DismissRegular />,
      p.onClear,
      90,
    );
  else
    add('search', t('search'), <SearchRegular />, () => setSearchOpen(true), 5, {
      node: (
        <Input
          className={s.search}
          aria-label={t('search')}
          placeholder={t('search')}
          value={p.search}
          contentBefore={<SearchRegular />}
          onChange={(_, data) => p.onSearch(data.value)}
        />
      ),
    });
  add(
    'views',
    t('preset.' + p.preset),
    p.view === 'list' ? <TextBulletListRegular /> : <GridRegular />,
    () => undefined,
    25,
    {
      node: (
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <ToolbarButton
              className={s.command}
              aria-label={t('views.label')}
              icon={p.view === 'list' ? <TextBulletListRegular /> : <GridRegular />}
            >
              {t('preset.' + p.preset)}
              <ChevronDownRegular />
            </ToolbarButton>
          </MenuTrigger>
          <MenuPopover>{viewItems}</MenuPopover>
        </Menu>
      ),
      overflowNode: (
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <MenuItem icon={<TextBulletListRegular />}>{t('views.label')}</MenuItem>
          </MenuTrigger>
          <MenuPopover>{viewItems}</MenuPopover>
        </Menu>
      ),
    },
  );
  add('filters', t('filters.title'), <FilterRegular />, p.onFilters, 15, {
    iconOnly: true,
    pressed: p.filters,
  });
  add('details', t('details.title'), <InfoRegular />, p.onDetails, 50, {
    iconOnly: true,
    pressed: p.details,
  });
  add(
    'expand',
    t(p.expanded ? 'expand.exit' : 'expand'),
    p.expanded ? <ArrowMinimizeRegular /> : <ArrowMaximizeRegular />,
    p.onExpand,
    0,
    { iconOnly: true },
  );
  return (
    <>
      <Overflow minimumVisible={1} padding={4} key={p.selectionCount ? 'selection' : 'default'}>
        <Toolbar className={s.root} aria-label={t('commands.label')}>
          {entries.map((entry, index) => (
            <React.Fragment key={entry.id}>
              {index === left && <div className={s.spacer} />}
              <OverflowItem id={entry.id} priority={entry.priority}>
                <div className={s.item}>
                  {entry.node || (
                    <Tooltip content={entry.label} relationship="label">
                      <ToolbarButton
                        className={
                          entry.primary ? s.primary : entry.pressed ? s.pressed : s.command
                        }
                        appearance={entry.primary ? 'primary' : 'subtle'}
                        aria-label={entry.label}
                        aria-pressed={entry.pressed}
                        icon={entry.icon}
                        onClick={entry.onClick}
                        disabled={entry.disabled || (p.busy && p.selectionCount > 0)}
                      >
                        {!entry.iconOnly && entry.label}
                      </ToolbarButton>
                    </Tooltip>
                  )}
                </div>
              </OverflowItem>
            </React.Fragment>
          ))}
          <OverflowMenu entries={entries} p={p} />
        </Toolbar>
      </Overflow>
      <Dialog open={searchOpen} onOpenChange={(_, data) => setSearchOpen(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{t('search')}</DialogTitle>
            <DialogContent className={s.fields}>
              <Input
                aria-label={t('search')}
                value={p.search}
                onChange={(_, data) => p.onSearch(data.value)}
              />
            </DialogContent>
            <DialogActions>
              <Toolbar>
                <ToolbarButton appearance="primary" onClick={() => setSearchOpen(false)}>
                  {t('action.done')}
                </ToolbarButton>
              </Toolbar>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
}
