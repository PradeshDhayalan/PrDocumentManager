import * as React from 'react';
import {
  InlineDrawer,
  Text,
  Toolbar,
  ToolbarButton,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  DismissRegular,
  EditRegular,
  OpenRegular,
  ArrowDownloadRegular,
  DocumentRegular,
} from '@fluentui/react-icons';
import { DocumentRow, T } from '../types/Documents';
import { DocumentRepository } from '../services/DocumentRepository';
import { resolveProvider, usable } from '../providers/registry';
import { DocumentThumbnail, DocumentPerson, StatusBadge } from './FileVisuals';
import { formatDate, formatSize } from '../services/format';
import { Action } from './DocumentCommandBar';
const useStyles = makeStyles({
  pane: {
    width: '300px',
    flexShrink: 0,
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
    padding: '18px 22px',
    boxSizing: 'border-box',
    overflowY: 'auto',
    '@media(max-width:1000px)': {
      position: 'fixed',
      top: 0,
      right: 0,
      bottom: 0,
      width: '320px',
      maxWidth: '100vw',
      zIndex: 40,
      backgroundColor: tokens.colorNeutralBackground1,
      boxShadow: tokens.shadow64,
    },
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '24px',
  },
  hero: { height: '170px', overflow: 'hidden', borderRadius: tokens.borderRadiusMedium },
  name: {
    fontSize: '17px',
    lineHeight: '22px',
    fontWeight: tokens.fontWeightSemibold,
    overflowWrap: 'anywhere',
    marginTop: '14px',
    marginBottom: '8px',
  },
  secondary: { fontSize: '12px', color: tokens.colorNeutralForeground2 },
  section: { padding: '20px 0', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: '14px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '16px',
  },
  props: {
    display: 'grid',
    gridTemplateColumns: '90px 1fr',
    gap: '14px 10px',
    fontSize: '12px',
    overflowWrap: 'anywhere',
  },
  hint: {
    padding: '12px',
    marginTop: '14px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    fontSize: '12px',
    lineHeight: '18px',
  },
  empty: {
    padding: '40px 0',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '16px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground2,
  },
  person: { marginBottom: '14px' },
  actions: { padding: '12px 0 0', gap: '6px' },
});
export function DetailsPane({
  rows,
  repo,
  t,
  onClose,
  onAction,
}: {
  rows: DocumentRow[];
  repo: DocumentRepository;
  t: T;
  onClose: () => void;
  onAction: (action: Action) => void;
}) {
  const s = useStyles(),
    row = rows.length === 1 ? rows[0] : undefined,
    provider = row ? resolveProvider(row.provider, repo.client) : undefined,
    can = row && provider && usable(row, provider);
  return (
    <InlineDrawer open position="end" className={s.pane} aria-label={t('details.title')}>
      <div className={s.header}>
        <Text weight="semibold">{t('details.title')}</Text>
        <Toolbar>
          <ToolbarButton
            appearance="subtle"
            icon={<DismissRegular />}
            aria-label={t('details.close')}
            onClick={onClose}
          />
        </Toolbar>
      </div>
      {row ? (
        <>
          <div className={s.hero}>
            <DocumentThumbnail row={row} repo={repo} t={t} />
          </div>
          <div className={s.name}>{row.name}</div>
          <Text className={s.secondary}>
            {row.sizeKb ? formatSize(row.sizeKb, t) : t('link.reference')}
          </Text>
          <Toolbar className={s.actions}>
            {can && (
              <ToolbarButton icon={<OpenRegular />} onClick={() => onAction('open')}>
                {t('action.open')}
              </ToolbarButton>
            )}
            {can && provider?.capabilities.canDownload && (
              <ToolbarButton icon={<ArrowDownloadRegular />} onClick={() => onAction('download')}>
                {t('action.download')}
              </ToolbarButton>
            )}
          </Toolbar>
          <div className={s.section}>
            <div className={s.sectionTitle}>
              {t('properties.title')}
              <Toolbar>
                <ToolbarButton
                  appearance="subtle"
                  icon={<EditRegular />}
                  aria-label={t('properties.edit')}
                  onClick={() => onAction('edit')}
                />
              </Toolbar>
            </div>
            <div className={s.props}>
              <Text className={s.secondary}>{t('column.dms_documenttype')}</Text>
              <span>{t('type.' + (row.documentType ?? 100000005))}</span>
              <Text className={s.secondary}>{t('column.dms_documentstatus')}</Text>
              <span>
                <StatusBadge row={row} t={t} />
              </span>
              <Text className={s.secondary}>{t('column.dms_expirydate')}</Text>
              <span>{formatDate(row.expiry, t)}</span>
              <Text className={s.secondary}>{t('description')}</Text>
              <span>{row.description || t('empty.value')}</span>
            </div>
          </div>
          <div className={s.section}>
            <div className={s.sectionTitle}>{t('activity.title')}</div>
            <div className={s.person}>
              <DocumentPerson
                name={row.modifiedBy}
                secondary={t('column._modifiedby_value')}
                activity
              />
            </div>
            <div className={s.person}>
              <DocumentPerson name={row.createdBy} secondary={t('createdBy')} activity />
            </div>
            <div className={s.props}>
              <Text className={s.secondary}>{t('column.modifiedon')}</Text>
              <span>{formatDate(row.modified, t)}</span>
              <Text className={s.secondary}>{t('created')}</Text>
              <span>{formatDate(row.created, t)}</span>
            </div>
          </div>
          <div className={s.section}>
            <div className={s.sectionTitle}>{t('storage.title')}</div>
            <div className={s.props}>
              <Text className={s.secondary}>{t('storage.source')}</Text>
              <span>{t('provider.' + row.provider)}</span>
              <Text className={s.secondary}>{t('column.dms_filesizekb')}</Text>
              <span>{row.sizeKb ? formatSize(row.sizeKb, t) : t('empty.value')}</span>
            </div>
            {provider && !provider.capabilities.ownsBinary && provider.capabilities.canOpen && (
              <div className={s.hint}>{t('link.permissions')}</div>
            )}
          </div>
        </>
      ) : (
        <div className={s.empty}>
          <DocumentRegular fontSize={44} />
          <Text>
            {rows.length ? t('selection.count', { count: rows.length }) : t('details.select')}
          </Text>
          {rows.length > 1 && (
            <Toolbar>
              <ToolbarButton icon={<EditRegular />} onClick={() => onAction('edit')}>
                {t('action.edit')}
              </ToolbarButton>
            </Toolbar>
          )}
        </div>
      )}
    </InlineDrawer>
  );
}
