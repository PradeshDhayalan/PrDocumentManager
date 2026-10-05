import * as React from 'react';
import { Persona, Badge, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { GlobeRegular, WarningRegular } from '@fluentui/react-icons';
import { DocumentRow, T } from '../types/Documents';
import { DocumentRepository } from '../services/DocumentRepository';
import { fileKind } from '../services/format';
import { resolveProvider, usable } from '../providers/registry';
import { fileIcons } from './iconSource';
const useStyles = makeStyles({
  icon: { width: '22px', height: '26px', flexShrink: 0, color: tokens.colorBrandForeground1 },
  scene: {
    height: '100%',
    width: '100%',
    backgroundColor: tokens.colorNeutralBackground3,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    overflow: 'hidden',
    position: 'relative',
  },
  mini: {
    width: '46px',
    height: '34px',
    flexShrink: 0,
    borderRadius: tokens.borderRadiusSmall,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  image: { width: '100%', height: '100%', objectFit: 'cover' },
  miniScale: {
    position: 'absolute',
    width: '192px',
    height: '142px',
    top: 0,
    left: 0,
    transform: 'scale(.24)',
    transformOrigin: 'top left',
    display: 'grid',
    placeItems: 'center',
  },
  paper: {
    height: '136px',
    width: '104px',
    padding: '12px',
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
    overflow: 'hidden',
  },
  title: {
    fontSize: '9px',
    lineHeight: '12px',
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: '10px',
    overflow: 'hidden',
    maxHeight: '26px',
  },
  line: { height: '2px', marginBottom: '6px', backgroundColor: tokens.colorNeutralStroke2 },
  rule: {
    height: '3px',
    width: '24px',
    backgroundColor: tokens.colorBrandForeground1,
    marginBottom: '10px',
  },
  sheet: {
    width: '174px',
    height: '104px',
    boxShadow: tokens.shadow4,
    backgroundColor: tokens.colorNeutralBackground1,
    overflow: 'hidden',
  },
  sheetHead: {
    height: '24px',
    padding: '4px 8px',
    fontSize: '9px',
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground2,
  },
  cells: { display: 'grid', gridTemplateColumns: 'repeat(4,1fr)' },
  cell: {
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    height: '14px',
  },
  slide: {
    width: '178px',
    height: '104px',
    padding: '14px',
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
  },
  slideTitle: {
    color: tokens.colorBrandForeground1,
    fontSize: '15px',
    lineHeight: '18px',
    fontWeight: tokens.fontWeightSemibold,
    maxHeight: '56px',
    overflow: 'hidden',
  },
  slideRule: {
    marginTop: '12px',
    height: '7px',
    width: '80px',
    backgroundColor: tokens.colorBrandBackground2,
  },
  status: { display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' },
  warning: { color: tokens.colorPaletteRedForeground1 },
  person: {
    minWidth: 0,
    '& .fui-Persona__primaryText': {
      fontSize: '12px',
      overflow: 'hidden',
      textOverflow: 'ellipsis',
      whiteSpace: 'nowrap',
    },
  },
});
export function FileTypeIcon({ row }: { row: DocumentRow }) {
  const s = useStyles(),
    kind = fileKind(row);
  return kind === 'link' ? (
    <GlobeRegular className={s.icon} />
  ) : (
    <img alt="" className={s.icon} src={fileIcons[kind] || fileIcons.generic} />
  );
}
export function DocumentPerson({
  name,
  activity = false,
  secondary,
}: {
  name: string;
  activity?: boolean;
  secondary?: string;
}) {
  const s = useStyles();
  return (
    <Persona
      className={s.person}
      name={name}
      secondaryText={secondary}
      size={activity ? 'medium' : 'extra-small'}
      avatar={{ color: 'colorful' }}
    />
  );
}
export function StatusBadge({ row, t }: { row: DocumentRow; t: T }) {
  const failed = row.uploadState === 100000002,
    pending = row.uploadState === 100000000;
  return (
    <Badge
      appearance="tint"
      color={
        failed
          ? 'danger'
          : pending
            ? 'warning'
            : row.status === 100000001
              ? 'success'
              : row.status === 100000002
                ? 'subtle'
                : 'subtle'
      }
    >
      {t(
        failed ? 'upload.failed' : pending ? 'upload.pending' : `status.${row.status ?? 100000000}`,
      )}
    </Badge>
  );
}
export function ExpiryBadge({ row, t }: { row: DocumentRow; t: T }) {
  if (!row.expiry) return null;
  const days = Math.ceil((Date.parse(row.expiry + 'T23:59:59') - Date.now()) / 86400000);
  return days < 0 ? (
    <Badge appearance="tint" size="small" color="danger">
      {t('expiry.expired')}
    </Badge>
  ) : days <= 30 ? (
    <Badge appearance="tint" size="small" color="warning">
      {t('expiry.soon', { days })}
    </Badge>
  ) : null;
}
export function DocumentThumbnail({
  row,
  repo,
  t,
  mini = false,
}: {
  row: DocumentRow;
  repo: DocumentRepository;
  t: T;
  mini?: boolean;
}) {
  const s = useStyles(),
    kind = fileKind(row),
    [image, setImage] = React.useState('');
  React.useEffect(() => {
    setImage('');
    if (kind !== 'image' || row.sizeKb > 4096) return;
    const provider = resolveProvider(row.provider, repo.client);
    if (!usable(row, provider) || !provider.capabilities.ownsBinary) return;
    const controller = new AbortController();
    let url = '';
    void provider
      .getContent(row, undefined, controller.signal)
      .then((content) => {
        if (content.kind === 'blob' && !controller.signal.aborted) {
          url = URL.createObjectURL(content.blob);
          setImage(url);
        }
      })
      .catch(() => undefined);
    return () => {
      controller.abort();
      if (url) URL.revokeObjectURL(url);
    };
  }, [row.id, row.etag, kind, repo]);
  let content: React.ReactNode;
  if (row.uploadState !== 100000001)
    content = <WarningRegular className={s.warning} fontSize={mini ? 24 : 40} />;
  else if (image)
    return (
      <div className={mergeClasses(s.scene, mini && s.mini)}>
        <img
          className={s.image}
          src={image}
          alt={t('thumbnail', { name: row.name })}
          loading="lazy"
          onError={() => setImage('')}
        />
      </div>
    );
  else if (kind === 'excel')
    content = (
      <div className={s.sheet}>
        <div className={s.sheetHead}>{row.name}</div>
        <div className={s.cells} aria-hidden="true">
          {Array.from({ length: 24 }, (_, i) => (
            <div className={s.cell} key={i} />
          ))}
        </div>
      </div>
    );
  else if (kind === 'powerpoint')
    content = (
      <div className={s.slide}>
        <div className={s.slideTitle}>{row.name.replace(/\.[^.]+$/, '')}</div>
        <div className={s.slideRule} />
      </div>
    );
  else if (['word', 'pdf', 'text', 'link'].includes(kind))
    content = (
      <div className={s.paper}>
        <div className={s.rule} />
        <div className={s.title}>{row.name.replace(/\.[^.]+$/, '')}</div>
        {Array.from({ length: 11 }, (_, i) => (
          <div className={s.line} key={i} />
        ))}
      </div>
    );
  else content = <FileTypeIcon row={row} />;
  return (
    <div
      className={mergeClasses(s.scene, mini && s.mini)}
      role="img"
      aria-label={t('thumbnail', { name: row.name })}
    >
      {mini ? <div className={s.miniScale}>{content}</div> : content}
    </div>
  );
}
