import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { DocumentRegular } from '@fluentui/react-icons';
import { HostContext } from '../types/HostContext';
import { createI18n } from '../services/i18n';
const useStyles = makeStyles({
  root: {
    minHeight: '420px',
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  heading: { padding: '24px', fontSize: '24px', fontWeight: tokens.fontWeightSemibold },
  empty: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '16px',
    minHeight: '320px',
    color: tokens.colorNeutralForeground2,
  },
  icon: { fontSize: '48px', color: tokens.colorNeutralForeground3 },
});
function Canvas({ host }: { host: HostContext }) {
  const s = useStyles(),
    t = createI18n(host.getString);
  return (
    <section
      className={s.root}
      aria-label={t('documents')}
      style={host.allocatedHeight > 0 ? { height: host.allocatedHeight } : undefined}
    >
      {host.showTitle && <div className={s.heading}>{t('documents')}</div>}
      <div className={s.empty}>
        <DocumentRegular className={s.icon} />
        <Text weight="semibold">{t(host.recordId ? 'empty.title' : 'unsaved.title')}</Text>
        {host.recordId && <Text>{t('empty.description')}</Text>}
      </div>
    </section>
  );
}
export function DmsGridHost({ host }: { host: HostContext }) {
  return (
    <FluentProvider theme={host.theme || webLightTheme}>
      <Canvas host={host} />
    </FluentProvider>
  );
}
