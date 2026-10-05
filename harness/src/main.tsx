import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  teamsHighContrastTheme,
  createLightTheme,
  Dropdown,
  Option,
  Toolbar,
  ToolbarButton,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { DmsGrid } from '../../control/DmsGrid';
import { HostContext } from '../../control/DmsGrid/src/types/HostContext';
import { createMockContext, createMockPcfContext, getString } from './createMockContext';
const teal = createLightTheme({
  10: '#001d1e',
  20: '#003034',
  30: '#004549',
  40: '#00575c',
  50: '#006167',
  60: '#006c72',
  70: '#00777e',
  80: '#008489',
  90: '#149399',
  100: '#2ba4a9',
  110: '#5cd0d6',
  120: '#7edce0',
  130: '#a1e6e9',
  140: '#beeef0',
  150: '#d9f5f6',
  160: '#eefbfb',
});
const themes = {
  blue: webLightTheme,
  teal,
  dark: webDarkTheme,
  highContrast: teamsHighContrastTheme,
};
const useStyles = makeStyles({
  page: { minHeight: '100vh', padding: '20px', backgroundColor: tokens.colorNeutralBackground2 },
  toolbar: {
    display: 'flex',
    gap: '8px',
    alignItems: 'center',
    flexWrap: 'wrap',
    marginBottom: '24px',
  },
  card: {
    maxWidth: '1440px',
    margin: '0 auto',
    boxShadow: tokens.shadow4,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
  },
  label: { color: tokens.colorNeutralForeground2, fontSize: '12px' },
});
function PcfMount({ host }: { host: HostContext }) {
  const context = createMockPcfContext(host);
  const control = React.useMemo(() => {
    const instance = new DmsGrid();
    instance.init(context);
    return instance;
  }, [host.recordId]);
  React.useEffect(() => () => control.destroy(), [control]);
  return control.updateView(context);
}
function App() {
  const s = useStyles(),
    [theme, setTheme] = React.useState<keyof typeof themes>('blue'),
    [record, setRecord] = React.useState('contoso'),
    [provider, setProvider] = React.useState('Note');
  const records: Record<string, { id: string; entity: string }> = {
    contoso: { id: '11111111-1111-4111-8111-111111111111', entity: 'account' },
    fabrikam: { id: '22222222-2222-4222-8222-222222222222', entity: 'account' },
    empty: { id: '33333333-3333-4333-8333-333333333333', entity: 'contact' },
    legacy: { id: '44444444-4444-4444-8444-444444444444', entity: 'opportunity' },
    unsaved: { id: '', entity: 'account' },
  };
  const changeProvider = async (value: string) => {
    const response = await fetch('/__mock/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ activeProvider: value }),
    });
    if (response.ok) {
      setProvider(value);
      window.dispatchEvent(new Event('dms-config-changed'));
    }
  };
  React.useEffect(() => {
    fetch('/__mock/state')
      .then((response) => response.json())
      .then((value: { config: { activeProvider: string } }) =>
        setProvider(value.config.activeProvider),
      )
      .catch(() => undefined);
  }, []);
  return (
    <FluentProvider theme={themes[theme]}>
      <div className={s.page}>
        <Toolbar className={s.toolbar} aria-label={getString('harness.title')}>
          <Text weight="semibold">{getString('harness.title')}</Text>
          <Text className={s.label}>{getString('harness.theme')}</Text>
          {(Object.keys(themes) as (keyof typeof themes)[]).map((name) => (
            <ToolbarButton
              key={name}
              appearance={name === theme ? 'primary' : 'subtle'}
              onClick={() => setTheme(name)}
            >
              {getString(`theme.${name}`)}
            </ToolbarButton>
          ))}
          <Dropdown
            aria-label={getString('harness.record')}
            value={getString('record.' + record)}
            selectedOptions={[record]}
            onOptionSelect={(_, data) => setRecord(data.optionValue || 'contoso')}
          >
            {Object.keys(records).map((key) => (
              <Option key={key} value={key}>
                {getString('record.' + key)}
              </Option>
            ))}
          </Dropdown>
          <Dropdown
            aria-label={getString('harness.provider')}
            value={getString('provider.' + provider)}
            selectedOptions={[provider]}
            onOptionSelect={(_, data) => void changeProvider(data.optionValue || 'Note')}
          >
            <Option value="Note">{getString('provider.Note')}</Option>
            <Option value="SharePoint">{getString('provider.SharePoint')}</Option>
          </Dropdown>
        </Toolbar>
        <main className={s.card}>
          <PcfMount
            host={createMockContext(themes[theme], records[record].id, records[record].entity)}
          />
        </main>
      </div>
    </FluentProvider>
  );
}
ReactDOM.render(<App />, document.getElementById('root'));
