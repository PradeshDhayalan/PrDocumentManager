import { customActionsJson, raiseCustomAction } from './customActionExample';
import './fonts.css';
import { microsoftFontFamily } from '../../control/DmsGrid/src/services/fonts';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { FluentProvider, webLightTheme, makeStyles, tokens } from '@fluentui/react-components';
import { DmsGrid } from '../../control/DmsGrid';
import { HostContext } from '../../control/DmsGrid/src/types/HostContext';
import { createMockContext, createMockPcfContext } from './createMockContext';
const useStyles = makeStyles({
  page: {
    minHeight: '100vh',
    fontFamily: tokens.fontFamilyBase,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  control: { width: '100%', margin: '0 auto' },
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
  const s = useStyles();
  // Sample records remain addressable for automated tests without harness chrome.
  const samples: Record<string, { id: string; entity: string }> = {
    contoso: { id: '11111111-1111-4111-8111-111111111111', entity: 'account' },
    fabrikam: { id: '22222222-2222-4222-8222-222222222222', entity: 'account' },
    empty: { id: '33333333-3333-4333-8333-333333333333', entity: 'account' },
    legacy: { id: '44444444-4444-4444-8444-444444444444', entity: 'opportunity' },
    unsaved: { id: '', entity: 'account' },
  };
  const sample =
    samples[new URLSearchParams(window.location.search).get('record') || 'contoso'] ||
    samples.contoso;
  return (
    <FluentProvider theme={{ ...webLightTheme, fontFamilyBase: microsoftFontFamily }}>
      <div className={s.page}>
        <main className={s.control}>
          <PcfMount
            host={{
              ...createMockContext(webLightTheme, sample.id, sample.entity),
              customActionsJson,
              raiseCustomAction,
            }}
          />
        </main>
      </div>
    </FluentProvider>
  );
}
ReactDOM.render(<App />, document.getElementById('root'));
