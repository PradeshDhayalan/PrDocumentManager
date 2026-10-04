import { Theme } from '@fluentui/react-components';
import { HostContext } from '../../control/DmsGrid/src/types/HostContext';
import resources from '../../control/DmsGrid/strings/DmsGrid.1033.resx?raw';
export const strings = new Map(
  Array.from(new DOMParser().parseFromString(resources, 'text/xml').querySelectorAll('data')).map(
    (element) => [
      element.getAttribute('name') || '',
      element.querySelector('value')?.textContent || '',
    ],
  ),
);
export function getString(key: string): string {
  return strings.get(key) || key;
}
export function createMockContext(theme: Theme, recordId: string): HostContext {
  return {
    recordId,
    entityName: 'account',
    clientUrl: window.location.origin,
    allocatedHeight: -1,
    theme,
    getString,
    showTitle: true,
  };
}

// The harness supplies only the PCF members consumed by this control. Production
// receives the complete platform context; this isolated cast never reaches PCF.
export function createMockPcfContext(
  host: HostContext,
): ComponentFramework.Context<import('../../control/DmsGrid/generated/ManifestTypes').IInputs> {
  const context = {
    mode: {
      allocatedHeight: host.allocatedHeight,
      allocatedWidth: 1440,
      contextInfo: { entityId: host.recordId, entityTypeName: host.entityName },
    },
    page: { getClientUrl: () => host.clientUrl },
    fluentDesignLanguage: { tokenTheme: host.theme },
    resources: { getString: host.getString },
    parameters: {
      documents: {
        records: {},
        sortedRecordIds: [],
        loading: false,
        paging: { setPageSize: (_size: number) => undefined },
        filtering: {
          setFilter: (_filter: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression) =>
            undefined,
        },
        refresh: () => undefined,
      },
      enableDragDrop: { raw: true },
      defaultView: { raw: '0' },
      pageSize: { raw: 50 },
      showTitle: { raw: host.showTitle },
      showDetailsPane: { raw: true },
      enableTelemetry: { raw: false },
    },
  };
  return context as unknown as ComponentFramework.Context<
    import('../../control/DmsGrid/generated/ManifestTypes').IInputs
  >;
}
