import * as React from 'react';
import { Theme } from '@fluentui/react-components';
import { IInputs, IOutputs } from './generated/ManifestTypes';
import { DmsGridHost } from './src/app/DmsGridHost';
import { HostContext } from './src/types/HostContext';
type RuntimeContext = ComponentFramework.Context<IInputs> & {
  fluentDesignLanguage?: { tokenTheme?: Theme };
  page?: { getClientUrl(): string };
};
export class DmsGrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  public init(context: ComponentFramework.Context<IInputs>): void {
    // TODO(SPIKE-S1): validate dataset paging and filtering on real form subgrids.
    context.parameters.documents.paging.setPageSize(1);
    const info = context.mode as ComponentFramework.Mode & {
      contextInfo?: { entityId: string; entityTypeName: string };
    };
    if (info.contextInfo?.entityId)
      context.parameters.documents.filtering.setFilter({
        filterOperator: 0,
        conditions: [
          {
            attributeName: 'dms_regardingid',
            conditionOperator: 0,
            value: info.contextInfo.entityId.replace(/[{}]/g, ''),
          },
          {
            attributeName: 'dms_regardingtype',
            conditionOperator: 0,
            value: info.contextInfo.entityTypeName,
          },
        ],
      });
  }
  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    const runtime = context as RuntimeContext;
    const mode = context.mode as ComponentFramework.Mode & {
      contextInfo?: { entityId: string; entityTypeName: string };
    };
    const host: HostContext = {
      recordId: mode.contextInfo?.entityId.replace(/[{}]/g, '') || '',
      entityName: mode.contextInfo?.entityTypeName || '',
      // TODO(SPIKE-S4): context.page is undocumented; validate the same-origin fallback.
      clientUrl: runtime.page?.getClientUrl() || window.location.origin,
      allocatedHeight: context.mode.allocatedHeight,
      theme: runtime.fluentDesignLanguage?.tokenTheme,
      getString: (key) => context.resources.getString(key),
      showTitle: context.parameters.showTitle.raw !== false,
    };
    return React.createElement(DmsGridHost, { host });
  }
  public getOutputs(): IOutputs {
    return {};
  }
  public destroy(): void {
    /* React host owns effect cleanup. */
  }
}
