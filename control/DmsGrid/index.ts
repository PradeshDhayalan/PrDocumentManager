import * as React from 'react';
import { IInputs, IOutputs } from './generated/ManifestTypes';
import { DmsGridHost } from './src/app/DmsGridHost';
import { RaiseCustomAction } from './src/services/customActions';
import { HostContext } from './src/types/HostContext';
type RuntimeContext = ComponentFramework.Context<IInputs> & {
  events?: { OnCustomAction?: RaiseCustomAction };
  page?: { getClientUrl(): string };
};
export class DmsGrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  public init(context: ComponentFramework.Context<IInputs>): void {
    // TODO(SPIKE-S1): validate dataset paging and filtering on real form subgrids.
    context.parameters.documents.paging.setPageSize(context.parameters.pageSize.raw || 10);
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
      customActionsJson: context.parameters.customActionsJson?.raw || undefined,
      gridConfigurationJson: context.parameters.gridConfigurationJson?.raw || undefined,
      raiseCustomAction: runtime.events?.OnCustomAction
        ? (request) => runtime.events!.OnCustomAction!(request)
        : undefined,
      recordId: mode.contextInfo?.entityId.replace(/[{}]/g, '') || '',
      entityName: mode.contextInfo?.entityTypeName || '',
      // TODO(SPIKE-S4): context.page is undocumented; validate the same-origin fallback.
      clientUrl: runtime.page?.getClientUrl() || window.location.origin,
      allocatedHeight: context.mode.allocatedHeight,
      getString: (key) => context.resources.getString(key),
      showTitle: context.parameters.showTitle.raw !== false,
      defaultView: String(context.parameters.defaultView.raw) === '1' ? 'Tiles' : 'List',
      pageSize: context.parameters.pageSize.raw || 10,
      enableDragDrop: context.parameters.enableDragDrop.raw !== false,
      showDetailsPane: context.parameters.showDetailsPane.raw !== false,
      userId: context.userSettings?.userId?.replace(/[{}]/g, ''),
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
