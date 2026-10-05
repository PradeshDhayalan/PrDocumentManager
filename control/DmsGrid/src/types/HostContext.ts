import { RaiseCustomAction } from '../services/customActions';
import { Theme } from '@fluentui/react-components';
export interface HostContext {
  customActionsJson?: string;
  raiseCustomAction?: RaiseCustomAction;
  recordId: string;
  entityName: string;
  clientUrl: string;
  allocatedHeight: number;
  theme?: Theme;
  getString: (key: string) => string;
  showTitle: boolean;
  defaultView?: 'List' | 'Tiles';
  pageSize?: number;
  enableDragDrop?: boolean;
  showDetailsPane?: boolean;
  userId?: string;
}
