import { Theme } from '@fluentui/react-components';
export interface HostContext {
  recordId: string;
  entityName: string;
  clientUrl: string;
  allocatedHeight: number;
  theme?: Theme;
  getString: (key: string) => string;
  showTitle: boolean;
}
