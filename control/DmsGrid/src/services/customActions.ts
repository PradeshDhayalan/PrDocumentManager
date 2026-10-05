import { DocumentRow, Entity } from './documents';
export interface CustomAction {
  id: string;
  label: string;
  handler: string;
  icon: string;
  order: number;
  selection: { min: number; max?: number };
  confirmMessage?: string;
  refreshOnSuccess: boolean;
}
export interface ActionResult {
  success: boolean;
  message?: string;
  refresh?: boolean;
}
export interface ActionRequest {
  requestId: string;
  actionId: string;
  handler: string;
  selectedRecordIds: string[];
  selectedRecords: { id: string; name: string; data: Entity }[];
  parent: { id: string; entityName: string };
  complete: (result: ActionResult) => void;
}
export type RaiseCustomAction = (request: ActionRequest) => void;
const icons = new Set(['Send', 'Checkmark', 'Document', 'Edit', 'Link', 'Open', 'Download']);
const record = (value: unknown): value is Record<string, unknown> =>
  Boolean(value && typeof value === 'object' && !Array.isArray(value));
export function parseCustomActions(json?: string): CustomAction[] {
  if (!json?.trim()) return [];
  const input: unknown = JSON.parse(json);
  if (!record(input) || !Array.isArray(input.actions) || input.actions.length > 30)
    throw new Error('customActionsJson must contain an actions array with at most 30 buttons.');
  const ids = new Set<string>();
  return input.actions
    .map((value: unknown, index: number): CustomAction => {
      if (!record(value)) throw new Error(`Custom action ${index + 1} must be an object.`);
      if (['code', 'javascript', 'js'].some((key) => key in value))
        throw new Error('Custom actions use registered handler names, not JavaScript source.');
      if (
        typeof value.id !== 'string' ||
        !/^[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(value.id) ||
        ids.has(value.id)
      )
        throw new Error('Custom action IDs must be unique identifiers.');
      ids.add(value.id);
      if (typeof value.label !== 'string' || !value.label.trim() || value.label.length > 100)
        throw new Error(`Custom action ${value.id} requires a label of up to 100 characters.`);
      if (
        typeof value.handler !== 'string' ||
        !/^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/.test(value.handler) ||
        value.handler
          .split('.')
          .some((part) => ['__proto__', 'prototype', 'constructor'].includes(part))
      )
        throw new Error(`Custom action ${value.id} requires a registered method name.`);
      const selection = value.selection === undefined ? {} : value.selection;
      if (!record(selection)) throw new Error(`Invalid selection rules for ${value.id}.`);
      const min = selection.min ?? 1,
        max = selection.max;
      if (
        typeof min !== 'number' ||
        !Number.isSafeInteger(min) ||
        min < 0 ||
        (max !== undefined && (typeof max !== 'number' || !Number.isSafeInteger(max) || max < min))
      )
        throw new Error(`Invalid selection limits for ${value.id}.`);
      if (value.icon !== undefined && (typeof value.icon !== 'string' || !icons.has(value.icon)))
        throw new Error(`Unsupported icon for ${value.id}.`);
      if (
        value.order !== undefined &&
        (typeof value.order !== 'number' || !Number.isFinite(value.order))
      )
        throw new Error(`Invalid order for ${value.id}.`);
      if (
        value.confirmMessage !== undefined &&
        (typeof value.confirmMessage !== 'string' || value.confirmMessage.length > 1000)
      )
        throw new Error(`Invalid confirmation message for ${value.id}.`);
      if (value.refreshOnSuccess !== undefined && typeof value.refreshOnSuccess !== 'boolean')
        throw new Error(`Invalid refreshOnSuccess for ${value.id}.`);
      return {
        id: value.id,
        label: value.label.trim(),
        handler: value.handler,
        icon: String(value.icon || 'Document'),
        order: Number(value.order ?? index),
        selection: { min, max: max as number | undefined },
        confirmMessage: value.confirmMessage as string | undefined,
        refreshOnSuccess: value.refreshOnSuccess === true,
      };
    })
    .sort((a, b) => a.order - b.order);
}
export function actionEnabled(action: CustomAction, count: number): boolean {
  return (
    count >= action.selection.min &&
    (action.selection.max === undefined || count <= action.selection.max)
  );
}
export function invokeCustomAction(
  action: CustomAction,
  rows: DocumentRow[],
  parent: ActionRequest['parent'],
  raise: RaiseCustomAction | undefined,
  signal: AbortSignal,
  timeoutMs = 30000,
): Promise<ActionResult> {
  if (!actionEnabled(action, rows.length))
    return Promise.reject(new Error('This action is not available for the current selection.'));
  if (!raise)
    return Promise.reject(
      new Error(
        'The host does not support OnCustomAction. Register the form handler in a compatible Power Apps environment.',
      ),
    );
  return new Promise((resolve, reject) => {
    let finished = false;
    const settle = (error?: Error, result?: ActionResult) => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      signal.removeEventListener('abort', abort);
      if (error) reject(error);
      else resolve(result!);
    };
    const abort = () => settle(new DOMException('Cancelled', 'AbortError'));
    const timer = setTimeout(
      () =>
        settle(
          new Error(
            'The custom action timed out. Check that its form handler is registered and calls complete().',
          ),
        ),
      timeoutMs,
    );
    signal.addEventListener('abort', abort, { once: true });
    if (signal.aborted) {
      abort();
      return;
    }
    const request: ActionRequest = {
      requestId: crypto.randomUUID(),
      actionId: action.id,
      handler: action.handler,
      selectedRecordIds: rows.map((row) => row.id),
      selectedRecords: rows.map((row) => ({ id: row.id, name: row.name, data: { ...row.raw } })),
      parent: { ...parent },
      complete: (result) => {
        if (
          !result ||
          typeof result.success !== 'boolean' ||
          (result.message !== undefined && typeof result.message !== 'string') ||
          (result.refresh !== undefined && typeof result.refresh !== 'boolean')
        ) {
          settle(new Error('The custom handler returned an invalid result.'));
          return;
        }
        if (!result.success) settle(new Error(result.message || 'The custom action failed.'));
        else settle(undefined, result);
      },
    };
    try {
      raise(request);
    } catch (error) {
      settle(error instanceof Error ? error : new Error('The custom event handler failed.'));
    }
  });
}
