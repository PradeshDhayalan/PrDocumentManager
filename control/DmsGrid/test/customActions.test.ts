import { webcrypto } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import {
  parseCustomActions,
  invokeCustomAction,
  actionEnabled,
  ActionRequest,
  RaiseCustomAction,
} from '../src/services/customActions';
import { mapDocument } from '../src/services/documents';
Object.defineProperty(globalThis, 'crypto', { value: webcrypto, configurable: true });
const definition = {
  id: 'review',
  label: 'Review',
  handler: 'Contoso.Documents.review',
  selection: { min: 1, max: 2 },
  refreshOnSuccess: true,
};
const action = parseCustomActions(JSON.stringify({ actions: [definition] }))[0];
const row = mapDocument({
  dms_documentid: '11111111-1111-4111-8111-111111111111',
  dms_name: 'Test.txt',
  dms_description: 'Original',
  '@odata.etag': 'W/"1"',
});
const parent = { id: '22222222-2222-4222-8222-222222222222', entityName: 'account' };
test('validates custom button configuration, identities and selection rules', () => {
  expect(parseCustomActions('')).toEqual([]);
  expect(actionEnabled(action, 0)).toBe(false);
  expect(actionEnabled(action, 2)).toBe(true);
  expect(actionEnabled(action, 3)).toBe(false);
  for (const changes of [
    { handler: 'Contoso.__proto__.run' },
    { handler: 'alert(1)' },
    { code: 'alert(1)' },
    { selection: { min: 2, max: 1 } },
    { icon: 'Unknown' },
    { refreshOnSuccess: 'yes' },
  ])
    expect(() =>
      parseCustomActions(JSON.stringify({ actions: [{ ...definition, ...changes }] })),
    ).toThrow();
  expect(() => parseCustomActions(JSON.stringify({ actions: [definition, definition] }))).toThrow(
    /unique/,
  );
});
test('sends an isolated selection snapshot and resolves exactly once through complete', async () => {
  let request: ActionRequest | undefined;
  const result = invokeCustomAction(
    action,
    [row],
    parent,
    (event) => {
      request = event;
      event.selectedRecords[0].data.dms_description = 'Changed outside the control';
      event.complete({ success: true, message: 'Reviewed', refresh: true });
      event.complete({ success: false, message: 'Late error' });
    },
    new AbortController().signal,
  );
  await expect(result).resolves.toEqual({ success: true, message: 'Reviewed', refresh: true });
  expect(request?.selectedRecordIds).toEqual([row.id]);
  expect(request?.parent).toEqual(parent);
  expect(row.raw.dms_description).toBe('Original');
});
test('reports missing host support, handler errors, timeouts and cancellation', async () => {
  await expect(
    invokeCustomAction(action, [row], parent, undefined, new AbortController().signal),
  ).rejects.toThrow('does not support');
  await expect(
    invokeCustomAction(
      action,
      [row],
      parent,
      (event) => event.complete({ success: false, message: 'Permission denied' }),
      new AbortController().signal,
    ),
  ).rejects.toThrow('Permission denied');
  jest.useFakeTimers();
  const timed = invokeCustomAction(
    action,
    [row],
    parent,
    () => undefined,
    new AbortController().signal,
    50,
  );
  const assertion = expect(timed).rejects.toThrow('timed out');
  jest.advanceTimersByTime(51);
  await assertion;
  const controller = new AbortController();
  let late: ActionRequest | undefined;
  const cancelled = invokeCustomAction(
    action,
    [row],
    parent,
    (event) => {
      late = event;
    },
    controller.signal,
  );
  const cancelledAssertion = expect(cancelled).rejects.toMatchObject({ name: 'AbortError' });
  controller.abort();
  await cancelledAssertion;
  late?.complete({ success: true });
  jest.useRealTimers();
});
test('the example web resource dispatches named handlers and registers once per form control', async () => {
  const sandbox: {
    Contoso?: {
      Documents: {
        register(name: string, handler: (request: ActionRequest) => unknown): void;
        createDispatcher(host: object): RaiseCustomAction;
        onLoad(context: object): void;
      };
    };
  } = {};
  vm.runInNewContext(
    fs.readFileSync(path.resolve('examples/custom-actions.webresource.js'), 'utf8'),
    sandbox,
  );
  const registry = sandbox.Contoso!.Documents;
  const addEventHandler = jest.fn();
  const control = { addEventHandler };
  const form = { getControl: () => control };
  const executionContext = { getFormContext: () => form };
  // The example reads Xrm only from the host form's script, never from the PCF.
  Object.assign(sandbox, { Xrm: { WebApi: {} } });
  registry.onLoad(executionContext);
  registry.onLoad(executionContext);
  expect(addEventHandler).toHaveBeenCalledTimes(1);
  expect(addEventHandler.mock.calls[0][0]).toBe('OnCustomAction');
  const raise = registry.createDispatcher({});
  await expect(
    invokeCustomAction(
      { ...action, handler: 'Contoso.Documents.showSelection' },
      [row],
      parent,
      raise,
      new AbortController().signal,
    ),
  ).resolves.toMatchObject({ success: true, message: 'Test.txt', refresh: false });
  await expect(
    invokeCustomAction(action, [row], parent, raise, new AbortController().signal),
  ).rejects.toThrow('No registered handler');
});
