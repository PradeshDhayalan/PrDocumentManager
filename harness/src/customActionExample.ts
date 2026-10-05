import config from '../../examples/custom-actions.json';
// The same JavaScript web resource is used locally and on the model-driven form.
import '../../examples/custom-actions.webresource.js';
import { RaiseCustomAction } from '../../control/DmsGrid/src/services/customActions';
interface ExampleRegistry {
  createDispatcher(host: {
    webApi: {
      updateRecord(entity: string, id: string, data: Record<string, unknown>): Promise<void>;
    };
  }): RaiseCustomAction;
}
const registry = (globalThis as typeof globalThis & { Contoso: { Documents: ExampleRegistry } })
  .Contoso.Documents;
export const customActionsJson = JSON.stringify(config);
export const raiseCustomAction = registry.createDispatcher({
  webApi: {
    async updateRecord(entity, id, data) {
      if (entity !== 'dms_document') throw new Error('The local example only updates documents.');
      const response = await fetch(`/api/data/v9.2/dms_documents(${id})`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json', 'If-Match': '*' },
        body: JSON.stringify(data),
      });
      if (!response.ok) {
        const result = await response.json();
        throw new Error(result.error?.message || 'Example update failed.');
      }
    },
  },
});
