/* Upload this file as a model-driven form JavaScript web resource.
   Add Contoso.Documents.onLoad to form OnLoad and pass executionContext.
   Replace DocumentsGrid below with the configured form control name. */
var Contoso = globalThis.Contoso || {};
(function (namespace) {
  'use strict';
  const handlers = Object.create(null);
  const registrations = new WeakMap();
  namespace.Documents = namespace.Documents || {};
  const documents = namespace.Documents;
  documents.register = function (name, handler) {
    if (typeof name !== 'string' || typeof handler !== 'function')
      throw new Error('A handler name and function are required.');
    handlers[name] = handler;
  };
  documents.createDispatcher = function (host) {
    return function (request) {
      const handler = handlers[request.handler];
      if (!handler) {
        request.complete({ success: false, message: 'No registered handler: ' + request.handler });
        return;
      }
      Promise.resolve()
        .then(() => handler(request, host))
        .then(
          (result) => request.complete({ success: true, ...(result || {}) }),
          (error) =>
            request.complete({
              success: false,
              message: error instanceof Error ? error.message : String(error),
            }),
        );
    };
  };
  documents.onLoad = function (executionContext) {
    const formContext = executionContext.getFormContext();
    const control = formContext.getControl('DocumentsGrid');
    if (!control || typeof control.addEventHandler !== 'function')
      throw new Error(
        'DocumentsGrid must support PCF custom events. Check the control name and environment.',
      );
    // Repeated form OnLoad must not register duplicate handlers.
    if (registrations.has(control)) return;
    const dispatcher = documents.createDispatcher({ formContext, webApi: Xrm.WebApi });
    control.addEventHandler('OnCustomAction', dispatcher);
    registrations.set(control, dispatcher);
  };
  documents.register('Contoso.Documents.showSelection', function (request) {
    return {
      message: request.selectedRecords.map((record) => record.name).join(', '),
      refresh: false,
    };
  });
  // A working example: persist a review marker in the Description field.
  // Supply the real review/approval workflow here for your business process.
  documents.register('Contoso.Documents.markReviewed', async function (request, host) {
    let updated = 0;
    try {
      for (const record of request.selectedRecords) {
        const existing = String(record.data.dms_description || '');
        const description = existing.startsWith('[Reviewed]') ? existing : '[Reviewed] ' + existing;
        if (description.length > 4000)
          throw new Error('Description exceeds the 4,000-character limit: ' + record.name);
        await host.webApi.updateRecord('dms_document', record.id, { dms_description: description });
        updated++;
      }
    } catch (error) {
      throw new Error(
        updated +
          ' document(s) updated. ' +
          (error instanceof Error ? error.message : String(error)) +
          ' Refresh before retrying.',
      );
    }
    return { message: updated + ' document(s) marked as reviewed.', refresh: true };
  });
})(Contoso);
globalThis.Contoso = Contoso;
