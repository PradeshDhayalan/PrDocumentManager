# Custom actions inside the PCF

The control exposes an optional text input, `customActionsJson`. Supply the contents of [custom-actions.json](../examples/custom-actions.json) when configuring the PCF. Empty configuration leaves the built-in commands available. The local harness already supplies the example configuration.

```json
{
  "actions": [
    {
      "id": "markReviewed",
      "label": "Mark as reviewed",
      "icon": "Checkmark",
      "selection": { "min": 1, "max": 20 },
      "handler": "Contoso.Documents.markReviewed",
      "confirmMessage": "Mark the selected documents as reviewed?",
      "refreshOnSuccess": true
    }
  ]
}
```

## Try the local example

Open http://localhost:5173, select one or more documents and click **Mark as reviewed**. Confirm the dialog. The example writes `[Reviewed]` at the beginning of each selected document's Description and refreshes the grid. Open Details to inspect the saved description. **Show selection** displays the selected document names without changing records. These examples use the same JavaScript dispatcher and handlers as the form web resource.

## Configure a model-driven form

1. Add [custom-actions.webresource.js](../examples/custom-actions.webresource.js) as a JavaScript web resource and include it in the form libraries.
2. Replace `DocumentsGrid` in its `getControl()` call with the actual form control name.
3. Add `Contoso.Documents.onLoad` to form OnLoad, with **Pass execution context as first parameter** enabled.
4. Configure `customActionsJson` on the PCF with your JSON, then save and publish the form.

The OnLoad handler subscribes with `control.addEventHandler('OnCustomAction', dispatcher)`. It obtains the form context from the form event and passes it to registered handlers. The PCF does not obtain a form context through globals or evaluate JavaScript source in JSON.

Register another handler in the same library, or a library loaded afterwards:

```js
Contoso.Documents.register('Contoso.Documents.myAction', async function (request, host) {
  // host.formContext is available on the model-driven form.
  // host.webApi is Xrm.WebApi there, and a mock adapter in the local harness.
  const names = request.selectedRecords.map(record => record.name);
  return { message: names.join(', '), refresh: false };
});
```

Set the button's `handler` to `Contoso.Documents.myAction`. Names are registry keys: a matching global function alone does not register it. Handler promises resolve to optional `message` and `refresh` fields; a thrown error or rejected promise displays an error in the control. The dispatcher calls the completion callback on your behalf. The review example updates records individually; failures may leave earlier updates saved and report the count. Implement transactions and concurrency requirements in your own business handler.

## Configuration and event contract

At most 30 buttons are accepted. IDs must be unique identifiers; labels are required and limited to 100 characters. `order` sorts custom buttons. Supported icon names are `Send`, `Checkmark`, `Document`, `Edit`, `Link`, `Open` and `Download`. `selection.min` defaults to 1; `selection.max` is optional. A minimum of 0 allows the button to appear with no selected records. Commands outside their selection limits are disabled, and commands share the control's busy state. On narrow screens they appear in More commands.

`confirmMessage` requests confirmation before dispatch. `refreshOnSuccess` defaults to false; a handler's explicit `refresh` overrides it. Invalid JSON displays an error without disabling built-in commands. Inline `code`, `javascript` and `js` configuration fields are rejected.

The `OnCustomAction` payload contains `requestId`, `actionId`, `handler`, `selectedRecordIds`, `selectedRecords` (each with `id`, `name` and a `data` snapshot), `parent` (`id`, `entityName`) and `complete(result)`. Direct event handlers must complete with `{ success: true, message?, refresh? }` or `{ success: false, message }`. Completion is accepted once. The control times out after 30 seconds and ignores late completions after cancellation; this does not roll back or cancel external work already started by the handler.

Microsoft currently documents PCF custom events as a preview feature. The local harness simulates this event surface and verifies the example; registration on a real model-driven form still requires tenant verification. See Microsoft's [custom event guide](https://learn.microsoft.com/en-us/power-apps/developer/component-framework/events) and [addEventHandler API](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/controls/addeventhandler).
