import * as React from 'react';
import {
  Button,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Field,
  Input,
  Dropdown,
  Option,
  Checkbox,
} from '@fluentui/react-components';
import { Attribute, DocumentClient } from '../../control/DmsGrid/src/services/documents';
import { HostContext } from '../../control/DmsGrid/src/types/HostContext';
import {
  GridConfiguration,
  parseGridConfiguration,
} from '../../control/DmsGrid/src/services/gridConfiguration';
export interface LocalConfiguration {
  pageSize: number;
  grid: GridConfiguration;
}
export const defaultConfiguration: LocalConfiguration = {
  pageSize: 10,
  grid: parseGridConfiguration(),
};
export function PcfConfiguration({
  host,
  value,
  onApply,
}: {
  host: HostContext;
  value: LocalConfiguration;
  onApply: (value: LocalConfiguration) => void;
}) {
  const [open, setOpen] = React.useState(false);
  const [draft, setDraft] = React.useState(value);
  const [attributes, setAttributes] = React.useState<Attribute[]>([]);
  const [name, setName] = React.useState('');
  const [label, setLabel] = React.useState('');
  const [type, setType] = React.useState('String');
  const [error, setError] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const load = async () => {
    const client = new DocumentClient(host);
    const [metadata, config] = await Promise.all([client.attributes(), client.config()]);
    setAttributes(metadata);
    setDraft({
      ...value,
      grid: {
        ...value.grid,
        visibleColumns: value.grid.visibleColumns ?? config.entity.visibleColumns,
        editableColumns: value.grid.editableColumns ?? config.entity.editableColumns,
      },
    });
  };
  const patch = (grid: Partial<GridConfiguration>) =>
    setDraft((previous) => ({ ...previous, grid: { ...previous.grid, ...grid } }));
  return (
    <>
      <div style={{ padding: '8px 24px', textAlign: 'right' }}>
        <Button
          onClick={() => {
            setOpen(true);
            setError('');
            void load().catch((e) => setError(String(e)));
          }}
        >
          PCF Configuration
        </Button>
      </div>
      <Dialog
        open={open}
        onOpenChange={(_, data) => {
          if (!busy) setOpen(data.open);
        }}
      >
        <DialogSurface style={{ maxWidth: 850 }}>
          <DialogBody>
            <DialogTitle>PCF Configuration</DialogTitle>
            <DialogContent style={{ maxHeight: '70vh', overflowY: 'auto' }}>
              {error && <p role="alert">{error}</p>}
              <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
                <Field label="Documents per page (1–250)">
                  <Input
                    aria-label="Documents per page"
                    type="number"
                    min={1}
                    max={250}
                    value={String(draft.pageSize)}
                    onChange={(_, data) => setDraft({ ...draft, pageSize: Number(data.value) })}
                  />
                </Field>
                <Field label="Tile size">
                  <Dropdown
                    aria-label="Tile size"
                    value={draft.grid.tileSize}
                    selectedOptions={[draft.grid.tileSize]}
                    onOptionSelect={(_, data) =>
                      patch({ tileSize: data.optionValue as GridConfiguration['tileSize'] })
                    }
                  >
                    {['small', 'medium', 'large'].map((size) => (
                      <Option key={size} value={size}>
                        {size}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
                <Field label="Search by field">
                  <Dropdown
                    aria-label="Search by field"
                    value={
                      attributes.find((a) => a.LogicalName === draft.grid.searchField)?.DisplayName
                        .UserLocalizedLabel.Label || 'Name'
                    }
                    selectedOptions={[draft.grid.searchField]}
                    onOptionSelect={(_, data) => patch({ searchField: data.optionValue })}
                  >
                    {attributes
                      .filter((a) => ['String', 'Memo'].includes(a.AttributeType))
                      .map((a) => (
                        <Option key={a.LogicalName} value={a.LogicalName}>
                          {a.DisplayName.UserLocalizedLabel.Label}
                        </Option>
                      ))}
                  </Dropdown>
                </Field>
              </div>
              <h3>Columns</h3>
              <p>
                Choose columns to display, edit, or filter. Drag grid headers to change their order.
              </p>
              <table style={{ width: '100%' }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: 'left' }}>Column</th>
                    <th>Show</th>
                    <th>Edit</th>
                    <th>Filter</th>
                  </tr>
                </thead>
                <tbody>
                  {attributes.map((a) => (
                    <tr key={a.LogicalName}>
                      <td>{a.DisplayName.UserLocalizedLabel.Label}</td>
                      {(['visibleColumns', 'editableColumns', 'filterColumns'] as const).map(
                        (key) => (
                          <td key={key}>
                            <Checkbox
                              aria-label={`${key === 'visibleColumns' ? 'Show' : key === 'editableColumns' ? 'Edit' : 'Filter'} ${a.DisplayName.UserLocalizedLabel.Label}`}
                              checked={draft.grid[key]?.includes(a.LogicalName) || false}
                              disabled={key === 'editableColumns' && !a.IsValidForUpdate}
                              onChange={(_, data) =>
                                patch({
                                  [key]: data.checked
                                    ? [...(draft.grid[key] || []), a.LogicalName]
                                    : draft.grid[key]?.filter((field) => field !== a.LogicalName),
                                })
                              }
                            />
                          </td>
                        ),
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
              <h3>Create a local document column</h3>
              <p>
                New local columns persist with the mock API. For Dataverse, create and publish the
                column on the Document table, then select its logical name in the PCF configuration.
              </p>
              <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
                <Field label="Column label">
                  <Input
                    aria-label="Column label"
                    value={label}
                    onChange={(_, data) => setLabel(data.value)}
                  />
                </Field>
                <Field label="Logical name">
                  <Input
                    aria-label="Logical name"
                    placeholder="dms_projectcode"
                    value={name}
                    onChange={(_, data) => setName(data.value)}
                  />
                </Field>
                <Field label="Column type">
                  <Dropdown
                    aria-label="Column type"
                    value={type}
                    selectedOptions={[type]}
                    onOptionSelect={(_, data) => setType(data.optionValue || 'String')}
                  >
                    {['String', 'Memo', 'Integer', 'Decimal', 'Boolean', 'DateTime'].map((kind) => (
                      <Option key={kind} value={kind}>
                        {kind}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
              </div>
              <Button
                disabled={busy || !name || !label}
                onClick={() => {
                  setBusy(true);
                  setError('');
                  void fetch('/__mock/columns', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ name, label, type }),
                  })
                    .then(async (response) => {
                      const body = await response.json();
                      if (!response.ok)
                        throw new Error(body.error?.message || 'Could not create column.');
                      setAttributes((previous) => [...previous, body]);
                      setDraft((previous) => ({
                        ...previous,
                        grid: {
                          ...previous.grid,
                          visibleColumns: [...(previous.grid.visibleColumns || []), name],
                          editableColumns: [...(previous.grid.editableColumns || []), name],
                        },
                      }));
                      setName('');
                      setLabel('');
                    })
                    .catch((e) => setError(String(e)))
                    .finally(() => setBusy(false));
                }}
              >
                Create column
              </Button>
            </DialogContent>
            <DialogActions>
              <Button disabled={busy} onClick={() => setOpen(false)}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                disabled={
                  busy ||
                  !Number.isInteger(draft.pageSize) ||
                  draft.pageSize < 1 ||
                  draft.pageSize > 250 ||
                  !draft.grid.visibleColumns?.length
                }
                onClick={() => {
                  onApply(draft);
                  setOpen(false);
                }}
              >
                Apply configuration
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
}
