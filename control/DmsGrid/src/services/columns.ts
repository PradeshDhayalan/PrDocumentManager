import { Attribute } from './documents';
export const builtInColumns = [
  { id: 'icon', label: 'File type', field: '' },
  { id: 'name', label: 'Name', field: 'dms_name' },
  { id: 'modified', label: 'Modified', field: 'modifiedon' },
  { id: 'by', label: 'Modified By', field: '_modifiedby_value' },
  { id: 'type', label: 'Document type', field: 'dms_documenttype' },
  { id: 'status', label: 'Status', field: 'dms_documentstatus' },
  { id: 'expiry', label: 'Expiry date', field: 'dms_expirydate' },
];
export const defaultColumns = builtInColumns.map((column) => column.id);
export function availableColumns(attributes: Attribute[]) {
  return [
    ...builtInColumns,
    ...attributes
      .filter(
        (attribute) => !builtInColumns.some((column) => column.field === attribute.LogicalName),
      )
      .map((attribute) => ({
        id: attribute.LogicalName,
        label: attribute.DisplayName.UserLocalizedLabel.Label,
        field: attribute.LogicalName,
      })),
  ];
}
export function configuredColumns(fields: string[], attributes: Attribute[]) {
  const available = availableColumns(attributes);
  const ids = fields
    .map((field) => available.find((column) => column.field === field || column.id === field)?.id)
    .filter((id): id is string => Boolean(id));
  return ids.length ? [...new Set(['icon', ...ids])] : defaultColumns;
}
export function moveColumn(columns: string[], source: string, target: string) {
  if (source === target || !columns.includes(source) || !columns.includes(target)) return columns;
  const next = columns.filter((id) => id !== source);
  next.splice(columns.indexOf(target), 0, source);
  return next;
}
