export interface GridConfiguration {
  visibleColumns?: string[];
  editableColumns?: string[];
  filterColumns: string[];
  searchField: string;
  tileSize: 'small' | 'medium' | 'large';
}
export function parseGridConfiguration(json?: string): GridConfiguration {
  const input = json ? JSON.parse(json) : {};
  if (!input || typeof input !== 'object' || Array.isArray(input))
    throw new Error('Grid configuration must be a JSON object.');
  for (const key of ['visibleColumns', 'editableColumns', 'filterColumns']) {
    if (
      input[key] !== undefined &&
      (!Array.isArray(input[key]) ||
        input[key].some((v: unknown) => typeof v !== 'string' || !/^[a-z_][a-z0-9_]*$/.test(v)))
    )
      throw new Error(`${key} must contain logical column names.`);
  }
  if (input.searchField !== undefined && !/^[a-z_][a-z0-9_]*$/.test(input.searchField))
    throw new Error('Search field must be a logical column name.');
  if (input.tileSize !== undefined && !['small', 'medium', 'large'].includes(input.tileSize))
    throw new Error('Tile size must be small, medium or large.');
  return {
    ...input,
    filterColumns: input.filterColumns ?? ['dms_documenttype', 'dms_documentstatus'],
    searchField: input.searchField ?? 'dms_name',
    tileSize: input.tileSize ?? 'medium',
  };
}
