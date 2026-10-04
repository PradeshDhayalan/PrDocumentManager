import { ApiError, Entity, Value, formatEntity, formattedSuffix } from './model';
type Predicate = (row: Entity) => boolean;
interface Token {
  kind: 'word' | 'string' | 'number' | 'punct';
  value: string;
}
function tokenize(source: string): Token[] {
  const tokens: Token[] = [];
  let offset = 0;
  while (offset < source.length) {
    const rest = source.slice(offset),
      space = /^\s+/.exec(rest);
    if (space) {
      offset += space[0].length;
      continue;
    }
    const quoted = /^'((?:[^']|'')*)'/.exec(rest);
    if (quoted) {
      tokens.push({ kind: 'string', value: quoted[1].replace(/''/g, "'") });
      offset += quoted[0].length;
      continue;
    }
    const number = /^-?\d+(?:\.\d+)?/.exec(rest);
    if (number) {
      tokens.push({ kind: 'number', value: number[0] });
      offset += number[0].length;
      continue;
    }
    const word = /^[A-Za-z_][A-Za-z0-9_]*/.exec(rest);
    if (word) {
      tokens.push({ kind: 'word', value: word[0] });
      offset += word[0].length;
      continue;
    }
    if ('(),'.includes(rest[0])) {
      tokens.push({ kind: 'punct', value: rest[0] });
      offset++;
      continue;
    }
    throw new ApiError(400, 'InvalidFilter', 'Unsupported filter syntax.');
  }
  return tokens;
}
export function parseFilter(source: string): Predicate {
  const tokens = tokenize(source);
  let cursor = 0;
  const peek = (word: string) => tokens[cursor]?.value === word;
  function take(expected?: string): Token {
    const token = tokens[cursor++];
    if (!token || (expected && token.value !== expected))
      throw new ApiError(400, 'InvalidFilter', 'Malformed filter.');
    return token;
  }
  function literal(): Value {
    const token = take();
    if (token.kind === 'string') return token.value;
    if (token.kind === 'number') return Number(token.value);
    if (token.value === 'null') return null;
    if (token.value === 'true') return true;
    if (token.value === 'false') return false;
    throw new ApiError(400, 'InvalidFilter', 'Expected literal.');
  }
  function atom(): Predicate {
    if (peek('(')) {
      take('(');
      const predicate = or();
      take(')');
      return predicate;
    }
    if (peek('contains')) {
      take();
      take('(');
      const key = take().value;
      take(',');
      const value = literal();
      take(')');
      return (row) =>
        String(row[key] ?? '')
          .toLowerCase()
          .includes(String(value).toLowerCase());
    }
    const key = take().value,
      operator = take().value,
      value = literal();
    if (!['eq', 'ne', 'gt', 'ge', 'lt', 'le'].includes(operator))
      throw new ApiError(400, 'InvalidFilter', 'Unsupported comparison.');
    return (row) => {
      const actual = row[key] ?? null;
      if (operator === 'eq') return actual === value;
      if (operator === 'ne') return actual !== value;
      if (actual === null || value === null || typeof actual !== typeof value) return false;
      const comparison = actual === value ? 0 : actual < value ? -1 : 1;
      return operator === 'gt'
        ? comparison > 0
        : operator === 'ge'
          ? comparison >= 0
          : operator === 'lt'
            ? comparison < 0
            : comparison <= 0;
    };
  }
  function and(): Predicate {
    let left = atom();
    while (peek('and')) {
      take();
      const previous = left,
        right = atom();
      left = (row) => previous(row) && right(row);
    }
    return left;
  }
  function or(): Predicate {
    let left = and();
    while (peek('or')) {
      take();
      const previous = left,
        right = and();
      left = (row) => previous(row) || right(row);
    }
    return left;
  }
  const predicate = or();
  if (cursor !== tokens.length)
    throw new ApiError(400, 'InvalidFilter', 'Unexpected filter token.');
  return predicate;
}
export function selectEntity(row: Entity, select: string | null): Entity {
  const formatted = formatEntity(row);
  if (!select) return formatted;
  const output: Entity = { '@odata.etag': row['@odata.etag'] };
  for (const key of select.split(',').map((k) => k.trim())) {
    if (key in formatted) output[key] = formatted[key];
    if (key + formattedSuffix in formatted)
      output[key + formattedSuffix] = formatted[key + formattedSuffix];
  }
  return output;
}
export function queryEntities(
  rows: Entity[],
  url: URL,
  prefer = '',
): { value: Entity[]; '@odata.count'?: number; '@odata.nextLink'?: string } {
  const params = url.searchParams;
  let filtered = rows.filter(
    params.has('$filter') ? parseFilter(params.get('$filter')!) : () => true,
  );
  const count = filtered.length,
    sorts = (params.get('$orderby') || 'modifiedon desc')
      .split(',')
      .map((s) => s.trim().split(/\s+/));
  filtered = [...filtered].sort((a, b) => {
    for (const [key, direction] of sorts) {
      const av = a[key],
        bv = b[key];
      if (av === bv) continue;
      const comparison =
        av == null
          ? -1
          : bv == null
            ? 1
            : typeof av === 'number' && typeof bv === 'number'
              ? av - bv
              : String(av).localeCompare(String(bv));
      return comparison * (direction === 'desc' ? -1 : 1);
    }
    return String(a.dms_documentid || a.annotationid).localeCompare(
      String(b.dms_documentid || b.annotationid),
    );
  });
  const integer = (value: string | null, fallback: number) => {
    if (value === null) return fallback;
    if (!/^\d+$/.test(value))
      throw new ApiError(400, 'InvalidQuery', 'Paging requires non-negative integers.');
    return Number(value);
  };
  const skip = integer(params.get('$skip'), 0),
    top = integer(params.get('$top'), count),
    pageSize = Math.max(
      1,
      Math.min(250, integer(/odata.maxpagesize=(\d+)/i.exec(prefer)?.[1] ?? null, 50)),
    );
  const end = Math.min(skip + pageSize, top),
    page = filtered.slice(skip, end),
    output: ReturnType<typeof queryEntities> = {
      value: page.map((row) => selectEntity(row, params.get('$select'))),
    };
  if (params.get('$count') === 'true') output['@odata.count'] = count;
  if (end < Math.min(count, top)) {
    const next = new URL(url);
    next.searchParams.set('$skip', String(end));
    output['@odata.nextLink'] = next.toString();
  }
  return output;
}
