import * as React from 'react';
export interface Gesture {
  shiftKey?: boolean;
  ctrlKey?: boolean;
  metaKey?: boolean;
}
export function chooseSelection(
  current: Set<string>,
  ids: string[],
  anchor: string | undefined,
  id: string,
  event: Gesture,
  toggle = false,
): Set<string> {
  if (event.shiftKey && anchor && ids.includes(anchor) && ids.includes(id)) {
    const a = ids.indexOf(anchor),
      b = ids.indexOf(id),
      range = ids.slice(Math.min(a, b), Math.max(a, b) + 1);
    return new Set(event.ctrlKey || event.metaKey ? [...current, ...range] : range);
  }
  if (toggle || event.ctrlKey || event.metaKey) {
    const next = new Set(current);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    return next;
  }
  return new Set([id]);
}
export function useSelection(ids: string[]) {
  const [selected, setSelected] = React.useState(new Set<string>()),
    anchor = React.useRef<string>();
  const key = ids.join(',');
  React.useEffect(() => {
    setSelected((old) => {
      const next = new Set([...old].filter((id) => ids.includes(id)));
      return next.size === old.size ? old : next;
    });
    if (anchor.current && !ids.includes(anchor.current)) anchor.current = undefined;
  }, [key]);
  const pick = (id: string, event: Gesture = {}, toggle = false) => {
    // React 16 pools events; capture modifiers before the state updater runs.
    const gesture = { shiftKey: event.shiftKey, ctrlKey: event.ctrlKey, metaKey: event.metaKey },
      currentAnchor = anchor.current;
    setSelected((current) => chooseSelection(current, ids, currentAnchor, id, gesture, toggle));
    if (!gesture.shiftKey) anchor.current = id;
  };
  const clear = () => {
    setSelected(new Set());
    anchor.current = undefined;
  };
  const all = () =>
    setSelected((old) => (ids.every((id) => old.has(id)) ? new Set() : new Set(ids)));
  const context = (id: string) => {
    if (!selected.has(id)) {
      setSelected(new Set([id]));
      anchor.current = id;
    }
  };
  return { selected, pick, clear, all, context, setSelected, anchor, ids };
}
export type Selection = ReturnType<typeof useSelection>;
