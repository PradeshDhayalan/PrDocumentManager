import * as React from 'react';
interface Point {
  x: number;
  y: number;
}
interface Gesture {
  start: Point;
  current: Point;
  initial: string[];
  additive: boolean;
  pointerId: number;
  active: boolean;
}
export function useDragSelection(selected: string[], setSelected: (ids: string[]) => void) {
  const container = React.useRef<HTMLDivElement>(null);
  const gesture = React.useRef<Gesture>();
  const suppressClick = React.useRef(false);
  const [rectangle, setRectangle] = React.useState<{
    left: number;
    top: number;
    width: number;
    height: number;
  }>();
  const finish = (cancel = false) => {
    const current = gesture.current;
    gesture.current = undefined;
    if (current?.active) {
      if (cancel) setSelected(current.initial);
      suppressClick.current = true;
    }
    setRectangle(undefined);
    if (current && container.current?.hasPointerCapture(current.pointerId))
      container.current.releasePointerCapture(current.pointerId);
  };
  React.useEffect(
    () => () => {
      gesture.current = undefined;
    },
    [],
  );
  const handlers = {
    onPointerDown: (event: React.PointerEvent<HTMLDivElement>) => {
      suppressClick.current = false;
      if (
        event.button !== 0 ||
        event.pointerType !== 'mouse' ||
        (event.target as HTMLElement).closest(
          'button,input,a,[role="checkbox"],[role="combobox"],.select-cell',
        )
      )
        return;
      gesture.current = {
        start: { x: event.clientX, y: event.clientY },
        current: { x: event.clientX, y: event.clientY },
        initial: [...selected],
        additive: event.ctrlKey || event.metaKey || event.shiftKey,
        pointerId: event.pointerId,
        active: false,
      };
    },
    onPointerMove: (event: React.PointerEvent<HTMLDivElement>) => {
      const current = gesture.current;
      if (!current || current.pointerId !== event.pointerId || !container.current) return;
      if (
        !current.active &&
        Math.hypot(event.clientX - current.start.x, event.clientY - current.start.y) < 5
      )
        return;
      if (!current.active) {
        current.active = true;
        container.current.setPointerCapture(event.pointerId);
        window.getSelection()?.removeAllRanges();
      }
      event.preventDefault();
      const bounds = container.current.getBoundingClientRect();
      current.current = {
        x: Math.max(bounds.left, Math.min(bounds.right, event.clientX)),
        y: Math.max(bounds.top, Math.min(bounds.bottom, event.clientY)),
      };
      const left = Math.min(current.start.x, current.current.x),
        top = Math.min(current.start.y, current.current.y);
      const right = Math.max(current.start.x, current.current.x),
        bottom = Math.max(current.start.y, current.current.y);
      const hits = Array.from(container.current.querySelectorAll<HTMLElement>('[data-document-id]'))
        .filter((element) => {
          const item = element.getBoundingClientRect();
          return item.left < right && item.right > left && item.top < bottom && item.bottom > top;
        })
        .map((element) => element.dataset.documentId!);
      setSelected([...new Set([...(current.additive ? current.initial : []), ...hits])]);
      setRectangle({ left, top, width: right - left, height: bottom - top });
    },
    onPointerUp: () => finish(),
    onPointerCancel: () => finish(true),
    onLostPointerCapture: () => {
      if (gesture.current) finish(true);
    },
    onClickCapture: (event: React.MouseEvent<HTMLDivElement>) => {
      if (suppressClick.current) {
        event.preventDefault();
        event.stopPropagation();
        suppressClick.current = false;
      }
    },
    onDragStart: (event: React.DragEvent<HTMLDivElement>) => {
      if (gesture.current) event.preventDefault();
    },
    onKeyDownCapture: (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (event.key === 'Escape' && gesture.current?.active) {
        event.preventDefault();
        event.stopPropagation();
        finish(true);
      }
    },
  };
  return { container, handlers, rectangle };
}
