import { chooseSelection } from '../src/hooks/useSelection';
const ids = ['a', 'b', 'c', 'd', 'e'];
test('selection replaces, toggles, selects ranges and adds ranges without mutating input', () => {
  const current = new Set(['e']);
  expect([...chooseSelection(current, ids, 'a', 'c', {})]).toEqual(['c']);
  expect([...chooseSelection(current, ids, 'a', 'c', { ctrlKey: true })]).toEqual(['e', 'c']);
  expect([...chooseSelection(current, ids, 'b', 'd', { shiftKey: true })]).toEqual(['b', 'c', 'd']);
  expect([...chooseSelection(current, ids, 'b', 'd', { shiftKey: true, metaKey: true })]).toEqual([
    'e',
    'b',
    'c',
    'd',
  ]);
  expect([...chooseSelection(current, ids, 'a', 'e', {}, true)]).toEqual([]);
  expect([...current]).toEqual(['e']);
});
