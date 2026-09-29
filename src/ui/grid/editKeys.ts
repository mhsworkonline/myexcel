// Edit-mode keyboard handling shared by the in-cell editor and the formula bar.
import { toggleRefAt } from '../../model/formula';
import { cancelEdit, commitEdit } from '../../state/actions/edit';
import { canPoint, pointMove } from '../../state/actions/point';
import { S, setState } from '../../state/store';

export interface AutoState {
  items: string[];
  index: number;
  start: number;
}

/** Returns true if handled (caller should preventDefault). */
export function handleEditKey(
  e: KeyboardEvent,
  el: HTMLTextAreaElement,
  auto: AutoState | null,
  acceptAuto: (name: string) => void,
  closeAuto: () => void,
  moveAuto: (d: number) => void,
): boolean {
  const st = S();
  const ed = st.edit;
  if (!ed) return false;
  const key = e.key;
  if (auto && auto.items.length) {
    if (key === 'ArrowDown' || key === 'ArrowUp') {
      moveAuto(key === 'ArrowDown' ? 1 : -1);
      return true;
    }
    if (key === 'Tab') {
      acceptAuto(auto.items[auto.index]);
      return true;
    }
    if (key === 'Escape') {
      closeAuto();
      return true;
    }
  }
  switch (key) {
    case 'Enter': {
      if (e.altKey) {
        const s = el.selectionStart;
        const t = ed.text.slice(0, s) + '\n' + ed.text.slice(el.selectionEnd);
        setState({ edit: { ...ed, text: t, caret: s + 1, mode: 'edit', point: undefined } });
        return true;
      }
      if (e.ctrlKey && e.shiftKey) {
        commitEdit('none');
        return true;
      }
      if (e.ctrlKey) {
        commitEdit('none', { fillSelection: true });
        return true;
      }
      commitEdit(e.shiftKey ? 'up' : 'down');
      return true;
    }
    case 'Tab':
      commitEdit(e.shiftKey ? 'left' : 'right');
      return true;
    case 'Escape':
      cancelEdit();
      return true;
    case 'F2':
      setState({ edit: { ...ed, mode: ed.mode === 'enter' ? 'edit' : 'enter', point: undefined } });
      return true;
    case 'F4': {
      const res = toggleRefAt(ed.text, el.selectionStart);
      setState({ edit: { ...ed, text: res.text, caret: res.caret, point: undefined } });
      return true;
    }
    case 'ArrowUp':
    case 'ArrowDown':
    case 'ArrowLeft':
    case 'ArrowRight': {
      if (ed.mode === 'edit' || ed.fromBar) return false;
      const dr = key === 'ArrowUp' ? -1 : key === 'ArrowDown' ? 1 : 0;
      const dc = key === 'ArrowLeft' ? -1 : key === 'ArrowRight' ? 1 : 0;
      if (canPoint()) {
        pointMove(dr, dc, e.shiftKey);
        return true;
      }
      commitEdit(dr > 0 ? 'down' : dr < 0 ? 'up' : dc > 0 ? 'right' : 'left');
      return true;
    }
  }
  if ((e.ctrlKey || e.metaKey) && key.toLowerCase() === 'a' && ed.text.startsWith('=')) {
    // Ctrl+A after a function name opens the arguments dialog
    import('../../state/store').then((m) => m.openDialog('insertFunction'));
    return true;
  }
  return false;
}
