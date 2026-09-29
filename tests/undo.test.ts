import { describe, expect, it } from 'vitest';
import { Engine } from '@/engine/Engine';
import { FnCommand, History, StructCommand, Tx } from '@/model/commands';
import { Sheet } from '@/model/sheet';
import { Workbook } from '@/model/workbook';

function setup() {
  const wb = Workbook.createDefault();
  const engine = new Engine(wb);
  const history = new History();
  const s = wb.sheets[0];
  const run = (label: string, fn: (tx: Tx) => void) => {
    const tx = new Tx(wb, label);
    fn(tx);
    const cmd = tx.finish();
    if (cmd) history.push({ cmd });
  };
  return { wb, engine, history, s, run };
}

describe('command layer undo/redo', () => {
  it('undoes and redoes cell edits, keeping the engine in sync', () => {
    const { wb, engine, history, s, run } = setup();
    run('Typing', (tx) => {
      tx.setCell(s, 0, 0, { v: 2 });
      tx.setCell(s, 1, 0, { v: 3 });
    });
    run('Typing', (tx) => tx.setCell(s, 2, 0, { f: '=A1*A2' }));
    expect(engine.getValue(s, 2, 0)).toBe(6);
    run('Typing', (tx) => tx.setCell(s, 0, 0, { v: 10 }));
    expect(engine.getValue(s, 2, 0)).toBe(30);
    history.undo(wb);
    expect(s.getCell(0, 0)?.v).toBe(2);
    expect(engine.getValue(s, 2, 0)).toBe(6);
    history.undo(wb);
    expect(s.getCell(2, 0)).toBeUndefined();
    history.redo(wb);
    history.redo(wb);
    expect(engine.getValue(s, 2, 0)).toBe(30);
    expect(history.canRedo).toBe(false);
  });

  it('undoes formatting and sheet metadata changes', () => {
    const { wb, history, s, run } = setup();
    const bold = wb.styles.intern({ bold: true });
    run('Bold', (tx) => tx.patchCell(s, 0, 0, { s: bold }));
    run('Merge', (tx) => tx.setMeta(s, 'merges', [{ r1: 0, c1: 0, r2: 1, c2: 1 }]));
    run('Width', (tx) => tx.setMeta(s, 'colWidths', new Map([[0, 100]])));
    history.undo(wb);
    expect(s.colWidths.size).toBe(0);
    history.undo(wb);
    expect(s.merges).toEqual([]);
    history.undo(wb);
    expect(s.getCell(0, 0)).toBeUndefined();
    history.redo(wb);
    expect(wb.styles.get(s.getCell(0, 0)?.s).bold).toBe(true);
  });

  it('undoes row deletion including #REF! formulas and restores dependents', () => {
    const { wb, engine, history, s, run } = setup();
    run('Data', (tx) => {
      tx.setCell(s, 0, 0, { v: 1 });
      tx.setCell(s, 1, 0, { v: 2 });
      tx.setCell(s, 2, 0, { v: 3 });
      tx.setCell(s, 3, 0, { f: '=SUM(A1:A3)' });
      tx.setCell(s, 4, 0, { f: '=A2*10' });
    });
    expect(engine.getValue(s, 3, 0)).toBe(6);
    run('Delete Rows', (tx) => tx.run(new StructCommand('Delete Rows', s.id, { axis: 'row', at: 1, delta: -1 })));
    expect(s.getCell(2, 0)?.f).toBe('=SUM(A1:A2)');
    expect(s.getCell(3, 0)?.f).toBe('=#REF!*10');
    expect(engine.getValue(s, 2, 0)).toBe(4);
    history.undo(wb);
    expect(s.getCell(1, 0)?.v).toBe(2);
    expect(s.getCell(3, 0)?.f).toBe('=SUM(A1:A3)');
    expect(s.getCell(4, 0)?.f).toBe('=A2*10');
    expect(engine.getValue(s, 3, 0)).toBe(6);
    expect(engine.getValue(s, 4, 0)).toBe(20);
    history.redo(wb);
    expect(s.getCell(2, 0)?.f).toBe('=SUM(A1:A2)');
  });

  it('undoes column insertion and sheet add/rename', () => {
    const { wb, engine, history, s, run } = setup();
    run('Data', (tx) => {
      tx.setCell(s, 0, 0, { v: 5 });
      tx.setCell(s, 0, 1, { f: '=A1+1' });
    });
    run('Insert Col', (tx) => tx.run(new StructCommand('Insert Columns', s.id, { axis: 'col', at: 0, delta: 1 })));
    expect(s.getCell(0, 2)?.f).toBe('=B1+1');
    expect(engine.getValue(s, 0, 2)).toBe(6);
    history.undo(wb);
    expect(s.getCell(0, 1)?.f).toBe('=A1+1');
    const s2 = new Sheet('Data');
    run('Insert Sheet', (tx) => tx.run(new FnCommand('Insert Sheet', (w) => w.insertSheet(s2, 1), (w) => w.removeSheet(s2))));
    run('Ref', (tx) => tx.setCell(s, 5, 5, { f: '=Data!A1+1' }));
    run('Rename', (tx) => tx.run(new FnCommand('Rename', (w) => w.renameSheet(s2, 'Numbers'), (w) => w.renameSheet(s2, 'Data'))));
    expect(s.getCell(5, 5)?.f).toBe('=Numbers!A1+1');
    history.undo(wb);
    expect(s.getCell(5, 5)?.f).toBe('=Data!A1+1');
    history.undo(wb);
    history.undo(wb);
    expect(wb.sheets.length).toBe(1);
  });
});
