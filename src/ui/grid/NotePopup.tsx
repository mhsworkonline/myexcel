'use client';
import { S } from '../../state/store';

export function NotePopup({ r, c, x, y }: { r: number; c: number; x: number; y: number }) {
  const note = S().wb.activeSheet.getCell(r, c)?.note;
  if (!note) return null;
  return (
    <div className="absolute z-30 w-[200px] min-h-[70px] px-1.5 py-1 text-[11px] shadow-md pointer-events-none" style={{ left: x, top: y, background: '#FFFFE1', border: '1px solid #767676', color: '#000' }}>
      {note.threaded ? (
        <div>
          <div className="font-semibold">{note.author ?? 'User'}</div>
          <div className="whitespace-pre-wrap">{note.text}</div>
          {note.replies?.map((rp, i) => (
            <div key={i} className="mt-1 border-t pt-1" style={{ borderColor: '#ddd' }}>
              <div className="font-semibold">{rp.author}</div>
              <div className="whitespace-pre-wrap">{rp.text}</div>
            </div>
          ))}
        </div>
      ) : (
        <div className="whitespace-pre-wrap">
          {note.author && <b>{note.author}:</b>}
          {note.author && <br />}
          {note.text}
        </div>
      )}
    </div>
  );
}
