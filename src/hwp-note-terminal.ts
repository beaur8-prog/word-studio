export type HwpTerminalNote = { mark: string; fingerprint: string };

export type HwpNoteTerminal = {
  width: number;
  fullBodyHeight: number;
  lineTop: number;
  secondBaseline: number;
  pitch: number;
  font: number;
  hanging: number;
  prior: string;
  notes: [HwpTerminalNote, HwpTerminalNote];
};

const fingerprint = (value: unknown): value is string => typeof value === 'string' && /^[0-9a-f]{16}$/.test(value);
const finite = (value: unknown, min: number, max: number): value is number =>
  typeof value === 'number' && Number.isFinite(value) && value >= min && value <= max;

export function noteTerminal(value: unknown): HwpNoteTerminal | undefined {
  if (!value || typeof value !== 'object') return;
  const { width, fullBodyHeight, lineTop, secondBaseline, pitch, font, hanging, prior, notes } = value as HwpNoteTerminal;
  if (!finite(width, 100, 3000) || !finite(fullBodyHeight, 100, 3000) ||
      !finite(lineTop, fullBodyHeight / 2, fullBodyHeight) || !finite(font, 6, 60) ||
      !finite(pitch, font + .5, font * 2) || !finite(hanging, .5, width / 4) ||
      !finite(secondBaseline, lineTop + pitch + font / 2, Math.min(lineTop + pitch + font, fullBodyHeight)) ||
      fullBodyHeight - (lineTop + pitch) > pitch * 4 || !fingerprint(prior) ||
      !Array.isArray(notes) || notes.length !== 2 ||
      !notes.every(note => note && typeof note.mark === 'string' && /^[1-9]\d{0,4}\)$/.test(note.mark) &&
        Number.parseInt(note.mark) <= 65535 && fingerprint(note.fingerprint)) || notes[0].mark === notes[1].mark) return;
  return { width, fullBodyHeight, lineTop, secondBaseline, pitch, font, hanging, prior,
    notes: [{ mark: notes[0].mark, fingerprint: notes[0].fingerprint }, { mark: notes[1].mark, fingerprint: notes[1].fingerprint }] };
}
