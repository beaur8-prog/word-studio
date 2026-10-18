/** Load local outline faces before the first editor measurement. */
export async function initializeOutlineFonts() {
  if (!window.wordDesktop?.outlineFonts) return;
  try {
    const faces = await window.wordDesktop.outlineFonts();
    await Promise.all(faces.flatMap(face => face.families.map(async family => {
      const font = new FontFace(family, new Uint8Array(face.data).buffer);
      await font.load(); document.fonts.add(font);
    })));
  } catch { /* An unavailable optional Windows font must not prevent opening documents. */ }
}
