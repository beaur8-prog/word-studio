import { expect, test } from '@playwright/test';

for (const scenario of ['sole balanced', 'web preview', 'one column', 'preceding sibling', 'following sibling'] as const) {
  test('measured desktop print column height has the expected scope for ' + scenario, async ({ page }) => {
    await page.goto('/');
    await page.locator('#root .tiptap[role="textbox"]').waitFor();
    await page.evaluate(async scenario => {
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { defaultPage } = await import('/src/types.ts');
      const columns = '<div class="document-columns" data-column-count="' + (scenario === 'one column' ? 1 : 2) +
        '" data-column-space="24" data-column-balance="true"><p style="margin:0">First paragraph</p>' +
        '<p style="margin:0">Second paragraph</p><p style="margin:0">Third paragraph</p><p style="margin:0">Fourth paragraph</p></div>';
      const before = scenario === 'preceding sibling' ? '<p style="margin:0">Preceding body paragraph</p>' : '';
      const after = scenario === 'following sibling' ? '<p style="margin:0">Following body paragraph</p>' : '';
      const layout = await createPrintLayout(before + columns + after, defaultPage, undefined, undefined,
        { nativeDesktop: scenario !== 'web preview' });
      layout.element.classList.replace('measuring', 'preview');
      document.body.classList.add('studio-printing');
    }, scenario);
    const capture = () => page.evaluate(() => {
      const sheet = document.querySelector<HTMLElement>('.print-page')!;
      const origin = sheet.getBoundingClientRect();
      const body = sheet.querySelector<HTMLElement>('.print-content')!;
      const group = body.querySelector<HTMLElement>('[data-column-count]');
      const rect = (element: HTMLElement) => {
        const box = element.getBoundingClientRect();
        return [box.x - origin.x, box.y - origin.y, box.width, box.height];
      };
      return {
        pages: document.querySelectorAll('.print-page').length,
        body: rect(body),
        marked: group?.hasAttribute('data-native-print-column-height') ?? false,
        reserved: group?.style.getPropertyValue('--native-print-column-height') || null,
        cap: group?.style.getPropertyValue('column-height') || null,
        height: group?.getBoundingClientRect().height ?? null,
        paragraphs: [...body.querySelectorAll<HTMLElement>('p')].map(p => ({ text: p.textContent, box: rect(p) })),
      };
    });
    const screen = await capture();
    await page.emulateMedia({ media: 'print' });
    const printed = await capture();
    expect(screen.pages).toBe(1);
    expect(printed.pages).toBe(1);
    expect(printed.paragraphs).toEqual(screen.paragraphs);
    expect(printed.body).toEqual(screen.body);
    if (scenario === 'sole balanced') {
      expect(printed.marked).toBe(true);
      expect(printed.reserved).toBe(screen.cap);
      expect(printed.height!).toBeCloseTo(parseFloat(screen.cap!), 1);
      expect(printed.height!).toBeGreaterThan(screen.height!);
      expect(new Set(printed.paragraphs.map(p => p.box[0])).size).toBe(2);
    } else {
      expect(printed.marked).toBe(false);
      expect(printed.reserved).toBeNull();
      expect(printed.height).toBe(screen.height);
    }
  });
}
