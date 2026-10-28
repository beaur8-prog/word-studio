import { test, expect, type Page } from "@playwright/test";

const pageSettings = {
  size: "A4",
  orientation: "portrait",
  margin: 25.4,
  header: "",
  footer: "",
  pageNumbers: false,
};
const paragraph = (text: string) => ({
  type: "paragraph",
  attrs: { lineSpacing: "20px", spaceAfter: "0px" },
  content: [{ type: "text", text }],
});

for (const behindDocument of [true, false]) test(`non-wrapping ${behindDocument ? 'background' : 'foreground'} art preserves mixed-font text and native line wrapping`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async behindDocument => {
    const { applyFloatingLayout } = await import('/src/floating-layout.ts');
    const root = document.createElement('div'); root.className = 'tiptap';
    root.style.cssText = 'position:fixed;left:-10000px;top:0;width:420px;padding:0;min-height:0';
    const image = document.createElement('img'); image.width = 5; image.height = 800;
    image.style.cssText = 'position:absolute';
    image.dataset.imageFloating = JSON.stringify({ horizontalPosition: { relative: 'column', offset: 200 * 9525 }, verticalPosition: { relative: 'paragraph', offset: 0 }, wrap: { type: 0 }, behindDocument });
    const anchor = document.createElement('p'); anchor.style.cssText = 'line-height:1px;font-size:1px;margin:0'; anchor.append(image); root.append(anchor);
    for (const align of ['left', 'center', 'right', 'justify']) {
      const p = document.createElement('p'); p.style.cssText = `font-family:Times New Roman;font-size:16px;line-height:23px;text-align:${align};margin:0`;
      p.innerHTML = ('Mixed <span style="font-size:20px">LARGER</span> text, punctuation and words. ').repeat(8); root.append(p);
    }
    document.body.append(root); await document.fonts.ready;
    const measure = () => [...root.querySelectorAll('p')].slice(1).map(p => {
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), rects = [];
      while (walker.nextNode()) {
        const range = document.createRange(); range.selectNodeContents(walker.currentNode);
        rects.push(...[...range.getClientRects()].map(r => [r.x, r.y, r.width, r.height]));
      }
      return { height: p.getBoundingClientRect().height, rects };
    });
    try { const before = measure(); applyFloatingLayout(root); return { before, after: measure(), positioned: image.style.position }; }
    finally { root.remove(); }
  }, behindDocument);
  expect(result.positioned).toBe('absolute');
  expect(result.after).toEqual(result.before);
});
async function open(page: Page, content: unknown[]) {
  await page.goto("/");
  const doc = {
    format: "word-studio",
    version: 1,
    title: "Layout flow",
    updatedAt: new Date().toISOString(),
    page: pageSettings,
    comments: [],
    content: { type: "doc", content },
  };
  await page
    .locator("input[type=file]")
    .first()
    .setInputFiles({
      name: "flow.wordstudio",
      mimeType: "application/json",
      buffer: Buffer.from(JSON.stringify(doc)),
    });
  await expect(page.locator(".document-stack .tiptap")).toContainText("FLOW");
}
async function preview(page: Page) {
  await page.getByRole("button", { name: "보기", exact: true }).click();
  await page.getByRole("button", { name: "인쇄 미리보기", exact: true }).click();
  await expect(page.locator(".print-page").first()).toBeVisible();
}
for (const balance of [false, true])
  test(`long ${balance ? "balanced" : "unbalanced"} columns flow down paper pages, retain a partial first page, and remain editable`, async ({
    page,
  }) => {
    const text = "FLOW 한글과 English column text 0123456789. ".repeat(240);
    await open(page, [
      paragraph("FLOW preceding full-width paragraph."),
      {
        type: "columnGroup",
        attrs: { count: 2, space: 48, balance },
        content: [paragraph(text)],
      },
      paragraph("FLOW following paragraph."),
    ]);
    await expect
      .poll(async () => Number(await page.locator(".document-stack").getAttribute("data-pages")))
      .toBeGreaterThan(1);
    await expect
      .poll(async () =>
        page.locator(".document-columns").evaluate((g) => g.scrollWidth <= g.clientWidth + 2),
      )
      .toBe(true);
    const metrics = await page.evaluate(() => {
      const sheet = document.querySelector<HTMLElement>(".document-stack")!,
        group = sheet.querySelector<HTMLElement>(".document-columns")!;
      const rect = sheet.getBoundingClientRect(),
        scale = rect.width / sheet.offsetWidth,
        h = Number(sheet.dataset.pageHeight),
        top = Number(sheet.dataset.margin),
        bottom = Number(sheet.dataset.bottomMargin);
      const walker = document.createTreeWalker(group, NodeFilter.SHOW_TEXT),
        points: { x: number; y: number; bottom: number }[] = [];
      while (walker.nextNode()) {
        const t = walker.currentNode;
        for (let i = 0; i < (t.textContent?.length || 0); i += 37) {
          const r = document.createRange();
          r.setStart(t, i);
          r.setEnd(t, i + 1);
          const b = r.getBoundingClientRect();
          if (b.width)
            points.push({
              x: (b.left - rect.left) / scale,
              y: (b.top - rect.top) / scale,
              bottom: (b.bottom - rect.top) / scale,
            });
        }
      }
      const invalid = points.filter((p) => {
        const n = Math.floor((p.y - top + 0.5) / (h + 24));
        return p.y < n * (h + 24) + top - 1 || p.bottom > n * (h + 24) + h - bottom + 1;
      });
      const before = sheet.querySelector(".tiptap > p")!.getBoundingClientRect();
      const range = document.createRange();
      range.selectNodeContents(group.querySelector("p")!);
      const first = range.getClientRects()[0];
      return {
        invalid: invalid.slice(0, 5),
        points,
        pages: Number(sheet.dataset.pages),
        firstY: (first.top - rect.top) / scale,
        beforeBottom: (before.bottom - rect.top) / scale,
      };
    });
    expect(metrics.invalid).toEqual([]);
    expect(metrics.firstY).toBeGreaterThanOrEqual(metrics.beforeBottom - 1);
    expect(metrics.firstY).toBeLessThan(200);
    // Place the caret on a later paper's column, then exercise editing and undo.
    const point = metrics.points.find((p) => p.y > 1200)!;
    await page.locator(".document-stack").evaluate((sheet, y) => {
      let parent = sheet.parentElement;
      while (parent && getComputedStyle(parent).overflowY !== "auto") parent = parent.parentElement;
      if (parent) parent.scrollTop = y - 250;
    }, point.y);
    const paper = await page.locator(".document-stack").boundingBox();
    await page.mouse.click(paper!.x + point.x + 1, paper!.y + point.y + 5);
    await page.keyboard.type("EDIT");
    await expect(page.locator(".document-columns")).toContainText("EDIT");
    await page.keyboard.press("ControlOrMeta+z");
    await expect(page.locator(".document-columns")).not.toContainText("EDIT");
    await preview(page);
    expect(await page.locator(".print-page").count()).toBe(metrics.pages);
    expect((await page.locator(".print-content").allTextContents()).join("")).toBe(
      "FLOW preceding full-width paragraph." + text + "FLOW following paragraph.",
    );
  });

const pixel =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aN1sAAAAASUVORK5CYII=";
const geometry = (polygon?: number[][]) =>
  JSON.stringify({
    horizontalPosition: { relative: "margin", align: "center" },
    verticalPosition: { relative: "paragraph", offset: 40 * 9525 },
    wrap: { type: polygon ? 2 : 1, side: "bothSides" },
    ...(polygon ? { polygon } : {}),
  });
const image = (polygon?: number[][]) => ({
  type: "inlineImage",
  attrs: { src: pixel, width: 200, height: 160, floating: geometry(polygon) },
});
async function flowMetrics(page: Page, selector: string) {
  return page.locator(selector).evaluate((root) => {
    const img = root.querySelector("img[data-image-floating]")!,
      ir = img.getBoundingClientRect(),
      rects: { left: number; right: number; top: number; bottom: number }[] = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode;
      if (text.parentElement?.closest("[data-flow-widget]")) continue;
      for (let i = 0; i < (text.textContent?.length || 0); i++) {
        if (!text.textContent![i].trim()) continue;
        const range = document.createRange();
        range.setStart(text, i);
        range.setEnd(text, i + 1);
        const r = range.getBoundingClientRect();
        if (r.width)
          rects.push({
            left: r.left,
            right: r.right,
            top: r.top,
            bottom: r.bottom,
          });
      }
    }
    const around = rects.filter((r) => r.top < ir.bottom - 0.5 && r.bottom > ir.top + 0.5),
      lines = new Map<number, { left: boolean; right: boolean }>();
    for (const r of around) {
      const key = Math.round(r.top),
        line = lines.get(key) || { left: false, right: false };
      if (r.right <= ir.left + 0.5) line.left = true;
      if (r.left >= ir.right - 0.5) line.right = true;
      lines.set(key, line);
    }
    return {
      image: ir.toJSON(),
      rects,
      around,
      both: [...lines.values()].filter((l) => l.left && l.right).length,
      overlap: around.filter((r) => r.left < ir.right - 0.5 && r.right > ir.left + 0.5).length,
      spacers: root.querySelectorAll("[data-flow-widget]").length,
    };
  });
}
test("a centered picture has text on both sides of the same lines in the editor and PDF tree without changing its paragraph", async ({
  page,
}) => {
  const text = "FLOW text around the centered picture with bold and normal words. ".repeat(45);
  await open(page, [
    {
      ...paragraph(text),
      content: [image(), { type: "text", text, marks: [{ type: "bold" }] }],
    },
  ]);
  await expect
    .poll(async () => (await flowMetrics(page, ".document-stack .tiptap")).both)
    .toBeGreaterThan(2);
  const editor = await flowMetrics(page, ".document-stack .tiptap");
  expect(editor.overlap).toBe(0);
  expect(editor.spacers).toBeGreaterThan(0);
  expect(await page.locator(".document-stack .tiptap > p").count()).toBe(1);
  await page.locator(".document-stack .tiptap").click();
  await page.keyboard.press("ControlOrMeta+End");
  await page.keyboard.type("EDIT");
  await expect(page.locator(".document-stack .tiptap")).toContainText("EDIT");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(page.locator(".document-stack .tiptap")).not.toContainText("EDIT");
  await preview(page);
  const printed = await flowMetrics(page, ".print-page:first-of-type .print-content");
  expect(printed.both).toBeGreaterThan(2);
  expect(printed.overlap).toBe(0);
  expect((await page.locator(".print-content").allTextContents()).join("")).toBe(text);
});

test("tight polygon wrapping uses the free region beside a sloping contour", async ({ page }) => {
  const polygon = [
      [0, 0],
      [1, 1],
      [0, 1],
    ],
    text = "FLOW a triangular picture allows more text near its narrow upper edge. ".repeat(35);
  await open(page, [{ ...paragraph(text), content: [image(polygon), { type: "text", text }] }]);
  await expect
    .poll(async () => (await flowMetrics(page, ".document-stack .tiptap")).spacers)
    .toBeGreaterThan(0);
  const result = await flowMetrics(page, ".document-stack .tiptap"),
    r = result.image;
  const reclaimed = result.around.filter(
    (c) =>
      c.left >= r.left &&
      c.right < r.right &&
      c.left > r.left + r.width * Math.min(1, (c.bottom - r.top) / r.height) + 0.5,
  );
  expect(reclaimed.length).toBeGreaterThan(5);
  expect(
    result.around.filter(
      (c) =>
        c.right > r.left + 0.5 &&
        c.left < r.left + r.width * Math.max(0, (c.top - r.top) / r.height) - 0.5,
    ),
  ).toEqual([]);
  await preview(page);
  expect(
    (await flowMetrics(page, ".print-page:first-of-type .print-content")).spacers,
  ).toBeGreaterThan(0);
});

test("centered wrapping spans later paragraphs, survives 200% zoom and saves without layout widgets", async ({
  page,
}) => {
  const text =
    "FLOW following paragraph wraps on both sides of a picture anchored in the previous paragraph. ".repeat(
      18,
    );
  await open(page, [
    {
      ...paragraph("FLOW anchor"),
      content: [image(), { type: "text", text: "FLOW anchor" }],
    },
    paragraph(text),
  ]);
  await expect
    .poll(async () => (await flowMetrics(page, ".document-stack .tiptap")).both)
    .toBeGreaterThan(2);
  await page.getByRole("slider", { name: "확대 축소" }).fill("200");
  await expect(page.getByRole("button", { name: "200%", exact: true })).toBeVisible();
  await expect
    .poll(async () => (await flowMetrics(page, ".document-stack .tiptap")).overlap)
    .toBe(0);
  expect((await flowMetrics(page, ".document-stack .tiptap")).both).toBeGreaterThan(2);
  const download = page.waitForEvent("download");
  await page.keyboard.press("ControlOrMeta+s");
  const file = await download;
  const { readFile } = await import("node:fs/promises");
  const saved = JSON.parse(await readFile((await file.path())!, "utf8"));
  expect(JSON.stringify(saved)).not.toContain("flow-widget");
  expect(saved.content.content).toHaveLength(2);
  await page.getByRole("button", { name: "200%", exact: true }).click();
  await page
    .locator("input[type=file]")
    .first()
    .setInputFiles((await file.path())!);
  await expect
    .poll(async () => (await flowMetrics(page, ".document-stack .tiptap")).both)
    .toBeGreaterThan(2);
  await preview(page);
  const result = await flowMetrics(page, ".print-page:first-of-type .print-content");
  expect(result.overlap).toBe(0);
  expect(result.both).toBeGreaterThan(2);
});

test("tight DrawingML polygon and wrap side survive repeated DOCX saves", async ({ page }) => {
  await page.goto("/");
  const polygon = [
    [0, 0],
    [0.75, 0],
    [1, 1],
    [0, 1],
  ];
  const result = await page.evaluate(
    async ({ pixel, geometry, pageSettings }) => {
      // @ts-expect-error Vite module.
      const io = await import("/src/document-io.ts");
      // @ts-expect-error Vite module.
      const { parseDocumentHtml } = await import("/src/extensions.ts");
      let doc = {
        format: "word-studio",
        version: 1,
        title: "Polygon",
        updatedAt: new Date().toISOString(),
        page: pageSettings,
        comments: [],
        content: {
          type: "doc",
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "inlineImage",
                  attrs: {
                    src: pixel,
                    width: 200,
                    height: 160,
                    alt: "Contour original",
                    floating: geometry,
                  },
                },
                { type: "text", text: "FLOW contour" },
              ],
            },
          ],
        },
      };
      const results = [];
      for (let i = 0; i < 2; i++) {
        const file = await io.exportDocx(doc),
          opened = await io.importFile(new File([file], "polygon.docx"));
        doc = { ...opened.document, content: parseDocumentHtml(opened.html) };
        io.validateDocument(doc);
        results.push(doc);
      }
      return results;
    },
    { pixel, geometry: geometry(polygon), pageSettings },
  );
  for (const doc of result) {
    const img = doc.content.content[0].content.find((n: any) => n.type === "inlineImage");
    const g = JSON.parse(img.attrs.floating);
    expect(g.polygon).toEqual(polygon);
    expect(g.wrap).toEqual({ type: 2, side: "bothSides" });
    expect(img.attrs.alt).toBe("Contour original");
  }
});

test("a picture below a short final paragraph remains within the editor paper and preview", async ({
  page,
}) => {
  await open(page, [
    {
      ...paragraph("FLOW short"),
      content: [image(), { type: "text", text: "FLOW short" }],
    },
  ]);
  await expect
    .poll(async () =>
      page
        .locator(".document-stack .tiptap")
        .evaluate(
          (root) =>
            root.getBoundingClientRect().bottom >=
            root.querySelector("img")!.getBoundingClientRect().bottom,
        ),
    )
    .toBe(true);
  await preview(page);
  await expect(page.locator(".print-page")).toHaveCount(1);
  expect((await flowMetrics(page, ".print-page .print-content")).image.height).toBe(160);
});

test("a picture inside a long column section wraps locally and keeps all columns on paper", async ({
  page,
}) => {
  const floating = JSON.stringify({
    horizontalPosition: { relative: "column", align: "center" },
    verticalPosition: { relative: "paragraph", offset: 40 * 9525 },
    wrap: { type: 1, side: "bothSides" },
  });
  const text = "FLOW narrow column with a floating picture and flowing body text. ".repeat(170);
  await open(page, [
    {
      type: "columnGroup",
      attrs: { count: 2, space: 48, balance: false },
      content: [
        {
          ...paragraph(text),
          content: [
            {
              type: "inlineImage",
              attrs: { src: pixel, width: 80, height: 120, floating },
            },
            { type: "text", text },
          ],
        },
      ],
    },
  ]);
  await expect
    .poll(async () => (await flowMetrics(page, ".document-stack .tiptap")).both)
    .toBeGreaterThan(2);
  expect((await flowMetrics(page, ".document-stack .tiptap")).overlap).toBe(0);
  expect(
    await page.locator(".document-columns").evaluate((g) => g.scrollWidth <= g.clientWidth + 2),
  ).toBe(true);
  await preview(page);
  expect((await page.locator(".print-content").allTextContents()).join("")).toBe(text);
  expect((await flowMetrics(page, ".print-page:first-of-type .print-content")).overlap).toBe(0);
  expect(
    await page
      .locator(".print-content [data-column-count]")
      .evaluateAll((gs) => gs.every((g) => g.scrollWidth <= g.clientWidth + 2)),
  ).toBe(true);
});

for (const textAlign of ["center", "right", "justify"])
  test(`${textAlign} paragraphs retain clear text around a centered picture`, async ({ page }) => {
    const text = "FLOW alignment and mixed font sizes near the picture. ".repeat(28);
    await open(page, [
      {
        ...paragraph(text),
        attrs: { textAlign, spaceAfter: "0px" },
        content: [
          image(),
          {
            type: "text",
            text,
            marks: [{ type: "textStyle", attrs: { fontSize: "18pt" } }],
          },
        ],
      },
    ]);
    await expect
      .poll(async () => (await flowMetrics(page, ".document-stack .tiptap")).spacers)
      .toBeGreaterThan(0);
    expect((await flowMetrics(page, ".document-stack .tiptap")).overlap).toBe(0);
    await preview(page);
    expect((await flowMetrics(page, ".print-page:first-of-type .print-content")).overlap).toBe(0);
  });

test("floating pictures use the margins of their own portrait and landscape sections", async ({
  page,
}) => {
  const landscape = { ...pageSettings, orientation: "landscape" };
  const text = "FLOW section-specific centered image with surrounding words. ".repeat(24);
  const body = [{ ...paragraph(text), content: [image(), { type: "text", text }] }];
  await open(page, [
    {
      type: "sectionContent",
      attrs: { page: JSON.stringify(pageSettings) },
      content: body,
    },
    { type: "sectionBreak", attrs: { page: JSON.stringify(landscape) } },
    {
      type: "sectionContent",
      attrs: { page: JSON.stringify(landscape) },
      content: body,
    },
  ]);
  await expect(page.locator(".document-section-content")).toHaveCount(2);
  await expect
    .poll(async () =>
      page.locator(".section-editable").evaluateAll((roots) =>
        roots.map((root) => {
          const p = root.querySelector("p")!.getBoundingClientRect(),
            img = root.querySelector("img[data-image-floating]")!.getBoundingClientRect();
          return Math.round((img.left + img.right - p.left - p.right) / 2);
        }),
      ),
    )
    .toEqual([0, 0]);
  for (let i = 0; i < 2; i++)
    expect(
      (
        await flowMetrics(
          page,
          `.document-section-content:nth-of-type(${i === 0 ? 1 : 3}) .section-editable`,
        )
      ).overlap,
    ).toBe(0);
  await preview(page);
  await expect(page.locator(".print-page")).toHaveCount(2);
  for (let i = 0; i < 2; i++)
    expect(
      (await flowMetrics(page, `.print-page:nth-of-type(${i + 1}) .print-content`)).overlap,
    ).toBe(0);
});

test("a short picture anchor near the bottom moves with its picture to the next page", async ({
  page,
}) => {
  await open(page, [
    ...Array.from({ length: 44 }, (_, i) => paragraph(`FLOW preceding line ${i + 1}`)),
    {
      ...paragraph("FLOW anchor at bottom"),
      content: [image(), { type: "text", text: "FLOW anchor at bottom" }],
    },
  ]);
  await expect
    .poll(async () =>
      page.locator(".document-stack .tiptap").evaluate((root) => {
        const img = root.querySelector("img[data-image-floating]")!.getBoundingClientRect(),
          sheet = root.closest<HTMLElement>(".document-stack")!,
          sr = sheet.getBoundingClientRect(),
          h = Number(sheet.dataset.pageHeight),
          top = Number(sheet.dataset.margin),
          bottom = Number(sheet.dataset.bottomMargin),
          i = Math.floor((img.top - sr.top - top + 0.5) / (h + 24));
        return i >= 1 && img.bottom - sr.top <= i * (h + 24) + h - bottom + 0.5;
      }),
    )
    .toBe(true);
  await expect(page.locator(".document-stack")).toHaveAttribute("data-pages", "2");
  await preview(page);
  await expect(page.locator(".print-page")).toHaveCount(2);
  expect(await page.locator(".print-page").nth(1).locator("img[data-image-floating]").count()).toBe(
    1,
  );
});
