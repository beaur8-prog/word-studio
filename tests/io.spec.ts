import { test, expect } from "@playwright/test";

test.beforeEach(async ({ page }) => {
  // A small same-origin harness avoids UI state and hot-reload navigation during IO checks.
  await page.route("**/__io_test__", (route) =>
    route.fulfill({
      contentType: "text/html",
      body: "<!doctype html><html><body>Document IO test</body></html>",
    }),
  );
  await page.goto("/__io_test__");
});

test("native files preserve the entire document and reject malformed structure", async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const io = await import("/src/document-io.ts");
    const source = {
      format: "word-studio",
      version: 1,
      title: "한국어 서식 문서",
      content: {
        type: "doc",
        content: [
          {
            type: "heading",
            attrs: {
              level: 2,
              textAlign: "center",
              indent: 0,
              lineSpacing: null,
            },
            content: [
              { type: "text", text: "한글 제목", marks: [{ type: "bold" }] },
            ],
          },
          {
            type: "paragraph",
            content: [
              {
                type: "text",
                text: "검토할 본문",
                marks: [
                  {
                    type: "textStyle",
                    attrs: {
                      fontFamily: "맑은 고딕",
                      fontSize: "14pt",
                      color: "#2454a6",
                    },
                  },
                  { type: "highlight", attrs: { color: "#ffee99" } },
                ],
              },
              { type: "hardBreak", marks: [{ type: "bold" }] },
              { type: "text", text: "다음 줄" },
            ],
          },
          { type: "pageBreak" },
          {
            type: "table",
            content: [
              {
                type: "tableRow",
                content: [
                  {
                    type: "tableCell",
                    attrs: { colspan: 1, rowspan: 1, colwidth: [220] },
                    content: [
                      {
                        type: "paragraph",
                        content: [{ type: "text", text: "표 내용" }],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
      page: {
        size: "Letter",
        orientation: "landscape",
        margin: 20,
        header: "머리글",
        footer: "바닥글",
        pageNumbers: true,
      },
      comments: [
        {
          id: "comment-one",
          quote: "검토할 본문",
          text: "메모를 보존",
          createdAt: "2026-10-17T09:00:00.000Z",
          resolved: false,
        },
      ],
      updatedAt: "2026-10-17T10:00:00.000Z",
    };
    const loaded = await io.importFile(
      new File([JSON.stringify(source)], "sample.wordstudio"),
    );
    const failures: string[] = [];
    const malformed = [
      { ...source, version: 2 },
      { ...source, page: { ...source.page, margin: -5 } },
      {
        ...source,
        content: {
          type: "doc",
          content: [{ type: "text", text: "top-level text" }],
        },
      },
      {
        ...source,
        content: {
          type: "doc",
          content: [
            {
              type: "image",
              attrs: { src: "https://remote.invalid/pixel.png" },
            },
          ],
        },
      },
      {
        ...source,
        content: {
          type: "doc",
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "text",
                  text: "unsafe",
                  marks: [
                    { type: "link", attrs: { href: "javascript:alert(1)" } },
                  ],
                },
              ],
            },
          ],
        },
      },
    ];
    for (const bad of malformed) {
      try {
        await io.importFile(new File([JSON.stringify(bad)], "bad.wordstudio"));
        failures.push("accepted");
      } catch {
        failures.push("rejected");
      }
    }
    const plain = await io.importFile(
      new File(["첫 줄\r\n\r\n마지막 <태그>"], "plain.txt"),
    );
    return {
      same: JSON.stringify(loaded.document) === JSON.stringify(source),
      warnings: loaded.warnings,
      failures,
      plain: plain.document.content,
    };
  });
  expect(result.same).toBe(true);
  expect(result.warnings).toEqual([]);
  expect(result.failures).toEqual(Array(5).fill("rejected"));
  expect(result.plain.content).toHaveLength(3);
  expect(result.plain.content[2].content[0].text).toBe("마지막 <태그>");
});

test("HTML import strips executable content and external resources while retaining safe formatting", async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const io = await import("/src/document-io.ts");
    const html =
      '<script>window.__unsafe=true</script><p style="color:rgb(20, 30, 40);background-image:url(https://remote.invalid/image);font-size:18pt" onclick="alert(1)">안전한 <strong>굵게</strong><a href="javascript:alert(1)">위험 링크</a><a href="https://example.com">일반 링크</a><img src="https://remote.invalid/track" alt="원격 그림"><iframe src="https://remote.invalid"></iframe></p>';
    const result = await io.importFile(new File([html], "test.html"));
    return { html: result.html, warnings: result.warnings };
  });
  expect(result.html).toContain("<strong>굵게</strong>");
  expect(result.html).toContain("font-size:18pt");
  expect(result.html).toContain('href="https://example.com"');
  expect(result.html).not.toMatch(
    /<script|onclick|javascript:|<iframe|remote\.invalid|background-image/,
  );
  expect(result.warnings).toHaveLength(1);
});

test("DOCX export contains formatted text, lists, tables, image, page breaks and page setup; import recovers body", async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const io = await import("/src/document-io.ts");
    const { default: JSZip } = await import(
      "/node_modules/.vite/deps/jszip.js"
    );
    const canvas = document.createElement("canvas");
    canvas.width = 20;
    canvas.height = 10;
    canvas.getContext("2d")!.fillRect(0, 0, 20, 10);
    const source = {
      format: "word-studio",
      version: 1,
      title: "DOCX 검증",
      content: {
        type: "doc",
        content: [
          {
            type: "heading",
            attrs: { level: 1, textAlign: "center" },
            content: [{ type: "text", text: "테스트 제목" }],
          },
          {
            type: "paragraph",
            attrs: { lineSpacing: "1.5", indent: 24 },
            content: [
              {
                type: "text",
                text: "굵은 본문",
                marks: [
                  { type: "bold" },
                  { type: "italic" },
                  { type: "underline" },
                  {
                    type: "textStyle",
                    attrs: {
                      fontSize: "16pt",
                      fontFamily: "Arial",
                      color: "#2454a6",
                    },
                  },
                ],
              },
            ],
          },
          {
            type: "orderedList",
            attrs: { start: 3 },
            content: [
              {
                type: "listItem",
                content: [
                  {
                    type: "paragraph",
                    content: [{ type: "text", text: "목록 항목" }],
                  },
                ],
              },
            ],
          },
          {
            type: "table",
            content: [
              {
                type: "tableRow",
                content: [
                  {
                    type: "tableHeader",
                    attrs: { colspan: 1, rowspan: 1, colwidth: [200] },
                    content: [
                      {
                        type: "paragraph",
                        content: [{ type: "text", text: "표의 셀" }],
                      },
                    ],
                  },
                ],
              },
            ],
          },
          {
            type: "image",
            attrs: {
              src: canvas.toDataURL("image/png"),
              width: 20,
              height: 10,
              alt: "검증 그림",
            },
          },
          { type: "pageBreak" },
          {
            type: "paragraph",
            content: [{ type: "text", text: "다음 페이지" }],
          },
        ],
      },
      page: {
        size: "Letter",
        orientation: "landscape",
        margin: 20,
        header: "내 머리글",
        footer: "내 바닥글",
        pageNumbers: true,
      },
      comments: [],
      updatedAt: new Date().toISOString(),
    };
    const blob = await io.exportDocx(source);
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const body = await zip.file("word/document.xml")!.async("string");
    const numbering = await zip.file("word/numbering.xml")!.async("string");
    const footer = await zip.file("word/footer1.xml")!.async("string");
    const imported = await io.importFile(new File([blob], "roundtrip.docx"));
    return {
      size: blob.size,
      body,
      numbering,
      footer,
      media: Object.keys(zip.files).filter(
        (p) => p.startsWith("word/media/") && !zip.files[p].dir,
      ).length,
      imported,
    };
  });
  expect(result.size).toBeGreaterThan(5000);
  expect(result.body).toContain("테스트 제목");
  expect(result.body).toMatch(/<w:b\b/);
  expect(result.body).toMatch(/<w:i\b/);
  expect(result.body).toContain('w:val="2454A6"');
  expect(result.body).toContain('w:val="32"');
  expect(result.body).toContain('w:line="360"');
  expect(result.body).toContain('w:left="360"');
  expect(result.body).toContain("<w:tbl>");
  expect(result.body).toContain('w:type="page"');
  expect(result.body).toContain('w:orient="landscape"');
  expect(result.body).toContain('w:w="15840"');
  expect(result.numbering).toContain('w:val="3"');
  expect(result.footer).toContain("PAGE");
  expect(result.media).toBe(1);
  expect(result.imported.html).toContain("테스트 제목");
  expect(result.imported.html).toContain("굵은 본문");
  expect(result.imported.html).toContain("<strong>");
  expect(result.imported.html).toMatch(/<table[\s>]/);
  expect(result.imported.html).toContain("data:image/png;base64,");
  expect(result.imported.document.page).toEqual({
    size: "Letter",
    orientation: "landscape",
    margin: 20,
    header: "내 머리글",
    footer: "내 바닥글",
    pageNumbers: true,
  });
  expect(result.imported.warnings.length).toBeGreaterThan(0);
});
