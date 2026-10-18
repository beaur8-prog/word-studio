"use strict";
const { expect } = require("@playwright/test");
const fs = require("node:fs/promises");
const path = require("node:path");
const assert = require("node:assert/strict");

/** Exercise the installed Electron/Chromium engine, rather than the browser
 * suite's newer Chrome. Use the caller's isolated profile and native PDF IPC. */
module.exports = async function layoutDesktopSmoke(app, window, directory) {
  const page = {
    size: "A4",
    orientation: "portrait",
    margin: 25.4,
    header: "",
    footer: "",
    pageNumbers: false,
  };
  const paragraph = (text) => ({
    type: "paragraph",
    attrs: { lineSpacing: "20px", spaceAfter: "0px" },
    content: [{ type: "text", text }],
  });
  const text = "FLOW 한글과 English document layout 0123456789. ".repeat(180);
  const pixel =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aN1sAAAAASUVORK5CYII=";
  const floating = JSON.stringify({
    horizontalPosition: { relative: "margin", align: "center" },
    verticalPosition: { relative: "paragraph", offset: 40 * 9525 },
    wrap: { type: 1, side: "bothSides" },
  });
  const body = "FLOW native text around both sides of a centered picture. ".repeat(35);
  const scenarios = [
    {
      name: "columns",
      text: "FLOW preceding" + text + "FLOW following",
      content: [
        paragraph("FLOW preceding"),
        {
          type: "columnGroup",
          attrs: { count: 2, space: 48, balance: false },
          content: [paragraph(text)],
        },
        paragraph("FLOW following"),
      ],
    },
    {
      name: "centered-picture",
      text: body,
      content: [
        {
          ...paragraph(body),
          content: [
            {
              type: "inlineImage",
              attrs: { src: pixel, width: 200, height: 160, floating },
            },
            { type: "text", text: body, marks: [{ type: "bold" }] },
          ],
        },
      ],
    },
  ];
  const results = [];
  for (const scenario of scenarios) {
    const fixture = path.join(directory, `${scenario.name}.wordstudio`),
      pdf = path.join(directory, `${scenario.name}.pdf`);
    await fs.writeFile(
      fixture,
      JSON.stringify({
        format: "word-studio",
        version: 1,
        title: scenario.name,
        updatedAt: new Date().toISOString(),
        page,
        comments: [],
        content: { type: "doc", content: scenario.content },
      }),
    );
    await app.evaluate(
      ({ dialog }, files) => {
        dialog.showOpenDialog = async () => ({
          canceled: false,
          filePaths: [files.fixture],
        });
        dialog.showSaveDialog = async () => ({
          canceled: false,
          filePath: files.pdf,
        });
      },
      { fixture, pdf },
    );
    await window.keyboard.press("ControlOrMeta+o");
    await expect(window.locator(".document-stack .tiptap")).toHaveText(scenario.text);
    await expect(window.getByText("이 PC에 저장됨", { exact: true })).toBeVisible();
    if (scenario.name === "columns") {
      await expect
        .poll(() => window.locator(".document-stack").getAttribute("data-pages"))
        .not.toBe("1");
      assert.equal(await window.evaluate(() => CSS.supports("column-wrap", "wrap")), true);
      assert.equal(
        await window
          .locator(".document-columns")
          .evaluate((g) => g.scrollWidth <= g.clientWidth + 2),
        true,
      );
    } else
      await expect
        .poll(() => window.locator(".document-stack [data-flow-widget]").count())
        .toBeGreaterThan(0);
    const editorPages = Number(await window.locator(".document-stack").getAttribute("data-pages"));
    await window.screenshot({
      path: path.join(directory, `${scenario.name}-editor.png`),
    });
    await window.getByRole("button", { name: "보기", exact: true }).click();
    await window.getByRole("button", { name: "인쇄 미리보기", exact: true }).click();
    await expect(window.locator(".print-page").first()).toBeVisible();
    const previewPages = await window.locator(".print-page").count();
    assert.equal(
      (await window.locator(".print-content").allTextContents()).join(""),
      scenario.text,
    );
    assert.equal(
      previewPages,
      editorPages,
      `${scenario.name}: editor and preview must use the same pages`,
    );
    let bothSides;
    if (scenario.name === "centered-picture") {
      bothSides = await window
        .locator(".print-content")
        .first()
        .evaluate((root) => {
          const image = root.querySelector("img").getBoundingClientRect(),
            lines = new Map(),
            walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
          let overlap = 0;
          while (walker.nextNode()) {
            const node = walker.currentNode;
            for (let i = 0; i < node.textContent.length; i++) {
              if (!node.textContent[i].trim()) continue;
              const range = document.createRange();
              range.setStart(node, i);
              range.setEnd(node, i + 1);
              const rect = range.getBoundingClientRect();
              if (!rect.width || rect.bottom <= image.top + 0.5 || rect.top >= image.bottom - 0.5)
                continue;
              const key = Math.round(rect.top),
                line = lines.get(key) || { left: false, right: false };
              if (rect.right <= image.left + 0.5) line.left = true;
              else if (rect.left >= image.right - 0.5) line.right = true;
              else overlap++;
              lines.set(key, line);
            }
          }
          return {
            lines: [...lines.values()].filter((line) => line.left && line.right).length,
            overlap,
          };
        });
      assert.ok(bothSides.lines > 2);
      assert.equal(bothSides.overlap, 0);
    }
    await window.screenshot({
      path: path.join(directory, `${scenario.name}-preview.png`),
    });
    await window.getByRole("button", { name: "PDF 저장", exact: true }).click();
    await expect(window.getByText("PDF를 저장했습니다", { exact: true })).toBeVisible();
    const bytes = await fs.readFile(pdf),
      pdfPages = (bytes.toString("latin1").match(/\/Type\s*\/Page\b/g) || []).length;
    assert.equal(pdfPages, previewPages);
    results.push({
      scenario: scenario.name,
      editorPages,
      previewPages,
      pdfPages,
      pdfBytes: bytes.length,
      ...(bothSides ? { bothSides } : {}),
    });
    await window.getByRole("button", { name: "편집으로 돌아가기", exact: true }).click();
  }
  return results;
};
