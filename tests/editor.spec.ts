import { test, expect } from "@playwright/test";

async function blank(page: any) {
  await page.goto("/");
  await page.locator(".tiptap").waitFor();
  await page.keyboard.press("Control+n");
  await page
    .getByRole("button", { name: "빈 문서 자유롭게 시작하세요" })
    .click();
  await expect(page.locator(".tiptap")).toHaveText("");
}

test("editing, formatting, undo, search/replace, and automatic restore", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await blank(page);
  const edit = page.locator(".tiptap");
  await edit.click();
  await page.keyboard.type("Hello Studio");
  await page.keyboard.press("ControlOrMeta+a");
  await page
    .getByRole("button", { name: "굵게 (Ctrl+B)", exact: true })
    .click();
  await expect(edit.locator("strong")).toHaveText("Hello Studio");
  await page.getByRole("button", { name: "실행 취소 (Ctrl+Z)" }).click();
  await expect(edit.locator("strong")).toHaveCount(0);
  await page.getByRole("button", { name: "다시 실행 (Ctrl+Y)" }).click();
  await expect(edit.locator("strong")).toHaveText("Hello Studio");
  const fontSize = page.getByRole("spinbutton", {
    name: "글꼴 크기",
    exact: true,
  });
  await fontSize.fill("");
  await fontSize.pressSequentially("17");
  await expect(fontSize).toHaveValue("17");
  await fontSize.press("Enter");
  await expect(edit.locator('span[style*="font-size: 17pt"]')).toContainText(
    "Hello Studio",
  );
  await page.getByRole("button", { name: "바꾸기", exact: true }).click();
  await page.getByRole("textbox", { name: "문서에서 찾기" }).fill("Studio");
  await expect(page.getByText("1개 결과", { exact: true })).toBeVisible();
  await page.getByRole("textbox", { name: "바꿀 내용" }).fill("World");
  await page.getByRole("button", { name: "모두 바꾸기", exact: true }).click();
  await expect(edit).toHaveText("Hello World");
  await page.getByRole("textbox", { name: "문서 제목" }).fill("테스트 문서");
  await expect(page.getByText("이 PC에 저장됨", { exact: true })).toBeVisible();
  await page.reload();
  await expect(page.getByRole("textbox", { name: "문서 제목" })).toHaveValue(
    "테스트 문서",
  );
  await expect(page.locator(".tiptap")).toHaveText("Hello World");
  expect(errors).toEqual([]);
});

test("table editing, comments, page settings, and native file save roundtrip", async ({
  page,
}) => {
  await blank(page);
  await page.locator(".tiptap").click();
  await page.keyboard.type("A document");
  await page.getByRole("button", { name: "삽입", exact: true }).click();
  await page.getByRole("button", { name: "표", exact: true }).click();
  await page.getByRole("button", { name: "2행 3열 표", exact: true }).click();
  await expect(page.locator(".tiptap table tr")).toHaveCount(2);
  await page.getByRole("button", { name: "표 레이아웃", exact: true }).click();
  await page.getByRole("button", { name: "아래에 삽입", exact: true }).click();
  await expect(page.locator(".tiptap table tr")).toHaveCount(3);
  await page.getByRole("button", { name: "검토", exact: true }).click();
  await page.getByRole("button", { name: "새 메모", exact: true }).click();
  await page.getByRole("textbox", { name: "메모 내용" }).fill("내용 확인 완료");
  await page.getByRole("button", { name: "메모 추가", exact: true }).click();
  await expect(page.locator(".comment-card")).toContainText("내용 확인 완료");
  await page.getByRole("button", { name: "해결", exact: true }).click();
  await expect(page.locator(".comment-card")).toHaveClass(/resolved/);
  await page.getByRole("button", { name: "레이아웃", exact: true }).click();
  await page.getByRole("button", { name: "용지 방향", exact: true }).click();
  await page.getByRole("button", { name: "가로", exact: true }).click();
  await expect(page.locator(".document-stack")).toHaveAttribute(
    "data-page-height",
    "793.7",
  );
  const wait = page.waitForEvent("download");
  await page.keyboard.press("Control+s");
  const d = await wait;
  expect(d.suggestedFilename()).toBe("문서1.wordstudio");
  const path = await d.path();
  expect(path).toBeTruthy();
  await page.locator("input[type=file]").first().setInputFiles(path!);
  await expect(page.locator(".tiptap table tr")).toHaveCount(3);
  await expect(page.locator(".comment-card")).toContainText("내용 확인 완료");
});

test("multi-page flow and explicit page breaks keep text inside page margins", async ({
  page,
}) => {
  await blank(page);
  const paragraphs = Array.from(
    { length: 65 },
    (_, i) =>
      `<p>${i + 1}. This paragraph verifies page flow and margin placement.</p>`,
  ).join("");
  await page
    .locator("input[type=file]")
    .first()
    .setInputFiles({
      name: "pagination.html",
      mimeType: "text/html",
      buffer: Buffer.from(paragraphs),
    });
  await expect
    .poll(async () =>
      Number(await page.locator(".document-stack").getAttribute("data-pages")),
    )
    .toBeGreaterThan(2);
  await expect(page.locator(".pagination-gap")).toHaveCount(2);
  const violation = await page.evaluate(() => {
    const stack = document.querySelector(".document-stack") as HTMLElement;
    const root = stack.getBoundingClientRect().top;
    const h = Number(stack.dataset.pageHeight),
      m = Number(stack.dataset.margin);
    return [...document.querySelectorAll(".tiptap>p")]
      .filter((el) => {
        const r = el.getBoundingClientRect();
        const top = r.top - root;
        const pos = top % (h + 24);
        return pos < m - 2 || pos + r.height > h - m + 2;
      })
      .map((el) => el.textContent);
  });
  expect(violation).toEqual([]);
  await page.locator(".tiptap").click();
  await page.keyboard.press(process.platform === "darwin" ? "Meta+ArrowDown" : "Control+End");
  await page.getByRole("button", { name: "삽입", exact: true }).click();
  await page
    .getByRole("button", { name: "페이지 나누기", exact: true })
    .click();
  await expect(page.locator(".page-break")).toHaveCount(1);
  await expect(page.locator(".tiptap")).toBeFocused();
  await page.keyboard.type("Next page");
  await expect(page.locator(".tiptap")).toContainText("Next page");
});

test("all ribbons, light theme, navigation, read mode and image insertion render", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto("/");
  await page.locator(".tiptap").waitFor();
  for (const tab of ["삽입", "디자인", "레이아웃", "검토", "보기", "홈"]) {
    await page.getByRole("button", { name: tab, exact: true }).click();
    await expect(
      page.getByRole("toolbar", { name: `${tab} 도구` }),
    ).toBeVisible();
  }
  await page
    .getByRole("button", { name: "화면 테마 전환", exact: true })
    .click();
  await expect(page.locator("html")).toHaveAttribute("data-chrome", "light");
  await page.getByRole("button", { name: "보기", exact: true }).click();
  await page.getByRole("button", { name: "탐색 창", exact: true }).click();
  await expect(page.locator(".outline button")).toHaveCount(4);
  await page.getByRole("button", { name: "읽기 모드", exact: true }).click();
  await expect(page.locator(".tiptap")).toHaveAttribute(
    "contenteditable",
    "false",
  );
  await page.getByRole("button", { name: "편집으로 돌아가기" }).click();
  await expect(page.locator(".tiptap")).toHaveAttribute(
    "contenteditable",
    "true",
  );
  await page.locator(".tiptap").click();
  await page.keyboard.press(process.platform === "darwin" ? "Meta+ArrowDown" : "Control+End");
  await page
    .locator("input[type=file]")
    .nth(1)
    .setInputFiles({
      name: "pixel.png",
      mimeType: "image/png",
      buffer: Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aN1sAAAAASUVORK5CYII=",
        "base64",
      ),
    });
  await expect(page.locator(".tiptap img")).toHaveCount(1);
  expect(errors).toEqual([]);
  await page.setViewportSize({ width: 1100, height: 800 });
  await page.screenshot({ path: "analysis/compact-ui.png" });
});
