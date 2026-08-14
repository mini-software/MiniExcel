import { expect, test } from "@playwright/test";

for (const project of ["desktop", "mobile"]) {
  test(`${project} renders the generated workbook`, async ({ page }, testInfo) => {
    test.skip(testInfo.project.name !== project);
    await page.goto("/");

    await expect(page.getByTestId("runtime-status")).toContainText("WASM");
    await expect(page.getByTestId("file-name")).toHaveText("miniexcel-browser-demo.xlsx");
    await expect(page.getByRole("cell", { name: "MiniExcel", exact: true })).toBeVisible();
    await expect(page.getByRole("cell", { name: "Browser WASM", exact: true })).toBeVisible();
    await expect(page.locator("#previewTable tbody tr")).toHaveCount(2);

    await page.screenshot({
      path: testInfo.outputPath(`${project}.png`),
      fullPage: true,
    });
  });
}

test("query controls refresh the preview", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("file-name")).toHaveText("miniexcel-browser-demo.xlsx");

  await page.getByLabel("Header row").uncheck();
  await page.getByRole("button", { name: "Refresh preview" }).click();

  await expect(page.locator("#previewTable thead th").filter({ hasText: /^A$/ })).toBeVisible();
  await expect(page.locator("#previewTable tbody tr")).toHaveCount(3);
});

test("end cell limits the preview range", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("file-name")).toHaveText("miniexcel-browser-demo.xlsx");

  await page.getByLabel("End cell").fill("B2");
  await page.getByRole("button", { name: "Refresh preview" }).click();

  await expect(page.locator("#previewTable thead th")).toHaveCount(3);
  await expect(page.locator("#previewTable tbody tr")).toHaveCount(1);
  await expect(page.getByRole("cell", { name: "MiniExcel", exact: true })).toBeVisible();
});
