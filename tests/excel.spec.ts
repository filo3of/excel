import { test, expect } from "@playwright/test";

const moment = require("moment");

let now = moment().format("L");

test("open excel sheet and add today date", async ({ page }) => {
  test.setTimeout(50000);

  await page.goto(
    "https://onedrive.live.com/edit?id=702781C0D1F0D235!107&resid=702781C0D1F0D235!107&ithint=file%2cxlsx&ct=1715877447058&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=945b64ad-9ad2-4b28-bbf1-ce2b65299c34&wdo=2&cid=702781c0d1f0d235"
  );

  await page.waitForURL(
    "https://onedrive.live.com/edit?id=702781C0D1F0D235!107&resid=702781C0D1F0D235!107&ithint=file%2cxlsx&ct=1715877447058&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=945b64ad-9ad2-4b28-bbf1-ce2b65299c34&wdo=2&cid=702781c0d1f0d235"
  );

  await expect(
    page
      .frameLocator('iframe[name="WacFrame_Excel_0"]')
      .getByRole("button", { name: "this is new test Saved to" })
  ).toBeVisible();

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .clear();

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .fill("B4");

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .press("Enter");

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//div[@id="formulaBarTextDivId_textElement"]')
    .fill("=TODAY()");

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//div[@id="formulaBarTextDivId_textElement"]')
    .press("Enter");

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .clear();

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .fill("B4");

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .press("Enter");

  await expect(
    page
      .frameLocator('iframe[name="WacFrame_Excel_0"]')
      .locator('//div[@id="m_excelWebRenderer_ewaCtl_selectionHighlight0-1-0"]')
  ).toHaveText(now);
});
