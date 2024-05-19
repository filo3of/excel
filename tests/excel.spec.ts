import { test, expect } from "@playwright/test";

const moment = require("moment");

let now = moment().format("DD/MM/YYYY");

test("open excel sheet and add today date", async ({ page, context }) => {
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

  await page.waitForTimeout(20000);

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .click();

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .clear();

  await page.waitForTimeout(2000);

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .fill("B4");

  await page.waitForTimeout(200);

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .press("Enter");

  await page.waitForTimeout(2000);

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//div[@id="formulaBarTextDivId_textElement"]')
    .click();

  await page.waitForTimeout(2000);

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//div[@id="formulaBarTextDivId_textElement"]')
    .pressSequentially("=TODAY()", { delay: 100 });

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//div[@id="formulaBarTextDivId_textElement"]')
    .press("Enter");

  await page.waitForTimeout(2000);

  const myElement = page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .getByLabel("Got it");

  if (await myElement.isVisible()) {
    await myElement.click();

    await page.waitForTimeout(2000);

    await page
      .frameLocator('iframe[name="WacFrame_Excel_0"]')
      .locator('//div[@id="formulaBarTextDivId_textElement"]')
      .click();

    await page.waitForTimeout(200);
  }

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .click();

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .clear();

  await page.waitForTimeout(2000);

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .fill("B4");

  await page.waitForTimeout(200);

  await page
    .frameLocator('iframe[name="WacFrame_Excel_0"]')
    .locator('//input[@id="FormulaBar-NameBox-input"]')
    .press("Enter");

  await page.waitForTimeout(2000);

  await expect(
    page
      .frameLocator('iframe[name="WacFrame_Excel_0"]')
      .locator('//div[@id="formulaBarTextDivId_textElement"]//div')
  ).toHaveText("=TODAY()");

  await expect(
    page
      .frameLocator('iframe[name="WacFrame_Excel_0"]')
      .locator(
        '(//label[@aria-label="' + now + ' . B4 . Contains Formula . "])[1]'
      )
  ).toBeHidden();

  await expect(
    page
      .frameLocator('iframe[name="WacFrame_Excel_0"]')
      .locator(
        '(//label[@aria-label="' + now + ' . B4 . Contains Formula . "])[2]'
      )
  ).toHaveAttribute("aria-label", now + " . B4 . Contains Formula . ");
});
