import { test as setup, expect } from "@playwright/test";

const authFile = "playwright/.auth/user.json";

const email = process.env.EMAIL;

const password = process.env.PASSWORD;

if (!email) {
  throw new Error("Unexpected error: Missing email");
}

if (!password) {
  throw new Error("Unexpected error: Missing password");
}

setup("authenticate", async ({ page }) => {
  await page.goto("https://login.microsoftonline.com/");

  await expect(page.getByRole("heading", { name: "Sign in" })).toBeVisible();

  await expect(
    page.getByPlaceholder("Email, phone, or Skype", { exact: true })
  ).toBeVisible();

  await page
    .getByPlaceholder("Email, phone, or Skype", { exact: true })
    .fill(email);

  await expect(page.getByRole("button", { name: "Next" })).toBeVisible();

  await expect(page.getByRole("button", { name: "Next" })).toBeEnabled();

  await page.getByRole("button", { name: "Next" }).click();

  await expect(
    page.getByRole("heading", { name: "Enter password" })
  ).toBeVisible();

  await expect(
    page.getByPlaceholder("Password", { exact: true })
  ).toBeVisible();

  await page.getByPlaceholder("Password", { exact: true }).fill(password);

  await expect(page.getByRole("button", { name: "Sign in" })).toBeVisible();

  await page.getByRole("button", { name: "Sign in" }).click();

  await expect(
    page.getByRole("heading", { name: "Stay signed in?" })
  ).toBeVisible();

  await expect(page.getByRole("button", { name: "No" })).toBeVisible();

  await expect(page.getByRole("button", { name: "No" })).toBeEnabled();

  await page.getByRole("button", { name: "No" }).click();

  await page.waitForURL("https://www.office.com/?auth=1");

  await expect(
    page.getByRole("heading", { name: "Welcome to Microsoft" })
  ).toBeVisible();

  // End of authentication steps.

  await page.context().storageState({ path: authFile });
});
