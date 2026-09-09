import { describe, expect, it } from "vitest";
import { buttonForBlockingDialog, isBlockingOfficeDialog } from "../src/dialogs.js";

describe("isBlockingOfficeDialog", () => {
  it("matches the WebView Stop On Load debugger prompt", () => {
    expect(isBlockingOfficeDialog("WebView Stop On Load")).toBe(true);
  });

  it("ignores the Office document window itself", () => {
    expect(isBlockingOfficeDialog("Book1 - Excel")).toBe(false);
    expect(isBlockingOfficeDialog("Document1 - Word")).toBe(false);
    expect(isBlockingOfficeDialog("Presentation1 - PowerPoint")).toBe(false);
  });
});

describe("buttonForBlockingDialog", () => {
  it("clicks Cancel on WebView Stop On Load so Office will not show it again", () => {
    expect(buttonForBlockingDialog("WebView Stop On Load")).toBe("Cancel");
  });
});
