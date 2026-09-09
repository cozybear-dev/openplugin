import { describe, expect, it } from "vitest";
import { disconnectCdp, pickTaskpaneTarget } from "../src/cdp.js";

describe("pickTaskpaneTarget", () => {
  it("prefers the taskpane.html page over other localhost targets", () => {
    const picked = pickTaskpaneTarget([
      { id: "1", type: "page", url: "https://localhost:3000/commands.html", title: "commands" },
      { id: "2", type: "page", url: "https://localhost:3000/taskpane.html", title: "OpenPlugin" },
      { id: "3", type: "iframe", url: "https://localhost:3000/taskpane.html", title: "nested" }
    ]);
    expect(picked?.id).toBe("2");
  });

  it("falls back to any localhost:3000 page", () => {
    const picked = pickTaskpaneTarget([
      { id: "1", type: "page", url: "https://excel.officeapps.live.com/", title: "Excel" },
      { id: "2", type: "page", url: "https://localhost:3000/functions.js", title: "fn" }
    ]);
    expect(picked?.id).toBe("2");
  });

  it("returns undefined when Office has no add-in page", () => {
    expect(
      pickTaskpaneTarget([{ id: "1", type: "page", url: "https://office.com", title: "Office" }])
    ).toBeUndefined();
  });
});

describe("disconnectCdp", () => {
  it("does not call browser.close", async () => {
    const browser = {
      close: async () => {
        throw new Error("close must not be called — it kills Office WebView2");
      },
      _connection: {
        closeCalls: 0,
        close() {
          this.closeCalls++;
        }
      }
    };
    await disconnectCdp(browser as never);
    expect(browser._connection.closeCalls).toBe(0);
  });

  it("closes the nested websocket transport", async () => {
    let wsClosed = 0;
    const browser = {
      close: async () => {
        throw new Error("close must not be called — it kills Office WebView2");
      },
      _connection: {
        close() {},
        _transport: {
          close() {},
          _ws: {
            close() {
              wsClosed++;
            }
          }
        }
      }
    };
    await disconnectCdp(browser as never);
    expect(wsClosed).toBe(1);
  });
});
