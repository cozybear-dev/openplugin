import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { hostKindFromOffice } from "./runtime-host";
import "./styles.css";

const root = createRoot(document.getElementById("root")!);

function render(hostKind: ReturnType<typeof hostKindFromOffice>, inOffice: boolean) {
  root.render(
    <StrictMode>
      <FluentProvider theme={webLightTheme}>
        <App hostKind={hostKind} inOffice={inOffice} />
      </FluentProvider>
    </StrictMode>
  );
}

const office = (globalThis as { Office?: typeof Office }).Office;
if (office?.onReady) {
  office.onReady((info) => {
    render(hostKindFromOffice(info.host), true);
  });
} else {
  render("excel", false);
}
