import { FluentProvider } from "@fluentui/react-components";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./functions";
import { hostKindFromOffice } from "./runtime-host";
import "./styles.css";
import { marginTheme } from "./theme";

const root = createRoot(document.getElementById("root")!);

function render(hostKind: ReturnType<typeof hostKindFromOffice>, inOffice: boolean) {
  root.render(
    <StrictMode>
      <FluentProvider className="op-provider" theme={marginTheme}>
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
