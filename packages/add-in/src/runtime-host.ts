import {
  FakeExcelHost,
  FakePowerPointHost,
  FakeWordHost,
  type HostAdapter,
  type HostKind
} from "@openplugin/core";
import { ExcelHost } from "@openplugin/host-excel";
import { PowerPointHost } from "@openplugin/host-powerpoint";
import { WordHost } from "@openplugin/host-word";

export function hostKindFromOffice(host?: Office.HostType): HostKind {
  if (host === Office.HostType.Excel) return "excel";
  if (host === Office.HostType.Word) return "word";
  if (host === Office.HostType.PowerPoint) return "powerpoint";
  return "excel";
}

export function createHost(kind: HostKind, inOffice: boolean): HostAdapter {
  if (inOffice) {
    if (kind === "excel") return new ExcelHost();
    if (kind === "word") return new WordHost();
    return new PowerPointHost();
  }
  if (kind === "word") return new FakeWordHost();
  if (kind === "powerpoint") return new FakePowerPointHost();
  return new FakeExcelHost();
}
