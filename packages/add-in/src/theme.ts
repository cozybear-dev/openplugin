import { createLightTheme, type BrandVariants, type Theme } from "@fluentui/react-components";

// Margin: paper surfaces, plum ink, and brass details. Keep portal controls in
// the same theme as the task pane, including focus, selected, and hover states.
const plum: BrandVariants = {
  10: "#170e14", 20: "#291b24", 30: "#392532", 40: "#483040",
  50: "#573b4e", 60: "#67465d", 70: "#77526c", 80: "#875e7b",
  90: "#996d8c", 100: "#aa809d", 110: "#ba94af", 120: "#cba9c0",
  130: "#dbc0d2", 140: "#e7d5e0", 150: "#f1e7ed", 160: "#f9f4f7"
};
export const marginTheme: Theme = {
  ...createLightTheme(plum),
  fontFamilyBase: '"Segoe UI", sans-serif',
  colorBrandBackground: plum[40],
  colorBrandBackgroundHover: plum[30],
  colorBrandBackgroundPressed: plum[20],
  colorBrandForeground1: plum[40],
  colorBrandForeground2: plum[40],
  colorNeutralBackground1: "#fffdf8",
  colorNeutralBackground1Hover: "#f2eee6",
  colorNeutralBackground1Pressed: "#e9e3d8",
  colorNeutralBackground2: "#f5f1e9",
  colorNeutralBackground3: "#eee8dd",
  colorNeutralForeground1: "#302831",
  colorNeutralForeground2: "#625862",
  colorNeutralForeground3: "#766c72",
  colorNeutralStroke1: "#cec5b9",
  colorNeutralStroke2: "#e2dbcf",
  borderRadiusSmall: "3px",
  borderRadiusMedium: "5px",
  borderRadiusLarge: "8px"
};
