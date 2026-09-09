/// <reference types="vite/client" />
/// <reference types="office-js" />

declare const OfficeRuntime: {
  storage: {
    getItem(key: string): Promise<string | null | undefined>;
    setItem(key: string, value: string): Promise<void>;
  };
};

declare module "*.md?raw" {
  const content: string;
  export default content;
}
