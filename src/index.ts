import { createMockEnvironment } from "./setup.js";

const env = createMockEnvironment();

globalThis.Excel = env.excel as any;
globalThis.Office = env.office as any;
globalThis.CustomFunctions = env.customFunctions as any;
globalThis.MockOfficeJs = env.mockOfficeJs;
