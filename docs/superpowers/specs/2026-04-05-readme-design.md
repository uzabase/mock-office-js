# README.md Design Spec

## Overview

Create a compact, user-facing README.md for the `office-js-mock` npm package. The target audience is Office Add-in developers who want to test their add-ins without the Excel host.

## Structure

1. **Title + one-line description**
2. **Warning blockquote** — early development notice with link to Limitations section
3. **Install** — `npm install office-js-mock`
4. **Setup** — ExcelMock instantiation and global registration (Vitest example)
5. **Usage** — one basic example: register custom function, set formula, assert value
6. **API** — ExcelMock method table (setup / assertion / reset)
7. **Limitations** — bulleted list of unsupported features
8. **License** — MIT

## Design Decisions

- **Compact over comprehensive:** lead with setup + one example so users can get started quickly.
- **Warning banner at top:** GitHub-flavored `> [!WARNING]` blockquote to set expectations about early development status, linking to Limitations section for specifics.
- **Limitations section kept:** the warning banner gives a high-level heads-up; the section gives concrete details (no native Excel functions, no events, no streaming, no cell reference resolution in formulas, no UI rendering).
- **Single usage example:** the custom function flow (associate → setCell with formula → getCell assertion) demonstrates the core value proposition in minimal code.
- **API table:** flat table of all ExcelMock methods grouped by purpose (setup, assertion, reset). Includes `setCells`, `getCells`, `setActiveWorksheet`, `addWorksheet`.
- **CellState type:** briefly document the shape (`{ value, formula?, spilledFrom? }`) since `getCell()` returns it and the usage example references `.value`.
- **setCell return type:** note that it returns `void` for plain values and `Promise<void>` for formulas.
- **TypeScript examples:** all code examples use TypeScript, matching the package's TypeScript-first nature.
- **No badges or repo URLs:** omitted for compactness; can be added later when publishing.
- **Framework-agnostic note:** Setup example uses Vitest but add a brief note that any test framework works.
- **English language:** standard for npm packages.
