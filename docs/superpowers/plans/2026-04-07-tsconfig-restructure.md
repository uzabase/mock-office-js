# tsconfig Restructure Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Separate tsconfig.json into a project-wide type-checking config and a build-only config for tsdown dts generation.

**Architecture:** `tsconfig.json` becomes type-check-only (`noEmit: true`) covering the full project. `tsconfig.build.json` extends it and adds build settings for tsdown. `tsconfig.test-d.json` is cleaned up to remove redundant overrides.

**Tech Stack:** TypeScript, tsdown

**Spec:** `docs/superpowers/specs/2026-04-07-tsconfig-restructure-design.md`

---

### Task 1: Create `tsconfig.build.json` and update tsdown reference

**Files:**
- Create: `tsconfig.build.json`
- Modify: `tsdown.config.ts:8`

- [ ] **Step 1: Create `tsconfig.build.json`**

This preserves the current build settings for tsdown's dts generation, extending the base config.

```jsonc
{
  "extends": "./tsconfig.json",
  "compilerOptions": {
    "noEmit": false,
    "declaration": true,
    "outDir": "dist",
    "rootDir": "src"
  },
  "include": ["src/**/*.ts"]
}
```

Note: At this point `tsconfig.json` still has the old settings, so `extends` will inherit them. The `noEmit: false` override is not yet needed but won't cause issues since the current config doesn't set `noEmit`.

- [ ] **Step 2: Update `tsdown.config.ts` to reference `tsconfig.build.json`**

In `tsdown.config.ts`, change the dts tsconfig reference:

```ts
// Before
dts: { tsconfig: "./tsconfig.json" },

// After
dts: { tsconfig: "./tsconfig.build.json" },
```

- [ ] **Step 3: Verify build still works**

Run: `npm run build`
Expected: Build succeeds, `dist/` contains `office.mjs`, `office.d.mts`, `office.js`

- [ ] **Step 4: Verify tests still pass**

Run: `npm run test:unit`
Expected: All unit tests and type tests pass

- [ ] **Step 5: Commit**

```bash
git add tsconfig.build.json tsdown.config.ts
git commit -m "refactor: extract tsconfig.build.json for tsdown dts generation"
```

---

### Task 2: Convert `tsconfig.json` to project-wide type-check config

**Files:**
- Modify: `tsconfig.json`

- [ ] **Step 1: Replace `tsconfig.json` with type-check-only config**

```jsonc
{
  "compilerOptions": {
    "target": "esnext",
    "module": "nodenext",
    "moduleResolution": "nodenext",
    "lib": ["esnext", "dom"],
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "noEmit": true
  },
  "exclude": ["tests/e2e/fixture/**", ".references/**", "node_modules/**", "dist/**"]
}
```

- [ ] **Step 2: Verify project-wide type checking works**

Run: `npx tsc --noEmit`
Expected: No type errors. This now covers `src/`, `tests/unit/`, `tests/e2e/*.ts`, and `*.config.ts`.

- [ ] **Step 3: Verify build still works**

Run: `npm run build`
Expected: Build succeeds (`tsconfig.build.json` extends the updated base and overrides with build settings)

- [ ] **Step 4: Verify tests still pass**

Run: `npm run test:unit`
Expected: All unit tests and type tests pass

- [ ] **Step 5: Commit**

```bash
git add tsconfig.json
git commit -m "refactor: make tsconfig.json project-wide type-check only"
```

---

### Task 3: Clean up `tsconfig.test-d.json`

**Files:**
- Modify: `tsconfig.test-d.json`

- [ ] **Step 1: Remove redundant overrides**

Replace `tsconfig.test-d.json` with:

```jsonc
{
  "extends": "./tsconfig.json",
  "include": ["src/**/*.ts", "tests/unit/**/*.test-d.ts"],
  "compilerOptions": {
    "types": ["office-js", "custom-functions-runtime"]
  }
}
```

Removed: `rootDir` (parent no longer has it), `noEmit` (parent already sets it).

- [ ] **Step 2: Verify type tests still pass**

Run: `npm run test:unit`
Expected: All unit tests and type tests pass (vitest typecheck uses this config)

- [ ] **Step 3: Commit**

```bash
git add tsconfig.test-d.json
git commit -m "refactor: clean up redundant overrides in tsconfig.test-d.json"
```
