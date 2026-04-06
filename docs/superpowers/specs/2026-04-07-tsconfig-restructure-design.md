# tsconfig restructure: Separate type checking from build

## Background

After migrating the build from tsc to tsdown, tsc is now used purely for type checking. However, `tsconfig.json` still contains build-oriented settings (`outDir`, `declaration`, `rootDir`) and only covers `src/`. This means tests, e2e test files, and config files are not type-checked by tsc.

## Goal

- Make `tsconfig.json` a project-wide type-checking config (with `noEmit: true`)
- Create `tsconfig.build.json` to preserve build settings for tsdown's dts generation
- Type-check all project TypeScript files except `tests/e2e/fixture/` (an independent Office Add-in project)

## Design

### `tsconfig.json` — Type checking only (modified)

Remove build settings, add `noEmit: true`, remove `include` to cover the entire project, use `exclude` to skip irrelevant directories.

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
  "exclude": ["tests/e2e/fixture/**", "node_modules/**", "dist/**"]
}
```

**Changes from current:**
- Removed: `outDir`, `declaration`, `rootDir`
- Added: `noEmit: true`
- Removed: `include` (defaults to all `.ts` files)
- Updated: `exclude` to cover `tests/e2e/fixture/`, `node_modules/`, `dist/`

### `tsconfig.build.json` — tsdown dts generation (new)

Preserves the current `tsconfig.json` settings needed by tsdown for `.d.mts` generation.

```jsonc
{
  "compilerOptions": {
    "target": "esnext",
    "module": "nodenext",
    "moduleResolution": "nodenext",
    "lib": ["esnext", "dom"],
    "declaration": true,
    "outDir": "dist",
    "rootDir": "src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true
  },
  "include": ["src/**/*.ts"]
}
```

### `tsdown.config.ts` — Update tsconfig reference (modified)

```ts
dts: { tsconfig: "./tsconfig.build.json" },
```

### `tsconfig.test-d.json` — No changes

Extends `tsconfig.json` and adds `noEmit: true` (now redundant but harmless). Continues to work as before for vitest typecheck.

### `package.json` scripts — No changes

Existing scripts remain functional. `tsc --noEmit` uses the updated `tsconfig.json` by default.

## Files affected

| File | Action |
|------|--------|
| `tsconfig.json` | Modify — type-check only, project-wide |
| `tsconfig.build.json` | Create — tsdown dts config |
| `tsdown.config.ts` | Modify — point to `tsconfig.build.json` |
| `tsconfig.test-d.json` | No change |
| `package.json` | No change |

## Out of scope

- Type checking `tests/e2e/fixture/` — independent project with its own tsconfig
- Adding a dedicated typecheck script to `package.json` — can be done separately if needed
