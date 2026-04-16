# CLAUDE.md

## Development Principles

- **Excel / office.js behavioral consistency**: This library mocks the Office JavaScript API. All behavior — especially formula parsing, type coercion, and custom function invocation — must match real Excel / office.js semantics. When in doubt, test against Excel and follow its behavior. For example, quoted formula arguments (e.g., `"2023"`) must always be passed as strings, never coerced to numbers.
