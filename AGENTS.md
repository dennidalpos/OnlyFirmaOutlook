# AGENTS.md

`v1.2 · 2026-08-30` — Non-derivable repository facts only. Cap ~2500 characters.

## 1. Identity & Scope
- **Purpose**: [Brief 1-line description of project goal and domain]
- **Runtime / Toolchain**: [e.g. Python 3.12 (uv) | Node 22 (pnpm) | Rust 1.80 (cargo) | Go 1.23 | Terraform / Docker]
- **Out of Scope**: [What this repo explicitly does NOT do]
- **Hard Constraints**: [Non-negotiable domain, performance, or technical invariants]

## 2. Verified Commands
Every command must be executed and verified before adding. Update date on verification.

| Workflow | Command | Shell / Cwd | Verified on | Notes / Examples |
| :--- | :--- | :--- | :--- | :--- |
| **Fast Verification** | | | | Unit tests, quick check, or dry-run |
| **Full Verification** | | | | Full test suite, e2e, plan, or release build |
| **Single Target** | | | | e.g. `<runner> <path> -- <filter>` |
| **Format / Lint** | | | | Linter, formatter, or static analysis |
| **Type / Schema Check**| | | | Typechecker, schema validator, or compiler check |
| **Build / Run / Plan** | | | | Build artifact, run entry point, or infra plan |

## 3. Architecture & Boundaries
- **Structure**: [Key entry points, data flows, and module boundaries not obvious from directory tree]
- **Generated / Vendored**: [Paths to regenerate or external copies — never hand-edit]
- **Protected Paths**: [Legacy, frozen, or sensitive paths to avoid modifying]
- **Conventions**: [Repo-specific patterns: error handling, logging, naming, state/DI]

## 4. Sensitive Areas & Gotchas
- **Sensitive Areas**: [Destructive operations, external paid APIs, production resources, live state]
- **Required Env / Config**: [Variable or secret names only (e.g. API_KEY, DB_URI) — no values]
- **Gotchas & Quirks**: [Non-obvious pitfalls, platform-specific quirks, edge cases]
