# AGENTS.md — Repository Instructions

Repository-specific instructions for Codex and AI agents. 

Primary local environment: **Windows + PowerShell**.

---

## 1. Core Philosophy & Priorities

1. **Always Greenfield:** Always treat the project as a new application. Do not assume integration with legacy systems, previous production constraints, or backward compatibility unless strictly instructed.
2. **Zero Technical Debt:** Proactively remove legacy code, deprecated patterns, dead files, and transitional structures. There is zero tolerance for technical debt.
3. **Architectural Authority:** You are explicitly authorized to make destructive changes, rewrite logic, and delete files to improve, modernize, or clean the architecture. 
4. **Cleanliness:** Keep the repository clean, navigable, and strictly organized.

---

## 2. Windows-First Environment

Assume local development is Windows-based.
- Use **PowerShell-compatible commands** unless repository evidence dictates otherwise (e.g., CI containers or specific deployment targets).
- Avoid Unix/Bash-only assumptions (`/tmp`, `chmod`, `rm -rf`, `sed -i`).
- Use repository-relative paths, quote paths that may contain spaces, and rely on cross-platform path APIs in code.

---

## 3. Workflow & Implementation

1. **Assess & Execute:** Inspect only the files necessary for the requested scope. Check the working tree (`git status --short`) before starting.
2. **Refactor Boldly:** Implement the highest-quality architectural solution. Do not hesitate to split, rename, or delete files if it resolves mixed responsibilities.
3. **Verify:** Add or update tests when behavior changes. If no test framework exists, provide a manual verification path.
4. **Document:** Update documentation (setup, commands, public APIs) whenever architectural or structural changes occur.

---

## 4. Structure & Files

- Keep files strictly focused on a single responsibility.
- **Naming:** Prefer explicit, descriptive names. Ban vague directories (`misc`, `old`, `temp`) and catch-all files (`utils.*`, `common.*`) unless strictly required by a framework.
- **Entrypoints:** `index.*` files must only contain exports, framework entrypoints, or lightweight composition logic—no heavy implementation.

---

## 5. Scripts

- Use repository-native scripts first. 
- When writing local scripts, prioritize PowerShell wrappers placed in `scripts/`.
- Scripts must run from the repository root, fail with non-zero exit codes on error, and be documented in `scripts/README.md`.

---

## 6. Security & Dependencies

- **No Secrets:** Never create, print, commit, or expose secrets, tokens, or personal data. Use `.env.example` for required variables.
- **Dependencies:** Prefer standard libraries. Before adding a new dependency, verify if an existing one suffices. Respect the current package manager and lockfile.

---

## 7. Cleanliness Check

Before finalizing any implementation, run `git status --short` and `git ls-files` (or equivalent) to verify:
- No temporary, debug, or build files were left behind.
- Stale references from renamed/deleted files are cleared.
- The `.gitignore` adequately covers any new generated outputs.
- Formatters, linters, and type-checkers pass. (State clearly if a check was bypassed and why).

---

## 8. PROJECT_STATUS.json (If Applicable)

If `PROJECT_STATUS.json` is present or requested, it must act as a strict, active todo list.

```json
{
  "todos": [
    "Short actionable task"
  ]
}
```
- **Rules:** Add active tasks only. Immediately delete completed, obsolete, or duplicated items. Do not store status notes, blockers, chat history, or changelogs here.

---

## 9. Final Response Format

For implementation tasks, your final output must concisely report:
- **What changed:** (Highlighting architectural improvements and legacy code removed).
- **Files modified/deleted.**
- **Checks run & results.**
- **Cleanliness status.**
- **Updated `PROJECT_STATUS.json`** (if applicable).
- **Next steps or remaining architectural risks.**
