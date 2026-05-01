# AGENTS.md

This file is the project constitution for AI coding agents (Codex CLI, Cursor,
Claude Code, Copilot, and friends). Read it before doing anything. If a rule
here conflicts with another instruction, this file wins unless the human user
explicitly overrides it in the current session.

> Tool support: Codex CLI, Cursor, Copilot, Windsurf, Amp, and Devin read this
> file natively. Claude Code does not yet — `CLAUDE.md` references this file so
> Claude reads it manually.

---

## Project context

<!-- Fill these in per project. Keep terse. -->
- **What this project is:** Automatic Aesthatic fomatting for Arabic Poetry in MS Word 
- **Primary language / stack:** VBA scripts/macros
- **How to run / test:** manual testing
- **Canonical examples:** auto-poetry-word-formatting/macros/FormatPoem.bas

---

## Non-negotiable rules

These apply to every task in every session.

### 1. Plan before you implement

For anything beyond a true one-liner (typo, rename, single-line bug fix):

1. Output a short numbered plan (3–8 steps).
2. **Stop. Wait for explicit approval** ("go", "yes", "proceed").
3. Only then write code.

If the user's request is ambiguous, ask one or two pointed clarifying
questions before planning. Do not guess on questions that change the design.

For bug fixes/new features, always write a failing test first. The fix is only complete
when that test passes with the fewest lines of code possible.

### 2. The scope contract

Every non-trivial task must begin with this exact 5-line block, then wait for
approval:

```
Goal:           <one sentence>
Files touched:  <explicit list, paths>
Out of scope:   <what we are NOT doing>
Tests added:    <Tests written for this feature/bugfix>
Done when:      <observable check>
Rollback:       <how to revert>
Reverse Context:<code that became redundant/deleted>
```

If during implementation you discover the scope was wrong, **stop and renegotiate
the contract**. Do not silently expand scope.

### 3. Diff budget: ~300 lines

Use the smallest possible edit. Do not rewrite an entire file to change one
function. If a change would exceed roughly 300 lines of new/modified code
(excluding generated files, lockfiles, snapshots), stop and propose a split
into smaller commits or PRs. The user must approve any larger change
explicitly.

Generated files, vendored dependencies, and lockfiles do not count toward this
budget but must be committed separately.

### 4. Scope discipline

- Touch only the files needed for the requested change.
- Never refactor unrelated code.
- Never add "nice-to-haves" the user did not ask for.
- Never add "just in case" logic, extra abstractions, or "defensive" code
  unless explicitly requested.
- Never add a new library or dependency, expand abstractions, or add config
  layers without asking for permission first.
- If you notice something worth changing, list it as a follow-up suggestion in
  the PR description — do not do it.

### 5. Interface-first for new features

When adding a new feature:

1. Define types, function signatures, or API shapes first.
2. Show them to the user.
3. Only after approval, write the implementation.

This prevents wasted compute on the wrong abstraction.

### 6. No placeholders, no half-features

- Never emit `// TODO: implement this`, `# rest of code...`, `pass  # later`,
  or equivalent.
- Every function you write is fully implemented or it is not added at all.
- If you cannot complete it, say so and ask for guidance.

### 7. Verify before you import

- Before importing a package, confirm it is already installed (check
  `package.json` / `requirements.txt` / `pyproject.toml` / etc., or run
  `npm list <pkg>` / `pip show <pkg>`).
- Never assume an API exists based on training data. If unsure about a
  framework's current API, say so and ask.

### 8. One logical change per commit

- One concern per commit. Refactor + feature = two commits.
- Conventional Commits format with a "why" focus:
  `type(scope): why-it-matters`
- See `.gitmessage` and `.cursor/rules/10-commits.mdc` for the full format and
  examples.

### 9. Trade-offs go in the PR description

Explain *why* you chose this approach, what alternatives you considered, and
what you deferred — in the PR description, not in code comments. Code comments
should explain non-obvious intent or constraints, never narrate what the code
does.

### 10. Follow the existing patterns

Before introducing a new pattern, library, or file structure, find the
canonical example in this codebase and mimic it. If no precedent exists, ask
which direction to set. Always search the codebase for an existing utility
before writing a new helper.

---

## Git boundaries (hard limits)

- **Never merge to `main` / `master`.** Merging is the human's job.
- **Never force-push** any branch unless the user explicitly says "force push".
- **Never run destructive git commands** (`reset --hard`, `clean -fd`,
  `branch -D`, `push --force`) without explicit confirmation.
- **Never delete config files** (`.env`, `.env.*`, `package.json`,
  `pyproject.toml`, `requirements.txt`, `Cargo.toml`, etc.) without explicit
  confirmation.
- **Never commit secrets.** If you see a credential, stop and report it.
- Commits and pushes to feature branches are fine. Opening a PR is fine.
  Merging the PR is not.

---

## Standard task workflow

```
1. Read the user's request.
2. Read relevant files. Find the canonical pattern.
3. Output the 5-line scope contract + numbered plan.
4. Wait for approval.
5. Create or switch to a feature branch (e.g. feat/<short-slug>).
6. Implement, respecting the diff budget.
7. Self-check: run tests / typecheck / linter the project provides.
8. Commit using the .gitmessage template (one logical change per commit).
9. Push the branch.
10. Open a PR using the PR template.
11. STOP. The user reviews and merges.
```

> Steps 7–11 are mandatory once `git remote -v` shows a remote. Do not stop
> after step 6 (implement) unless the user explicitly says "don't commit" or
> the project has no remote configured yet.

---

## When the user says "fix it" or gives a tiny task

You can skip the formal scope contract for clearly trivial work
(typo fixes, single-variable renames, obvious one-line bugs). Use judgment:
if you have to think about it for more than a moment, it is not trivial.

---

## Self-improvement

When the user pushes back on something you did, propose a one-line addition to
this file or to `.cursor/rules/` so the same mistake does not happen again.
Do not pre-emptively invent rules for problems that have not occurred.

Keep this file under ~200 lines. If a section grows large, move it to
`.cursor/rules/*.mdc` with appropriate scoping.
