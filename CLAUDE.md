# CLAUDE.md

Claude Code does not yet read `AGENTS.md` natively. This file bridges the gap.

## Read this first

**Read and follow `@AGENTS.md` at the project root before any task.** It is
the project constitution. If anything below conflicts with `AGENTS.md`, this
file's overrides win — but only for the items explicitly listed here.

## Claude-specific notes

- Use Plan mode for any non-trivial task. Do not start editing until the user
  approves the plan.
- When in doubt about which file to modify, ask. Do not guess across multiple
  candidates.
- The `.gitmessage` template is set as the project commit template. When you
  run `git commit`, follow that structure (see also
  `.cursor/rules/10-commits.mdc`).

## Permission boundaries

These are the same as `AGENTS.md` "Git boundaries" but worth restating because
Claude Code's tool permissions can be configured per-project:

- Allowed without asking: read files, search, run tests, run typechecker,
  run linter, `git status`, `git diff`, `git log`, `git add`, `git commit`,
  `git push origin <feature-branch>`, `gh pr create`, `gh pr view`.
- Ask first: install dependencies, modify CI config, modify
  `package.json` / `pyproject.toml` / `requirements.txt`, modify `.env*`,
  delete files.
- Never without explicit "yes, force-push" / "yes, merge" from the user:
  `git push --force`, `git reset --hard`, `git clean -fd`, `gh pr merge`,
  pushing to `main` / `master`, deleting branches that are not your own
  feature branch.

## Tips for working well with this project

- Prefer the canonical pattern. Find an existing example before inventing a
  new one.
- Surgical edits over rewrites for any file >500 lines.
- If a change crosses ~300 lines, stop and propose splitting it.
