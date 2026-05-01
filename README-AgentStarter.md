# agent-starter

A drop-in kit that keeps AI-generated code small, intent-clear, and reviewable
without you having to read every line. Designed for solo developers who use
Cursor, Claude Code, or Codex CLI and want one workflow they can apply to
every project.

## What's in the box

| File | Read by | Purpose |
|---|---|---|
| `AGENTS.md` | Codex CLI, Cursor, Copilot, Windsurf, Amp, Devin (native) | Project constitution. Hard rules every agent obeys. |
| `CLAUDE.md` | Claude Code | Points Claude at `AGENTS.md` and lists permission boundaries. |
| `.cursor/rules/00-core.mdc` | Cursor | Always-on pointer to `AGENTS.md` + Cursor-specific behaviors. |
| `.cursor/rules/10-commits.mdc` | Cursor | Commit message format with examples. |
| `.cursor/rules/20-scope-contract.mdc` | Cursor | Forces a 5-line scope contract before any non-trivial task. |
| `.gitmessage` | git (and the agent) | Commit template: `type(scope): why` + structured body. |
| `.coderabbit.yaml` | CodeRabbit | Auto-review config tuned for solo devs. |
| `.github/pull_request_template.md` | GitHub | PR template enforcing Goal / Out-of-scope / Done-when / Rollback. |
| `setup.ps1` / `setup.sh` | You | One-shot wiring of `git config commit.template` + verification. |

## The four-layer defense

1. **Agent rules** keep agents on a leash (constitution + Cursor rules + Claude bridge).
2. **Scope contract** forces a tiny written commitment before any code is generated.
3. **Atomic commits** with `type(scope): why` make `git log` a readable changelog.
4. **CodeRabbit** auto-reviews every PR so you can skim a summary instead of reading every line.

You stay the human gatekeeper at one and only one step: **the merge button**.

## Use it on a new project

```powershell
# Option A: GitHub template (recommended)
gh repo create my-thing --template <your-username>/agent-starter --clone --private
cd my-thing
.\setup.ps1                # or: bash setup.sh

# Option B: copy-in
git init my-thing
cd my-thing
# copy contents of agent-starter/ into here
.\setup.ps1
```

Then:

1. Open `AGENTS.md` and fill in the `Project context` section
   (what this project is, language/stack, how to run/test).
2. (Optional) Tweak `.coderabbit.yaml` if your project has unusual generated
   files to ignore.
3. Make your first feature branch and start working. The agent will follow
   the rules.

## Backport into an existing project

Drop the files in incrementally, in this order, so you can stop at any layer
that already feels like enough:

1. `AGENTS.md` — biggest signal, smallest cost. Start here.
2. `.cursor/rules/` — only if you use Cursor.
3. `CLAUDE.md` — only if you use Claude Code.
4. `.gitmessage` + run `git config commit.template .gitmessage`.
5. `.github/pull_request_template.md` — for GitHub-hosted projects.
6. `.coderabbit.yaml` + install CodeRabbit on the repo.

You can copy them all at once and run `setup.ps1` / `setup.sh` to wire up the
git template, but adopt the workflow one layer at a time.

## The per-task workflow

```
1. Describe task to agent.
2. Agent returns 5-line scope contract + numbered plan. STOP.
3. You eyeball the contract. Approve, or ask to narrow.
4. Agent implements on a feature branch (<= ~300 lines).
5. Agent commits with structured "why"-focused message.
6. Agent pushes branch and opens a PR using the template.
7. CodeRabbit reviews automatically.
8. You skim the PR description + commit log + CodeRabbit summary.
9. You merge. (Agent never merges.)
```

Total of *your* attention per change: about 2 minutes.

## Customising

- **Add a project rule:** prefer `AGENTS.md` for cross-tool rules; use
  `.cursor/rules/*.mdc` for Cursor-only or for keeping Cursor's always-on
  context lean. Keep `AGENTS.md` under ~200 lines.
- **Stricter mode for one repo:** add a "Checkpoint review" line to that
  repo's `AGENTS.md`: *"After implementation, show me the diff and commit
  message; do not run `git commit` until I say so."*
- **Less strict for throwaway repos:** delete `20-scope-contract.mdc` and
  loosen the diff budget in `AGENTS.md`. The rest still helps.

## Iterating

Don't pre-write rules for problems that haven't happened. Every time the AI
does something annoying or overkill, add one line to `AGENTS.md` (or a Cursor
rule) so it doesn't happen again. The kit is a starting point, not a
finished spec.

## What this kit deliberately does NOT include

- A specific test runner, linter, or formatter — those are project-specific.
- Pre-commit hooks — they're great, but they're language-specific and add
  setup friction. Add them per-project once you know the stack.
- CI pipelines — same reason. The PR template + CodeRabbit cover the
  highest-leverage 80%.

## Credits / further reading

- [Cursor Rules docs](https://cursor.sh/docs/rules)
- [AGENTS.md Guide 2026 — vibecoding.app](https://vibecoding.app/blog/agents-md-guide)
- [Cursor agent best practices](https://www.cursor.com/blog/agent-best-practices)
- [How Senior Engineers Actually Review PRs — Manav Gandhi](https://medium.com/@27manavgandhi/how-senior-engineers-actually-review-pull-requests-cecb6041e661)
- [Stop Vibe Merging — dev.to](https://dev.to/shmulc/stop-vibe-merging-1jpo)
- [CodeRabbit docs](https://docs.coderabbit.ai/)
