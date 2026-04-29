#!/usr/bin/env bash
# setup.sh - one-shot project setup for the agent-starter kit (macOS / Linux / Git Bash)
#
# Usage (from the project root, after cloning the template):
#   bash setup.sh
#
# What this does:
#   1. Verifies you're inside a git repo (initialises one if not).
#   2. Wires up the .gitmessage commit template for this repo.
#   3. Prints the next-step checklist.

set -euo pipefail

c_cyan='\033[0;36m'
c_green='\033[0;32m'
c_yellow='\033[0;33m'
c_reset='\033[0m'

step() { printf "${c_cyan}==> %s${c_reset}\n" "$1"; }
ok()   { printf "    ${c_green}%s${c_reset}\n" "$1"; }
warn() { printf "    ${c_yellow}%s${c_reset}\n" "$1"; }

step "Checking git repository"
if git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
    ok "git repo detected."
else
    warn "Not a git repo yet. Initialising..."
    git init -q
    git branch -M main 2>/dev/null || true
    ok "git initialised on branch 'main'."
fi

step "Wiring up .gitmessage commit template"
if [ -f ".gitmessage" ]; then
    git config commit.template .gitmessage
    ok "git config commit.template = .gitmessage"
else
    warn ".gitmessage not found in this directory. Skipping."
fi

step "Verifying agent-starter files are in place"
expected=(
    "AGENTS.md"
    "CLAUDE.md"
    ".cursor/rules/00-core.mdc"
    ".cursor/rules/10-commits.mdc"
    ".cursor/rules/20-scope-contract.mdc"
    ".gitmessage"
    ".coderabbit.yaml"
    ".github/pull_request_template.md"
)
missing=()
for f in "${expected[@]}"; do
    [ -e "$f" ] || missing+=("$f")
done
if [ "${#missing[@]}" -eq 0 ]; then
    ok "All starter-kit files present."
else
    warn "Missing files (kit may be incomplete):"
    for m in "${missing[@]}"; do warn "  - $m"; done
fi

cat <<'EOF'

==> Next steps

  1. Edit AGENTS.md and fill in the 'Project context' section
     (what this project is, language/stack, run/test commands).

  2. (Optional) Tweak .coderabbit.yaml path filters if your project has
     unusual generated files.

  3. Push to GitHub:
        git add -A
        git commit          # the template will prompt you for the structured message
        gh repo create <name> --private --source=. --push

  4. Install CodeRabbit on the repo:
        https://github.com/marketplace/coderabbitai

  5. From now on, every change goes:
        feature branch -> commit (atomic, why-focused) -> push -> PR -> CodeRabbit -> you merge.

EOF
