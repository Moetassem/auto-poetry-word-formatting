# setup.ps1 - one-shot project setup for the agent-starter kit (Windows / PowerShell)
#
# Usage (from the project root, after cloning the template):
#   .\setup.ps1
#
# What this does:
#   1. Verifies you're inside a git repo (initialises one if not).
#   2. Wires up the .gitmessage commit template for this repo.
#   3. Prints the next-step checklist.
#
# It does NOT install dependencies, create branches, or push anything.

$ErrorActionPreference = "Stop"

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "    $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "    $msg" -ForegroundColor Yellow }

Write-Step "Checking git repository"
$inRepo = $false
try {
    git rev-parse --is-inside-work-tree 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { $inRepo = $true }
} catch { }

if (-not $inRepo) {
    Write-Warn "Not a git repo yet. Initialising..."
    git init | Out-Null
    git branch -M main 2>$null
    Write-Ok "git initialised on branch 'main'."
} else {
    Write-Ok "git repo detected."
}

Write-Step "Wiring up .gitmessage commit template"
if (-not (Test-Path ".gitmessage")) {
    Write-Warn ".gitmessage not found in this directory. Skipping."
} else {
    git config commit.template .gitmessage
    Write-Ok "git config commit.template = .gitmessage"
}

Write-Step "Verifying agent-starter files are in place"
$expected = @(
    "AGENTS.md",
    "CLAUDE.md",
    ".cursor\rules\00-core.mdc",
    ".cursor\rules\10-commits.mdc",
    ".cursor\rules\20-scope-contract.mdc",
    ".gitmessage",
    ".coderabbit.yaml",
    ".github\pull_request_template.md"
)
$missing = @()
foreach ($f in $expected) {
    if (-not (Test-Path $f)) { $missing += $f }
}
if ($missing.Count -eq 0) {
    Write-Ok "All starter-kit files present."
} else {
    Write-Warn "Missing files (kit may be incomplete):"
    foreach ($m in $missing) { Write-Warn "  - $m" }
}

Write-Host ""
Write-Step "Next steps"
Write-Host @"

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

"@ -ForegroundColor White
