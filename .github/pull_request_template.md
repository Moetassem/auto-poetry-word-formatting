<!--
  Fill out every section. The whole point of this template is to make this PR
  reviewable in 60 seconds. If a section does not apply, write "n/a" — do not
  delete it.
-->

## Goal

<!-- One sentence. What outcome does this PR deliver? -->

## Out of scope

<!--
  What this PR is deliberately NOT doing, even though it might be tempting.
  This is the most important section for AI-generated PRs — it pins down what
  was kept out so reviewers know what not to look for.
-->

-

## Files touched

<!-- Brief list with one-line per file describing the change. -->

-

## Done when

<!--
  The observable check that proves this PR works.
  e.g. "test_parser.py passes", "running `app --help` shows the new flag",
       "manually opening the doc in Word renders correctly".
-->

-

## How to roll back

<!--
  Single command or short steps to undo this if it breaks something in prod.
  e.g. "git revert <this-commit>", "feature flag FOO_ENABLED off",
       "redeploy previous tag v1.2.3".
-->

-

## Trade-offs / alternatives considered

<!--
  Why this approach? What did you consider and reject? What did you defer?
  This goes here, NOT in code comments.
-->

-

## Follow-ups deferred

<!-- Things you noticed but intentionally did not do in this PR. -->

-

---

### Self-checklist

- [ ] Diff is <= ~300 lines (or this PR explicitly explains why it must be larger).
- [ ] One logical change. No mixed refactor + feature.
- [ ] No placeholder code (`TODO: implement`, `# rest of code...`, empty `pass`).
- [ ] No unrelated files touched.
- [ ] No new dependencies introduced silently — they are listed above if any.
- [ ] No secrets, tokens, or `.env*` content committed.
- [ ] Tests / typecheck / linter pass locally.
- [ ] Commit messages follow the `.gitmessage` template (one logical change each).
