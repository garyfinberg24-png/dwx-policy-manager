# Policy Manager — QA Workflow Guide

This document is the **QA-specific overlay** to the main `CLAUDE.md`. Read `CLAUDE.md` first for full project context (architecture, SharePoint lists, design system, etc.), then use this file for QA-specific workflow rules.

## Who this is for

The QA lead testing Policy Manager and fixing defects using Claude Code. If you're the owner (Gary), the rules in `CLAUDE.md` are the source of truth — this file is stricter in places.

## Repository access

- Repo: `https://dev.azure.com/gfinberg/DWx/_git/dwx-policy-manager` (Azure DevOps — GitHub is no longer used)
- Main branch: `master`
- You have Contribute + Create branch + PR permissions on this repo only
- You cannot push directly to `master` — ADO branch policy requires a PR with owner approval

## Testing reference

Your primary QA reference is the interactive test guide:

- **Location:** `docs/testing-guide-qa-lead.html`
- **Contents:** 79 interactive test cases across 13 sections (Functional F1–F6, Process P1–P3, Role-Based R1–R4)
- **Usage:** Open the file in a browser, work through each test, click Pass/Fail/Block/Skip, take notes, export CSV when done
- The test guide includes a role visibility matrix and a policy lifecycle reference

The SharePoint test site is `https://mf7m.sharepoint.com/sites/PolicyManager`.

## Branch naming convention

All QA work goes on a branch. Never modify `master` directly.

Format: `qa/<short-description>` or `qa/<test-id>-<short-description>`

Examples:
- `qa/r1-user-role-nav-missing-items`
- `qa/f3-policy-hub-filter-bug`
- `qa/p2-publish-pipeline-missing-email`

Keep branch names short, descriptive, and lowercase-with-hyphens.

## Pull request rules

1. **One issue per PR.** If you find three bugs while testing, open three PRs on three branches. Do not bundle unrelated fixes.
2. **Small, focused changes.** If a fix touches more than ~5 files or more than ~200 lines, stop and discuss with the owner first.
3. **Always tag the owner as reviewer.** Gary Finberg is the required reviewer on `master`.
4. **Reference the test case.** In the PR description, include the test ID from `docs/testing-guide-qa-lead.html` (e.g., "Fixes F3.2: Policy Hub category filter returns empty results").
5. **Describe what you tested.** Include a "Test plan" section in the PR body listing what you verified before pushing.
6. **Never skip hooks.** If a pre-commit or pre-push hook fails, fix the underlying issue — don't use `--no-verify`. The `.claude/settings.json` blocks this anyway.

## What you can do with Claude Code

- Read any file in the repo
- Run `npm install`, `npm run build`, `gulp bundle` (non-ship build)
- Run unit tests (`npx jest`), Playwright E2E tests (`npx playwright test`)
- Create branches, commit, push to your QA branches
- Open PRs against `master` via `az repos pr create` or the ADO web UI
- Edit source files to fix defects
- Run the TinyMCE/Playwright mock servers for local testing

## What you CANNOT do

These are blocked by `.claude/hooks/guard-production.sh` — the commands will be rejected even if Claude tries to run them. This is intentional.

1. **Production packaging.** `gulp package-solution --ship` and `npm run ship` are blocked. Only the owner packages for production deployment.
2. **SharePoint provisioning scripts.** Anything under `scripts/policy-management/*.ps1` is blocked (Deploy-AllPolicyLists, Create-PolicyManagementLists, Provision-SharePointPages, etc.). These write to the live SharePoint site.
3. **Azure Function / Logic App deployments.** `az deployment group create`, `az functionapp deployment`, `deploy.ps1` are blocked. Only the owner deploys infrastructure changes.
4. **Force push, reset --hard origin, --no-verify.** All blocked. If git state is confusing, ask the owner before attempting recovery.
5. **Direct push to master/main.** Blocked by both the hook and ADO branch policy. Use a `qa/*` branch and open a PR.
6. **Global git config changes.** Blocked. Set config locally (`git config user.email ...`) if needed.
7. **Key Vault writes.** `az keyvault set` is blocked to prevent accidental secret overwrites.

If you hit a hook block and genuinely need the command, ask the owner. Do not try to work around the hook.

## Mandatory task execution rules (from CLAUDE.md)

These apply to you as well. Copy directly:

1. **Always read `CLAUDE.md` before you do anything.**
2. **Always ask questions if you are unsure of the task or requirement.**
3. **Be systematic in your planning and execution.**
4. **After you complete a task, always validate the result.**
5. **One task at a time, verify before moving on.** Don't batch multiple items into a single large edit.
6. **A successful build does NOT mean the task is done.** After every edit, re-read the changed section and verify it matches the requirement point by point.
7. **Track every sub-item in the todo list.** If the user asks for 3 fixes, create 3 separate todo items.
8. **Explain back before implementing.** Before writing code, describe what you'll do in your own words so the owner can confirm you understood. Wait for explicit approval ("go for it", "approved") before editing.

## Typical QA session workflow

1. Pick a test case from `docs/testing-guide-qa-lead.html`
2. Execute the test on `https://mf7m.sharepoint.com/sites/PolicyManager`
3. If it fails, gather details: screenshot, console output (F12), network tab (F12), steps to reproduce
4. Ask Claude Code to investigate: "Test case F3.2 is failing — filter returns empty. Here's the console error: [paste]. Please investigate the Policy Hub filter code."
5. Let Claude propose a fix. Review it. Ask questions.
6. Approve the fix. Claude creates a `qa/*` branch, commits, pushes.
7. Verify the fix locally with `gulp bundle` (non-ship) + browser test if possible
8. Open a PR via `az repos pr create --title "..." --description "..." --target-branch master` (or via the ADO web UI)
9. Tag Gary as reviewer. Add a clear test plan in the description.
10. Move to the next test case

## Common gotchas

- **PowerShell scripts starting with numbers** need `.\` prefix: `.\08-Approval-Lists.ps1`
- **SPFx CDN caching** — after a build, you may need a hard refresh (Ctrl+F5) in SharePoint to see changes. Some changes also require an app catalog re-upload (owner task).
- **`@ts-nocheck` is in many files** — don't remove it unless you're specifically hardening types in that file and fixing all resulting errors.
- **Class components, not hooks** — this codebase uses React 17 class components throughout. Use `setState`, not `useState`, when editing.
- **SharePoint list names** — always use `PM_LISTS.POLICIES` etc. from `src/constants/SharePointListNames.ts`, never hardcode list names.
- **Never commit `.env` files, secrets, or `sharepoint/solution/*.sppkg`** — these are gitignored but double-check before staging.

## Design system rules (mandatory, from CLAUDE.md)

Do not violate these in any fix:

- **Icons:** ONLY SVG line icons (`<svg viewBox="0 0 24 24" fill="none">` with stroke paths). NO emoji. NO Fluent `<Icon>` for decorative use.
- **Colour palette:** Forest Teal `#0d9488` primary, `#0f766e` dark teal. NO Microsoft Blue (`#0078d4`).
- **Border radius:** 4px on form controls (dropdowns, inputs, buttons), 10px on KPI cards.
- **Cards:** `background: #fff`, `border: 1px solid #e2e8f0`, `borderRadius: 10px`.
- **Panels:** Use `StyledPanel` (slide-in from right), not persistent side panels. Do NOT modify any Fluent UI `<Panel>` component directly — log issues for owner review.

## Escalation

If Claude Code reports something weird, Azure deployments fail, SharePoint site is down, or you see data loss — stop immediately and message Gary. Do not attempt recovery with destructive commands.

## Contact

Owner: Gary Finberg (gary@firsttech.digital)
