#!/usr/bin/env bash
# Policy Manager — production safety hook
# Blocks commands that could affect production without explicit owner approval.
# Invoked by Claude Code as a PreToolUse hook for Bash commands.

set -u

# Read the tool input JSON from stdin
input=$(cat)

# Extract the command string (Bash tool uses `.tool_input.command`)
command=$(printf '%s' "$input" | python -c "
import json, sys
try:
    data = json.load(sys.stdin)
    cmd = data.get('tool_input', {}).get('command', '')
    print(cmd)
except Exception:
    print('')
" 2>/dev/null)

if [ -z "$command" ]; then
    exit 0
fi

# Normalize for pattern matching (lowercase, collapse whitespace)
lc=$(printf '%s' "$command" | tr '[:upper:]' '[:lower:]' | tr -s '[:space:]' ' ')

# Strip anything after the first heredoc start ("<<'EOF'", "<<EOF", "<<-EOF") or
# the opening quote of a long -m argument. This prevents false positives where
# commit message bodies mention flag names like --no-verify or commands like
# "git push" that are being described rather than executed.
# We keep only the portion of the line before the heredoc/body begins.
head=$(printf '%s' "$lc" | sed -E "s/<<-?'?[a-z0-9_]+'?.*$//" | sed -E 's/-m "[^"]*".*$/-m/' | sed -E "s/-m '[^']*'.*$/-m/")

block() {
    # Write reason to stderr so Claude Code surfaces it to the agent
    printf 'BLOCKED by .claude/hooks/guard-production.sh: %s\n' "$1" >&2
    printf 'If this is intentional and authorised by the repo owner (Gary Finberg),\n' >&2
    printf 'ask the owner to run the command manually or adjust the hook.\n' >&2
    exit 2
}

# 1. Block packaging for production ship
case "$lc" in
    *"gulp package-solution --ship"*) block "gulp package-solution --ship (production packaging requires owner approval per CLAUDE.md)";;
    *"npm run ship"*)                 block "npm run ship (production packaging requires owner approval per CLAUDE.md)";;
esac

# 2. Block PnP PowerShell provisioning scripts (they write to live SharePoint)
case "$lc" in
    *"invoke-pnpsitetemplate"*)       block "PnP site template invocation (writes to live SharePoint)";;
    *"deploy-allpolicylists.ps1"*)    block "Deploy-AllPolicyLists.ps1 (provisions SP lists on live site)";;
    *"seed-approvalandnotification"*) block "Seed-ApprovalAndNotificationData (writes sample data to live SP)";;
    *"deploy-sampledata.ps1"*)        block "Deploy-SampleData.ps1 (writes sample data to live SP)";;
    *"create-policymanagementlists"*) block "Create-PolicyManagementLists.ps1 (provisions SP lists on live site)";;
    *"provision-sharepointpages"*)    block "Provision-SharePointPages.ps1 (creates SP pages on live site)";;
esac

# 3. Block scripts directory PowerShell execution pattern-wide
case "$lc" in
    *"scripts/policy-management/"*".ps1"*) block "scripts/policy-management/ PowerShell execution (live SharePoint writes)";;
    *"scripts\\\\policy-management\\\\"*".ps1"*) block "scripts\\policy-management\\ PowerShell execution (live SharePoint writes)";;
esac

# 4. Block Azure infrastructure deployments
case "$lc" in
    *"az deployment group create"*)   block "az deployment group create (production Azure deployment)";;
    *"az deployment sub create"*)     block "az deployment sub create (subscription-level Azure deployment)";;
    *"az functionapp deployment"*)    block "az functionapp deployment (production Function App deployment)";;
    *"az functionapp create"*)        block "az functionapp create (production Function App provisioning)";;
    *"az logic"*)                     block "az logic (Logic App modifications on production)";;
    *"az keyvault"*"set"*)            block "az keyvault set (modifying production Key Vault)";;
    *"deploy.ps1"*)                   block "deploy.ps1 (Azure Function / Logic App deployment script)";;
esac

# 5. Block dangerous git operations (match against $head so commit message bodies
#    that mention these patterns don't trigger false positives)
case "$head" in
    *"git push --force"*)       block "git push --force (history rewrite disabled per repo policy)";;
    *"git push -f "*)           block "git push -f (history rewrite disabled per repo policy)";;
    *"git push --no-verify"*)   block "git push --no-verify (hook bypass disabled)";;
    *"git commit"*"--no-verify"*) block "git commit --no-verify (hook bypass disabled)";;
    *"git reset --hard origin"*) block "git reset --hard origin (destructive — ask owner)";;
    *"git config --global"*)    block "git config --global (global git config changes disabled)";;
    # Direct push to master/main is enforced server-side by ADO branch policy (PR-required).
    # The hook no longer duplicates that check — ADO is authoritative for everyone,
    # including QA contributors who cannot bypass server-side policy.
esac

# 6. Block destructive filesystem operations at repo root scope
case "$lc" in
    *"rm -rf /"*)          block "rm -rf / (destructive)";;
    *"rm -rf ~"*)          block "rm -rf ~ (destructive)";;
    *"rm -rf *"*)          block "rm -rf * (destructive)";;
    *"rmdir /s /q c:"*)    block "rmdir /s /q c: (destructive)";;
esac

exit 0
