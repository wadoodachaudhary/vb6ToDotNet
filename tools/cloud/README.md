# HomeFront in Claude Code cloud sessions

A cloud session starts from this repo (`vb6ToDotNet`, public). Everything else comes from
private GitHub copies:

| What | Where | Cloud path |
|---|---|---|
| Outer repo (VB6 source, tools, skills) | `wadoodachaudhary/vb6ToDotNet` main (public) | the session checkout = `VBToCSharp/HomeFront` |
| HomeFront app | `wadoodachaudhary/HomeFront` main (**private** mirror of the Mac's nested repo) | `<checkout>/MobileSource/HomeFront` |
| FlexKit | `wadoodachaudhary/FlexKit` branch `telerik-parity-20260904` (**private**) | `<checkout>/../FlexKit` |
| Database | release `db-20260930` asset `HOMEFRONTSQL-20260930.bak` on the private HomeFront repo (full copy, client data) | SQL Server 2022 in Docker, `localhost,1433`, database `HOMEFRONTSQL` |

GitLab (hyphen-pb, homefront, flexkit) stays the deploy target. Deploys still run only from
the Mac with the `/update-repositories` skill.

## One-time setup (owner, in claude.ai/code → environment selector)

1. Grant Claude GitHub access to `wadoodachaudhary/HomeFront` and `wadoodachaudhary/FlexKit`:
   install the Claude GitHub App on both, or run `/web-setup` once.
2. Create an environment named `HomeFront`:
   - **Network access**: Full, or Custom with the defaults plus `mcr.microsoft.com` and
     `*.data.mcr.microsoft.com` (the SQL Server image).
   - **Environment variables**: `HF_CLOUD=1` (plus `ASPNETCORE_ENVIRONMENT=Development` if you like;
     the launch profile sets it too). No password is needed: the SA password is generated
     inside each sandbox.
   - **`HF_GITLAB_TOKEN`**: a GitLab token with only `read_repository` that can read
     `application-modernization/hyphen-pb` and `flexkit` (a group access token with the Reporter
     role, or a personal access token). It lets each session bring in teammates' merged work.
     It cannot push, and the cloud never pushes to GitLab.
3. Connect the **Atlassian** connector at claude.ai/customize/connectors (Jira reads; writes only
   on the owner's explicit request).
   - **Setup script**:
     ```bash
     curl -fsSL https://raw.githubusercontent.com/wadoodachaudhary/vb6ToDotNet/main/tools/cloud/environment-setup.sh | bash
     ```

## Each session

The `SessionStart` hook in `.claude/settings.json` runs `session-hook.sh`. Only when
`HF_CLOUD=1` (or `CLAUDE_CODE_REMOTE=true`) does it start `session-start.sh` in the background. On
the Mac it does nothing. `session-start.sh`:

- places the app and FlexKit in the Mac's layout (from the setup cache), fast-forwarding them to GitHub;
- starts SQL Server and restores `HOMEFRONTSQL` (`db.sh`, checksum-verified);
- writes the app's `Database:Password` user-secret;
- runs `teammates.sh`: fetches GitLab `hyphen-pb` main and `flexkit` main (read-only) and
  merges teammates' work since the last deploy into the app's `main` and FlexKit's
  `telerik-parity-20260904`, file by file as `deploy_to_repos.sh --pull` does. Branch-owned files
  and deletions are only reported. Clean results are built and pushed to the GitHub mirrors;
  conflicts stay local with an `ATTENTION` line for the session to resolve. Each sync records a
  `GitLab-Synced: <repo> <sha>` trailer, so the same work is never merged twice;
- touches `/tmp/hf-cloud-ready` (log: `/tmp/hf-cloud-start.log`).

Query the database with `bash tools/cloud/sql.sh "SELECT ..."`. It never prints the password.

Browser checks: the setup script installs Google Chrome (the checks launch Playwright with
`channel: 'chrome'`), and the sandbox ships the `playwright` package in `/opt/node-tools`.
Run them with both `NODE_PATH=/opt/node-tools/node_modules` and
`PLAYWRIGHT_MODULE=/opt/node-tools/node_modules/playwright/test`. Most checks default to a Mac
path, and `playwright/test` exports `chromium`, `webkit` and `expect`.

If the log shows `could not read Username for 'https://github.com'` or `could not download`,
the session cannot reach the private repos. Claude can attach them with `add_repo`
(`access: push`) and rerun `bash tools/cloud/session-start.sh` in the foreground. If `add_repo`
answers "you don't have access", the Claude GitHub App is not installed on HomeFront and
FlexKit yet (step 1 above). Release assets are fetched through the REST API, because
`gh release download` uses GraphQL, which cloud sessions refuse.

## Jira tickets: cloud fixes, Mac tests

The `jira-ticket` skill (`.claude/skills/jira-ticket`) is the cloud routine: read the ticket
through the Atlassian connector, branch `jira/<KEY>` from the current main branch in each repo
the fix touches, verify, push the branch to GitHub, and report. Fixes never go straight to
`main` or `telerik-parity-20260904`.

On the Mac (the app and FlexKit each have a `github` remote):

```bash
bash tools/cloud/ticket.sh list            # ticket branches on GitHub, merged or still to test
bash tools/cloud/ticket.sh test HHM-1234   # check out jira/HHM-1234 in the app and/or FlexKit
bash tools/cloud/ticket.sh accept HHM-1234 # merge into main / telerik-parity-20260904, push to GitHub
bash tools/cloud/ticket.sh back            # return both repos to their main branches
```

Teammates' merges reach the GitHub mirrors through `teammates.sh`, so before deploying pull them
too (`/update-repositories` then judges teammate work by content, so nothing is applied twice):

```bash
git -C MobileSource/HomeFront pull --ff-only github main
git -C ../FlexKit pull --ff-only github telerik-parity-20260904
```

Every FlexKit change still needs its FlexCore port on the Mac. Cloud database changes stay in
that sandbox and are never written back to the Mac.

## Refreshing the database copy

On the Mac take a copy-only compressed backup, then attach it to a new private release. Update
`DB_TAG`, `DB_FILE` and `DB_SHA256` in `db.sh` and `environment-setup.sh`.
