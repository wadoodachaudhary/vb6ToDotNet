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
- touches `/tmp/hf-cloud-ready` (log: `/tmp/hf-cloud-start.log`).

Query the database with `bash tools/cloud/sql.sh "SELECT ..."`. It never prints the password.

If the log shows `could not read Username for 'https://github.com'` or `could not download`,
the session cannot reach the private repos. Claude can attach them with `add_repo`
(`access: push`) and rerun `bash tools/cloud/session-start.sh` in the foreground. If `add_repo`
answers "you don't have access", the Claude GitHub App is not installed on HomeFront and
FlexKit yet (step 1 above). Release assets are fetched through the REST API, because
`gh release download` uses GraphQL, which cloud sessions refuse.

## Getting cloud work back to the Mac

The cloud pushes the app to GitHub `HomeFront` main and FlexKit to GitHub `telerik-parity-20260904`.
On the Mac the app and FlexKit repos each have a `github` remote:

```bash
git -C MobileSource/HomeFront pull --ff-only github main
git -C ../FlexKit pull --ff-only github telerik-parity-20260904
```

Pull both before deploying, or a GitLab push will not contain the cloud work. Cloud database
changes stay in that sandbox and are never written back to the Mac.

## Refreshing the database copy

On the Mac take a copy-only compressed backup, then attach it to a new private release. Update
`DB_TAG`, `DB_FILE` and `DB_SHA256` in `db.sh` and `environment-setup.sh`.
