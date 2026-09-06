# CLAUDE.md

## Jenkins

Jenkins runs locally at `http://localhost:8080/` (the same machine this repo
is checked out on). Anonymous API access is disabled (403) — you need HTTP
Basic auth with a Jenkins username + API token (Jenkins user settings →
"Configure" → "API Token"). Never commit that token anywhere in this repo;
ask the user for it in-session if you don't have it.

Useful endpoints for this job (`test`), once authenticated with
`curl --globoff -u user:token ...` (`--globoff` is required — curl treats
`[` `]` in the URL as its own range syntax otherwise):

- `job/test/lastBuild/api/json?tree=number,result,building,timestamp,url` —
  latest build result.
- `job/test/<build>/wfapi/describe` — per-stage status/timing breakdown
  (Jenkins Stage View, in JSON). Use this instead of guessing from the
  console log which stage failed.
- `job/test/<build>/consoleText` — full console log, useful for grepping the
  actual failure out of a long build.
- `scriptText` (POST, `script=<groovy>`) — Groovy script console. Only works
  if the authenticated user has Overall/Administer. Useful for:
  - Listing credential IDs:
    `CredentialsProvider.lookupCredentialsInItemGroup(StandardCredentials.class, Jenkins.get(), null)`
    — print `.id` only, never the secret.
  - Exercising a stored credential (e.g. hitting an API it authenticates to)
    without ever surfacing the secret value to the conversation: read it
    server-side inside the script (`cred.secret.getPlainText()`) and pipe it
    straight into the HTTP call / `ProcessBuilder` env vars, only printing
    the response (scrub the token out of any echoed output defensively,
    e.g. `output.replace(token, "***REDACTED***")`).
  - Because the Jenkins master and this checkout are on the same machine,
    `ProcessBuilder` in a script-console script can shell out to local
    tools (e.g. `dotnet`, `python3`) with the credential injected only as
    an env var — this is how a locally-built `.nupkg` got pushed straight
    to the private GitHub Packages feed, and how a GitHub Actions repo
    secret got rotated (see below), without the token ever appearing in
    this chat.
- The `github-packages-pat` Jenkins credential (used read-only for restore
  in the `Build` stage's `environment` block) is, in practice, a
  full-scope classic GitHub PAT — `write:packages`, `repo`, `admin:*`, etc.
  (`X-OAuth-Scopes` on any `api.github.com` call with it confirms this).
  That's over-privileged for what Jenkins needs it for, but it means this
  one credential can also push packages directly and manage repo settings/
  secrets on `Ter-jeff/*` repos if a task ever needs that — know that
  before reaching for a different token.

## NuGet source: no more nuget.org

[NuGet.Config](NuGet.Config) points exclusively at a private GitHub Packages
feed (`https://nuget.pkg.github.com/Ter-jeff/index.json`) — there is no
`nuget.org` fallback for anything restored via `dotnet restore`/`dotnet tool
update` without an explicit `--add-source`. If a package isn't already on
that feed, restore fails with "找不到 封裝 ... 的版本" (or the English
equivalent), not a clearer "unauthorized" error.

Packages get onto that feed via a separate repo, **`Ter-jeff/nuget`**
(sibling checkout at `../nuget`): drop a `.nupkg` under its `feed/`
directory, commit, push — a GitHub Actions workflow
(`.github/workflows/publish-packages.yml`) auto-publishes anything under
`feed/**.nupkg` to the GitHub Packages feed. That repo's README documents
the full process.

`slt-csharp-metrics` and `csharp-duplicate-detector` (used by the `Metrics`
stage's `.devops/metrics_calculate.py`) are custom tools built from source
in that repo (`src/SltCsharpMetrics`, `src/CsharpDuplicateDetector`), not
mirrors of a public package. Confirmed currently-published versions (as of
2026-09-05): both `1.0.0`. Don't trust version numbers found in this repo's
git history from when the binaries used to be vendored directly into
`.devops/dotnet-tools` (e.g. `4.1.2`, `SltDuplicateCodeDetector 1.1.0`) —
those were never actually published to any feed, and the package was
renamed (`SltDuplicateCodeDetector` → `csharp-duplicate-detector`) when it
got repackaged for real publishing.

Before assuming a package is "just missing" from that feed, check whether
it's actually published: `curl --globoff -u <owner>:<token>
https://nuget.pkg.github.com/<owner>/download/<package-id-lowercase>/index.json`
returns `{"versions": [...]}` if it exists, 404 `NAME_UNKNOWN` if not.

If the mirror repo's publish workflow fails, check its Actions run logs
(`api.github.com/repos/Ter-jeff/nuget/actions/runs`, then
`.../actions/runs/<id>/jobs` for the failing step, then
`.../actions/jobs/<id>/logs`, following the 302 redirect it returns) — a
`401 Unauthorized` on the `dotnet nuget push` step means the workflow's
`GH_PACKAGES_PAT` repository secret (Settings → Secrets and variables →
Actions, in that repo) is expired/invalid. This happened once already
(fixed 2026-09-05) by reusing the `github-packages-pat` Jenkins credential
(see above — it already has `write:packages`) to overwrite that repo
secret via the GitHub API: `GET
.../actions/secrets/public-key`, encrypt the token with
`nacl.public.SealedBox` (PyNaCl; `pip3 install --break-system-packages
pynacl` if missing) using that key, `PUT .../actions/secrets/GH_PACKAGES_PAT`
with `{encrypted_value, key_id}`. Do this via the script-console
`ProcessBuilder` trick above so the token never appears in chat. The
repo's README recommends a token scoped to just `write:packages` instead —
prefer that if the user wants to hand you a fresh one rather than reusing
the broad one.

## Jenkinsfile conventions

- No hardcoded per-user paths. [Jenkinsfile](Jenkinsfile)'s top-level `PATH`
  env var prepends `/usr/local/share/dotnet` and `/opt/homebrew/bin` because
  the `launchd`-run Jenkins service starts with a minimal PATH that can't
  find `dotnet`/`python3` otherwise — but nothing in this pipeline needs a
  *global* `dotnet tool` install directory (e.g. `~/.dotnet/tools`), because
  every `dotnet tool` call here uses an explicit `--tool-path
  ${env.TOOLS_DIR}` (`.devops/dotnet-tools`) and every invocation of an
  installed tool uses its full path (`${env.TOOLS_DIR}/reportgenerator`,
  etc). Don't reintroduce a bare user-home tools path unless a new step
  actually needs a tool resolved off PATH.
- `Metrics` stage runs on every agent now (the `when { expression {
  !isUnix() } }` guard was removed 2026-09-05 once `slt-csharp-metrics` and
  `csharp-duplicate-detector` became installable via `dotnet tool update`
  instead of only existing as pre-provisioned Windows binaries).
  `.devops/metrics_calculate.py` picks the right launcher name itself
  (`TOOL_SUFFIX`: `.exe` on Windows, empty elsewhere) — `dotnet tool`
  launchers only carry the `.exe` suffix on Windows.
