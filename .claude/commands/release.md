---
name: release
description: Assess version bump, tag, publish to PyPI, and update changelog
argument-hint: "[patch|minor|major]"
disable-model-invocation: true
allowed-tools: Read, Edit, Grep, Bash(git status:*), Bash(git log:*), Bash(git diff:*), Bash(git add:*), Bash(git commit:*), Bash(git tag:*), Bash(git push:*), Bash(uv sync:*), Bash(uv run ruff:*), Bash(uv run ty:*), Bash(uv run pytest:*), Bash(gh release create:*), Bash(gh release delete:*), Bash(gh run list:*), Bash(gh run watch:*), Bash(gh run view:*), Bash(uvx --refresh:*)
---

## State

- Recent commits (tagged): !`git log --oneline --decorate -20`
- Latest tag: !`git tag --sort=-v:refname | head -1`
- Status: !`git status --short`
- Version: !`grep '^version' pyproject.toml`

## Steps

Stop and report if any step fails.

1. **Pre-flight**: Verify clean working tree and on `main`. If there are no commits since the last tag, report "Nothing to release" and stop.

2. **Assess bump**: Analyse all commits since the last tag:
   - `BREAKING CHANGE` or `!:` in any commit → `major`
   - Any `feat:` commit → `minor`
   - Only `fix:`, `chore:`, `docs:`, `refactor:` → `patch`
   - Take the highest applicable level.
   - If `$ARGUMENTS` is provided and **matches** the assessed level, proceed.
   - If `$ARGUMENTS` is provided but **disagrees** with the assessment, flag the mismatch with reasoning and ask the user to confirm or revise.
   - If `$ARGUMENTS` is empty, propose the assessed bump and **ask the user to confirm**.

3. **Compute version**: Parse current from `pyproject.toml`, apply bump. State `old → new`.

4. **Update versions** in exactly these 2 files:
   - `pyproject.toml` → `version = "X.Y.Z"`
   - `src/agent_xlsx/__init__.py` → `__version__ = "X.Y.Z"`

   Grep the old version string across `pyproject.toml` and `src/` to confirm zero remaining hits.

5. **Sync lockfile**: `uv sync --group dev`

6. **Changelog**: Add entry to `docs/src/app/changelog/page.mdx` below `# Changelog`, above the previous release entry. Review all commits since the previous tag. Follow the existing format and tone. Keep it concise; consolidate related fixes into single bullets.

7. **Validate**: `uv run ruff check && uv run ruff format --check && uv run ty check && uv run pytest -v`

8. **Commit**: Stage `pyproject.toml`, `src/agent_xlsx/__init__.py`, `uv.lock`, and `docs/src/app/changelog/page.mdx`. Commit message: `chore: release vX.Y.Z`

9. **Tag + push**: `git tag vX.Y.Z && git push origin main --tags`

10. **GitHub release**: Derive notes from the changelog entry. Use `gh release create vX.Y.Z --title "vX.Y.Z" --notes "..."` (note: the flag is `--notes`/`-n`, NOT `--body`).

11. **Monitor**: `gh run watch $(gh run list --limit 1 --json databaseId -q '.[0].databaseId')` — report final status. If CI fails after push, fix the issue, commit, then force-move the tag (`git tag -f vX.Y.Z && git push origin main --tags --force`), delete the GitHub release (`gh release delete vX.Y.Z --yes`), and recreate it to retrigger the publish workflow.

12. **Verify**: `uvx --refresh agent-xlsx@X.Y.Z --version` — confirm the published version resolves correctly from PyPI.

13. **Summary**: Old → new version, PyPI URL, GitHub release URL, CI status.
