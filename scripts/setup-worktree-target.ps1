<#
Share one target/ dir across the main checkout and every git worktree of
this repo (windows). For contributors who don't use direnv or
`nix develop` -- both of those already set CARGO_TARGET_DIR dynamically,
see .envrc and flake.nix. This script gets the same effect via
.cargo/config.toml's [build] target-dir for anyone running plain `cargo`
outside of direnv/nix.

Run once per worktree, from inside that worktree:
  scripts/setup-worktree-target.ps1

Safe to re-run (idempotent) and safe to run in a worktree that already
gets CARGO_TARGET_DIR from direnv/nix -- CARGO_TARGET_DIR, if set, takes
precedence over this file's target-dir, so this is inert for those users.

NOTE: this edits .cargo/config.toml *in the worktree you run it from*.
That file is git-tracked, so this intentionally makes it locally dirty
(git status will show it modified). This is expected and is why the
tracked copy of .cargo/config.toml deliberately omits target-dir -- see
the comment at the top of that file. Do not commit this edit; it bakes
in a machine-specific absolute path.
#>

$ErrorActionPreference = "Stop"

$repoRoot = (git rev-parse --show-toplevel).Trim()
Set-Location $repoRoot

$gitCommonDir = (git rev-parse --git-common-dir).Trim()
# Resolve symlinks the same way .envrc's `realpath` does.
$commonDirAbs = (Resolve-Path -LiteralPath $gitCommonDir).Path
# Use forward slashes so the TOML value matches the .sh script's output
# byte-for-byte when both compute the same underlying path, and because
# Cargo accepts forward slashes in target-dir on Windows.
$targetDir = ((Split-Path -Parent $commonDirAbs) -replace '\\', '/') + "/target"

$configDir = ".cargo"
$configFile = ".cargo/config.toml"
if (-not (Test-Path -LiteralPath $configDir)) {
  New-Item -ItemType Directory -Path $configDir | Out-Null
}
if (-not (Test-Path -LiteralPath $configFile)) {
  New-Item -ItemType File -Path $configFile | Out-Null
}

$content = Get-Content -LiteralPath $configFile -Raw
if ($null -eq $content) { $content = "" }

$targetLine = "target-dir = `"$targetDir`""

if ($content.Contains($targetLine)) {
  Write-Host "Already set: [build] $targetLine in $configFile"
  exit 0
}

if ($content -match '(?m)^target-dir\s*=.*$') {
  # Stale target-dir line present (e.g. copied from a different machine or
  # a moved repo) -- replace it in place rather than duplicating the key.
  # Using a MatchEvaluator (scriptblock) instead of a replacement string
  # avoids $targetLine's `$` and backslashes being reinterpreted as regex
  # replacement syntax.
  $newContent = [regex]::Replace($content, '(?m)^target-dir\s*=.*$', { param($m) $targetLine })
  Set-Content -LiteralPath $configFile -Value $newContent -NoNewline
  Write-Host "Updated: [build] $targetLine in $configFile"
} elseif ($content -match '(?m)^\[build\]\s*$') {
  # [build] table exists but has no target-dir key -- insert right after
  # the table header.
  $newContent = [regex]::Replace($content, '(?m)^\[build\]\s*$', { param($m) "[build]`n$targetLine" }, 1)
  Set-Content -LiteralPath $configFile -Value $newContent -NoNewline
  Write-Host "Added to existing [build] table: $targetLine in $configFile"
} else {
  # No [build] table at all -- append a new one.
  $append = "`n# Local-only: shared target dir across worktrees of this repo.`n# Set by scripts/setup-worktree-target.sh or .ps1 -- do not commit this edit.`n[build]`n$targetLine`n"
  Add-Content -LiteralPath $configFile -Value $append -NoNewline
  Write-Host "Appended new [build] table: $targetLine in $configFile"
}
