#!/usr/bin/env bash
set -euo pipefail

# Cross-platform installer for weekly-report-skill (macOS / Linux).
# Usage:
#   ./scripts/install.sh <target> [target ...] [--home DIR] [--dry-run]
# Targets:
#   pi | codex | codex-plugin | dsh | dsh-bundle |
#   antigravity | antigravity-skill | claude | claude-plugin | opencode | all

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
HOME_DIR="${HOME}"
DRY_RUN=0
TARGETS=()

while [[ $# -gt 0 ]]; do
  case "$1" in
    --home)
      HOME_DIR="$2"
      shift 2
      ;;
    --dry-run)
      DRY_RUN=1
      shift
      ;;
    -h|--help)
      sed -n '3,9p' "$0"
      exit 0
      ;;
    *)
      TARGETS+=("$1")
      shift
      ;;
  esac
done

if [[ ${#TARGETS[@]} -eq 0 ]]; then
  echo "Error: specify at least one target (pi, codex, dsh, antigravity, claude, all)." >&2
  exit 2
fi

copy_skill_core() {
  local dest="$1"
  if [[ "$DRY_RUN" -eq 1 ]]; then
    echo "[DryRun] Install skill core -> $dest"
    return
  fi

  mkdir -p "$dest/scripts" "$dest/templates"
  for f in LICENSE README.md SKILL.md; do
    [[ -f "$REPO_ROOT/$f" ]] && cp -f "$REPO_ROOT/$f" "$dest/$f"
  done

  for d in agents evals references tests workflows; do
    if [[ -d "$REPO_ROOT/$d" ]]; then
      rm -rf "$dest/$d"
      cp -R "$REPO_ROOT/$d" "$dest/$d"
      find "$dest/$d" -name '__pycache__' -type d -exec rm -rf {} + 2>/dev/null || true
      find "$dest/$d" -name '*.pyc' -type f -delete 2>/dev/null || true
    fi
  done

  for s in generate_report.py validate_payload.py; do
    [[ -f "$REPO_ROOT/scripts/$s" ]] && cp -f "$REPO_ROOT/scripts/$s" "$dest/scripts/$s"
  done

  # Preserve any custom company templates already placed in destination templates/.
  if [[ -f "$REPO_ROOT/templates/工作周报模板.xlsx" ]]; then
    cp -f "$REPO_ROOT/templates/工作周报模板.xlsx" "$dest/templates/工作周报模板.xlsx"
  fi
}

copy_wrapper() {
  local subdir="$1"
  local root="$2"
  local src="$REPO_ROOT/platforms/$subdir"
  if [[ ! -d "$src" ]]; then
    echo "Error: wrapper directory not found: $src" >&2
    exit 1
  fi
  if [[ "$DRY_RUN" -eq 1 ]]; then
    echo "[DryRun] Overlay wrapper ($subdir) -> $root"
    return
  fi
  mkdir -p "$root"
  cp -R "$src/." "$root/"
}

install_one() {
  local target="$1"
  case "$target" in
    pi)
      local dest="$HOME_DIR/.pi/agent/skills/weekly-report"
      copy_skill_core "$dest"
      echo "Installed [pi] -> $dest (run /reload in Pi)"
      ;;
    codex)
      local dest="$HOME_DIR/.codex/skills/weekly-report"
      copy_skill_core "$dest"
      echo "Installed [codex] -> $dest (includes agents/openai.yaml)"
      ;;
    codex-plugin)
      local root="$HOME_DIR/.codex/plugins/weekly-report"
      copy_wrapper "codex" "$root"
      copy_skill_core "$root/skills/weekly-report"
      echo "Installed [codex-plugin] -> $root"
      ;;
    dsh)
      local dest="$HOME_DIR/.dsh/skills/weekly-report"
      copy_skill_core "$dest"
      echo "Installed [dsh] -> $dest"
      ;;
    dsh-bundle)
      local root="$HOME_DIR/.dsh/local-plugins/dsh-weekly-report"
      copy_wrapper "dsh" "$root"
      copy_skill_core "$root/skills/weekly-report"
      echo "Installed [dsh-bundle] -> $root (register in DSH profile bundles)"
      ;;
    antigravity)
      local root="$HOME_DIR/.gemini/config/plugins/weekly-report"
      copy_wrapper "antigravity" "$root"
      copy_skill_core "$root/skills/weekly-report"
      echo "Installed [antigravity plugin] -> $root"
      ;;
    antigravity-skill)
      local dest="$HOME_DIR/.gemini/config/skills/weekly-report"
      copy_skill_core "$dest"
      echo "Installed [antigravity standalone skill] -> $dest"
      ;;
    claude)
      local dest="$HOME_DIR/.claude/skills/weekly-report"
      copy_skill_core "$dest"
      echo "Installed [claude] -> $dest"
      ;;
    claude-plugin)
      local root="$HOME_DIR/.claude/plugins/weekly-report"
      copy_wrapper "claude-code" "$root"
      copy_skill_core "$root/skills/weekly-report"
      echo "Installed [claude-plugin] -> $root"
      ;;
    opencode)
      local dest="$HOME_DIR/.config/opencode/superpowers/skills/weekly-report"
      copy_skill_core "$dest"
      echo "Installed [opencode] -> $dest"
      ;;
    all)
      for t in pi codex dsh antigravity claude; do
        install_one "$t"
      done
      ;;
    *)
      echo "Error: unknown target '$target'" >&2
      exit 2
      ;;
  esac
}

for t in "${TARGETS[@]}"; do
  install_one "$t"
done
