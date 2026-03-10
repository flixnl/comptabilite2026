#!/bin/bash
# auto_push.sh — Pousse les commits locaux non-pushés vers GitHub
# Appelé automatiquement par launchd toutes les heures

REPO_DIR="$HOME/Library/Mobile Documents/com~apple~CloudDocs/Comptabilité/2026/dashboard"
LOG="$REPO_DIR/auto_push.log"

cd "$REPO_DIR" || exit 1

# Supprimer les locks stale
find .git -name "*.lock" -delete 2>/dev/null

# Vérifier s'il y a des commits à pusher
UNPUSHED=$(git log origin/main..HEAD --oneline 2>/dev/null)

if [ -z "$UNPUSHED" ]; then
    exit 0
fi

echo "[$(date '+%Y-%m-%d %H:%M')] Commits à pusher:" >> "$LOG"
echo "$UNPUSHED" >> "$LOG"

git push origin main >> "$LOG" 2>&1

if [ $? -eq 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M')] ✅ Push réussi" >> "$LOG"
else
    echo "[$(date '+%Y-%m-%d %H:%M')] ❌ Push échoué" >> "$LOG"
fi
