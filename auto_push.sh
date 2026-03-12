#!/bin/bash
# auto_push.sh — Commit + push automatique du dashboard vers GitHub
# Appelé automatiquement par launchd toutes les heures

REPO_DIR="$HOME/Library/Mobile Documents/com~apple~CloudDocs/Comptabilité/2026/_OUTILS/dashboard"
LOG="$REPO_DIR/auto_push.log"

cd "$REPO_DIR" || exit 1

# Supprimer les locks stale et tmp objects
find .git -name "*.lock" -delete 2>/dev/null
find .git/objects -name "tmp_obj_*" -delete 2>/dev/null

# Commiter les changements non-commités (data.json, index.html)
if ! git diff --quiet data.json index.html 2>/dev/null; then
    git add data.json index.html
    git commit -m "Scan auto — dashboard mis à jour ($(date '+%b %d'))" >> "$LOG" 2>&1
    echo "[$(date '+%Y-%m-%d %H:%M')] 📝 Commit auto créé" >> "$LOG"
fi

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
