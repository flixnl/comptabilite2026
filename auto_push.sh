#!/bin/bash
# auto_push.sh — Commit + push automatique du dashboard vers GitHub (SSH)
# Appelé par launchd (com.felix.dashboard-autopush.plist) ou manuellement.
#
# Prérequis :
#   • remote = git@github.com:flixnl/comptabilite2026.git  (SSH, pas HTTPS)
#   • ssh-agent actif avec clé chargée (normalement géré par macOS Keychain)

REPO_DIR="$HOME/Documents/Comptabilite/2026/_OUTILS/dashboard"
LOG="$REPO_DIR/auto_push.log"

cd "$REPO_DIR" || exit 1

# ── Nettoyage locks stale + objets temporaires ─────────────────────────────
find .git -name "*.lock" -delete 2>/dev/null
find .git/objects -name "tmp_obj_*" -delete 2>/dev/null

# ── Vérifier la clé SSH (macOS Keychain doit avoir chargé la clé) ──────────
if ! ssh-add -l &>/dev/null; then
    # Tenter de charger la clé depuis le Keychain macOS
    ssh-add --apple-use-keychain "$HOME/.ssh/id_ed25519" &>/dev/null 2>&1 || true
fi

# ── Committer les changements (data.json, index.html, + tout le reste) ─────
if ! git diff --quiet HEAD 2>/dev/null || git ls-files --others --exclude-standard | grep -q .; then
    git add -A
    git commit -m "Scan auto — dashboard mis à jour ($(date '+%b %d'))" >> "$LOG" 2>&1
    echo "[$(date '+%Y-%m-%d %H:%M')] 📝 Commit auto créé" >> "$LOG"
fi

# ── Vérifier s'il y a des commits à pusher ─────────────────────────────────
UNPUSHED=$(git log origin/main..HEAD --oneline 2>/dev/null)
if [ -z "$UNPUSHED" ]; then
    exit 0  # Rien à pousser — sortie silencieuse
fi

echo "[$(date '+%Y-%m-%d %H:%M')] Commits en attente:" >> "$LOG"
echo "$UNPUSHED" >> "$LOG"

# ── Push avec retry 1× ─────────────────────────────────────────────────────
push_once() {
    GIT_SSH_COMMAND="ssh -o BatchMode=yes -o ConnectTimeout=10" \
        git push origin main >> "$LOG" 2>&1
}

if push_once; then
    echo "[$(date '+%Y-%m-%d %H:%M')] ✅ Push réussi" >> "$LOG"
    exit 0
fi

# Retry après 2 secondes
sleep 2
if push_once; then
    echo "[$(date '+%Y-%m-%d %H:%M')] ✅ Push réussi (retry)" >> "$LOG"
    exit 0
fi

# Échec définitif — log court, pas de notification bloquante
echo "[$(date '+%Y-%m-%d %H:%M')] ❌ Push échoué — Git non synchronisé (voir auto_push.log)" >> "$LOG"
exit 1
