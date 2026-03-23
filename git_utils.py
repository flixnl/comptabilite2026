#!/usr/bin/env python3
"""
dashboard/git_utils.py — Utilitaires Git robustes pour le menubar.

Fonctions
─────────
find_git_repo(start)     Remonte les parents pour trouver le répertoire .git
run_git_command(cmd)     Exécute une commande git avec timeout + logging propre
git_commit_and_push(…)   Workflow complet add → commit → push avec retry 1x

Règles de conception
────────────────────
• Ne jamais dépendre du cwd — toujours passer --git-dir et --work-tree explicitement.
• Robuste aux chemins iCloud NFC/NFD : find_git_repo() normalise les chemins.
• Push silencieux en cas de succès ; log court en cas d'erreur.
• Retry 1× après 2 secondes si le push échoue (réseau transitoire).
• Timeout 30 s par commande (évite les blocages indéfinis).
• Aucune raise propagée hors du module — toujours retourner (ok, message).
"""
from __future__ import annotations

import logging
import os
import subprocess
import time
import unicodedata
from typing import Optional

# ── Logger ────────────────────────────────────────────────────────────────────

logger = logging.getLogger(__name__)

# ── Constantes ────────────────────────────────────────────────────────────────

_GIT_TIMEOUT: int = 30          # secondes par commande
_RETRY_DELAY: float = 2.0       # secondes entre les essais


# ══════════════════════════════════════════════════════════════════════════════
#  find_git_repo
# ══════════════════════════════════════════════════════════════════════════════

def find_git_repo(start: Optional[str] = None) -> Optional[str]:
    """Remonte la hiérarchie depuis *start* pour trouver un répertoire .git.

    Robuste NFC/NFD : normalise chaque chemin en NFC avant de vérifier.

    Paramètres
    ----------
    start : chemin de départ (défaut : répertoire de ce module)

    Retourne
    --------
    Chemin absolu du répertoire contenant .git, ou None si introuvable.
    """
    if start is None:
        start = os.path.dirname(os.path.abspath(__file__))

    # Normaliser en NFC pour une comparaison cohérente
    current = unicodedata.normalize('NFC', os.path.abspath(start))

    # Monter jusqu'à la racine
    for _ in range(20):   # max 20 niveaux — évite les boucles infinies
        candidate_nfc = os.path.join(current, '.git')
        candidate_nfd = os.path.join(
            unicodedata.normalize('NFD', current), '.git')

        if os.path.isdir(candidate_nfc) or os.path.isdir(candidate_nfd):
            return current   # répertoire contenant .git

        parent = os.path.dirname(current)
        if parent == current:
            break   # racine du système de fichiers
        current = parent

    return None


# ══════════════════════════════════════════════════════════════════════════════
#  run_git_command
# ══════════════════════════════════════════════════════════════════════════════

def run_git_command(
    args:      list[str],
    repo_dir:  Optional[str] = None,
    timeout:   int = _GIT_TIMEOUT,
    log_path:  Optional[str] = None,
    quiet:     bool = False,
) -> tuple[bool, str]:
    """Exécute une commande git avec timeout et logging propre.

    Paramètres
    ----------
    args      : arguments git sans le binaire (ex: ['add', '-A'])
    repo_dir  : répertoire contenant .git (auto-détecté si None)
    timeout   : secondes avant abandon
    log_path  : chemin vers le fichier log (optionnel)
    quiet     : si True, ne pas loguer les erreurs (l'appelant gère lui-même)

    Retourne
    --------
    (ok: bool, message: str)
    """
    if repo_dir is None:
        repo_dir = find_git_repo()
    if not repo_dir:
        msg = 'find_git_repo: aucun dépôt Git trouvé'
        _log(log_path, f'❌ {msg}')
        return False, msg

    # Résoudre le .git (NFD sur macOS iCloud)
    git_dir_nfc = os.path.join(repo_dir, '.git')
    git_dir_nfd = os.path.join(
        unicodedata.normalize('NFD', repo_dir), '.git')
    git_dir = git_dir_nfd if os.path.isdir(git_dir_nfd) else git_dir_nfc

    work_tree = repo_dir

    cmd = [
        'git',
        '--git-dir',   git_dir,
        '--work-tree', work_tree,
    ] + args

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        out  = (result.stdout or '').strip()
        err  = (result.stderr or '').strip()
        combined = out or err or '(aucune sortie)'

        if result.returncode == 0:
            return True, combined
        else:
            msg = f'returncode={result.returncode} | {combined}'
            if not quiet:
                _log(log_path, f'❌ git {" ".join(args[:2])}: {msg}')
            return False, msg

    except subprocess.TimeoutExpired:
        msg = f'Timeout ({timeout}s) — git {" ".join(args[:2])}'
        _log(log_path, f'❌ {msg}')
        return False, msg
    except Exception as exc:
        msg = f'Erreur inattendue : {exc}'
        _log(log_path, f'❌ {msg}')
        return False, msg


# ══════════════════════════════════════════════════════════════════════════════
#  git_commit_and_push
# ══════════════════════════════════════════════════════════════════════════════

def git_commit_and_push(
    files:     Optional[list[str]] = None,
    message:   str = 'Scan auto — dashboard mis à jour',
    repo_dir:  Optional[str] = None,
    log_path:  Optional[str] = None,
    retry:     int = 1,
) -> tuple[bool, str]:
    """Workflow complet : git add → git commit → git push.

    Comportement intelligent :
    - Si files est fourni, n'ajoute QUE ces fichiers (jamais git add -A)
    - Détecte le dirty tree (fichiers dev modifiés non liés) et ne panique pas
    - Distingue "rien à committer" de "erreur réelle"
    - Convertit les chemins absolus en relatifs au repo

    Paramètres
    ----------
    files    : liste de chemins à ajouter (requis — None = erreur)
    message  : message de commit
    repo_dir : répertoire racine du repo (auto-détecté si None)
    log_path : fichier log optionnel
    retry    : nombre de tentatives supplémentaires si push échoue

    Retourne
    --------
    (ok: bool, message: str)
      ok=True  → push réussi (ou nothing to push)
      ok=False → erreur, voir message pour détail
    """
    if repo_dir is None:
        repo_dir = find_git_repo()
    if not repo_dir:
        return False, 'Aucun dépôt Git trouvé'

    _log(log_path, f'📌 Git commit+push — repo: {os.path.basename(repo_dir)}')

    # ── 0. Diagnostic du working tree ─────────────────────────────────────
    ok_st, status_out = run_git_command(
        ['status', '--porcelain'], repo_dir=repo_dir)
    if ok_st and status_out and status_out.strip():
        status_lines = [l for l in status_out.strip().split('\n') if l.strip()]
        n_modified = sum(1 for l in status_lines if l[:2].strip() in ('M', 'MM', 'AM'))
        n_untracked = sum(1 for l in status_lines if l.startswith('??'))
        if n_modified > 0 or n_untracked > 0:
            _log(log_path,
                 f'ℹ️  Working tree : {n_modified} modifié(s), {n_untracked} non-tracké(s) '
                 f'(dev local — ne bloque pas le push)')

    # ── 1. git add (fichiers explicites uniquement) ───────────────────────
    if not files:
        _log(log_path, 'ℹ️  Aucun fichier spécifié — skip git add')
        # Même sans add, tenter le push pour les commits précédents
    else:
        # Convertir les chemins absolus en relatifs au repo
        rel_files = []
        for f in files:
            if os.path.isabs(f):
                try:
                    rel = os.path.relpath(f, repo_dir)
                except ValueError:
                    rel = f  # Si sur un autre drive (Windows)
            else:
                rel = f
            rel_files.append(rel)

        ok, msg = run_git_command(
            ['add'] + rel_files, repo_dir=repo_dir, log_path=log_path)
        if not ok:
            return False, f'git add échoué : {msg}'

    # ── 2. git commit (quiet — on gère les messages nous-mêmes) ─────────
    ok, msg = run_git_command(
        ['commit', '-m', message],
        repo_dir=repo_dir,
        log_path=log_path,
        quiet=True,
    )
    if not ok:
        msg_lower = msg.lower()
        if 'nothing to commit' in msg_lower or 'no changes added' in msg_lower:
            _log(log_path, 'ℹ️  Rien à committer (dashboard inchangé)')
            # Continuer quand même pour pusher les commits précédents
        elif 'changes not staged' in msg_lower:
            # Fichiers dev modifiés localement — pas une erreur du pipeline
            _log(log_path,
                 'ℹ️  Fichiers dev modifiés localement — '
                 'seuls les fichiers générés sont committés')
            # Continuer pour tenter le push
        else:
            return False, f'git commit échoué : {msg}'
    else:
        _log(log_path, f'✅ Commit : {msg[:80]}')

    # ── 3. git push (avec retry) ───────────────────────────────────────────
    attempts = 1 + max(0, retry)
    for attempt in range(1, attempts + 1):
        ok, msg = run_git_command(
            ['push', 'origin', 'main'],
            repo_dir=repo_dir,
            log_path=log_path,
        )
        if ok:
            _log(log_path, '✅ Push réussi')
            return True, 'Push réussi'
        else:
            if attempt < attempts:
                _log(log_path,
                     f'⚠️  Push tentatif {attempt} échoué — retry dans {_RETRY_DELAY}s')
                time.sleep(_RETRY_DELAY)
            else:
                _log(log_path,
                     f'❌ Push échoué après {attempts} tentative(s) : {msg[:120]}')
                return False, f'Git non synchronisé (voir logs) — {msg[:80]}'

    return False, 'Push échoué (retry épuisé)'


# ══════════════════════════════════════════════════════════════════════════════
#  check_ssh_agent
# ══════════════════════════════════════════════════════════════════════════════

def check_ssh_agent() -> tuple[bool, str]:
    """Vérifie que ssh-agent est actif et qu'une clé est chargée.

    Retourne
    --------
    (ok: bool, message: str)
    """
    ssh_auth_sock = os.environ.get('SSH_AUTH_SOCK', '')
    if not ssh_auth_sock:
        return False, 'SSH_AUTH_SOCK non défini — ssh-agent inactif'
    if not os.path.exists(ssh_auth_sock):
        return False, f'SSH_AUTH_SOCK introuvable : {ssh_auth_sock}'

    try:
        result = subprocess.run(
            ['ssh-add', '-l'],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0:
            n = len(result.stdout.strip().splitlines())
            return True, f'{n} clé(s) chargée(s) dans ssh-agent'
        else:
            return False, 'Aucune clé dans ssh-agent (ssh-add ~/. ssh/id_ed25519)'
    except Exception as exc:
        return False, f'ssh-add -l échoué : {exc}'


# ══════════════════════════════════════════════════════════════════════════════
#  Helpers internes
# ══════════════════════════════════════════════════════════════════════════════

def _log(log_path: Optional[str], message: str) -> None:
    """Écrit dans le fichier log et dans le logger Python."""
    logger.debug(message)
    if log_path:
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                from datetime import datetime
                f.write(f'[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}\n')
        except OSError:
            pass
