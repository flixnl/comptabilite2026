#!/usr/bin/env python3
"""
dashboard/score.py — Helpers de formatage du score global.

Usage
─────
    from dashboard.score import menubar_icon, format_score_line

    icon = menubar_icon(score_label, m_a_verifier, n_a_appliquer)
    # ex: '📊!!', '📊!', '📊(3)', '📊[2]', '📊'

    line = format_score_line(score_label, score_emoji)
    # ex: 'Score : 🟢 OK'
"""
from __future__ import annotations


# ─── Icône menubar ────────────────────────────────────────────────────────────

def menubar_icon(
    score_label:  str,
    m_a_verifier: int = 0,
    n_a_appliquer: int = 0,
) -> str:
    """Calcule l'icône du menubar en combinant le score financier et les pendants.

    Priorités (de la plus haute à la plus basse) :
      1. score 'Risque'    → '📊!!'  (override de tout)
      2. score 'Attention' → '📊!'
      3. m > 0 (fichiers à vérifier) → '📊(m)'
      4. n > 0 (décisions à appliquer) → '📊[n]'
      5. sinon → '📊'

    Parameters
    ----------
    score_label    : 'OK' | 'Attention' | 'Risque'
    m_a_verifier   : nombre de fichiers en attente dans _à_vérifier
    n_a_appliquer  : nombre de décisions à appliquer en DÉPENSES
    """
    if score_label == 'Risque':
        return '📊!!'
    if score_label == 'Attention':
        return '📊!'
    if m_a_verifier > 0:
        return f'📊({m_a_verifier})'
    if n_a_appliquer > 0:
        return f'📊[{n_a_appliquer}]'
    return '📊'


# ─── Ligne de score formatée ──────────────────────────────────────────────────

def format_score_line(score_label: str, score_emoji: str) -> str:
    """Retourne une ligne lisible pour l'item de menu ou le rapport.

    Exemple : 'Score global : 🟢 OK'
    """
    return f'Score global : {score_emoji} {score_label}'


# ─── Résumé compact (pour notification / tooltip) ─────────────────────────────

def format_score_summary(
    score_label:  str,
    score_emoji:  str,
    n_alerts:     int,
    n_high:       int,
    n_medium:     int,
) -> str:
    """Résumé compact du score avec compteur d'alertes.

    Exemple : '🟡 Attention — 1 HIGH, 2 MEDIUM'
    """
    parts = []
    if n_high:
        parts.append(f'{n_high} HIGH')
    if n_medium:
        parts.append(f'{n_medium} MEDIUM')
    if not parts and n_alerts:
        parts.append(f'{n_alerts} alerte(s)')
    detail = ' — ' + ', '.join(parts) if parts else ''
    return f'{score_emoji} {score_label}{detail}'
