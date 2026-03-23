#!/usr/bin/env python3
"""
dashboard/menubar_format.py — Formatage lisible des données cockpit / autopilot
                               pour les items du menubar macOS (rumps).

Fonctions exportées
-------------------
    format_human_status(score_label)          → str   # 🟢 Tout va bien …
    format_due_date_relative(days_left, status) → str # «dans 12j» / «EN RETARD»
    format_cashflow_status(ctx)               → str   # 'OK' | 'serré' | 'risque' | 'inconnu'
    format_decision_line(actions, line_index) → str   # action prioritaire
    format_action_summary(top, line_index)    → str   # alias de format_decision_line
    format_autopilot_simple(result)           → str   # 🟢 OK — 2 anomalies …
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from dashboard.context import CockpitContext


# ─── Score → texte humain ─────────────────────────────────────────────────────

_SCORE_MAP: dict[str, str] = {
    'OK':        '🟢 Tout va bien',
    'Attention': '🟡 Attention requise',
    'Risque':    '🔴 Risque détecté',
    'Critique':  '🔴 Situation critique',
}


def format_human_status(score_label: str) -> str:
    """Convertit un score_label en texte humain pour le menubar.

    Exemples
    --------
    'OK'        → '🟢 Tout va bien'
    'Attention' → '🟡 Attention requise'
    'Risque'    → '🔴 Risque détecté'
    """
    return _SCORE_MAP.get(score_label, f'⚪ {score_label}')


# ─── Date relative ────────────────────────────────────────────────────────────

def format_due_date_relative(days_left: int | None, status: str) -> str:
    """Retourne une chaîne de date relative courte pour le menubar.

    Exemples
    --------
    days_left=12, status='À VENIR'   → 'dans 12j'
    days_left=0,  status='IMMINENT'  → 'aujourd\'hui'
    days_left=-3, status='EN RETARD' → '3j de retard'
    days_left=None, status=…         → ''
    """
    if days_left is None:
        return ''

    status_up = status.upper() if status else ''

    if status_up == 'EN RETARD':
        n = abs(int(days_left))
        return f'{n}j de retard'
    if days_left == 0:
        return "aujourd'hui"
    if days_left == 1:
        return 'demain'
    if days_left < 0:
        return f'{abs(int(days_left))}j de retard'
    return f'dans {int(days_left)}j'


# ─── Trésorerie 30j ───────────────────────────────────────────────────────────

def format_cashflow_status(ctx: 'CockpitContext') -> str:
    """Retourne un statut de trésorerie parmi : 'OK' | 'serré' | 'risque' | 'inconnu'.

    Logique
    -------
    • Cherche dans les alertes actives si une alerte de type cashflow / trésorerie
      est présente et détermine le niveau.
    • Fallback vers les métriques si disponibles.
    """
    try:
        # Lecture des alertes — priorité aux alertes HIGH/MEDIUM sur la trésorerie
        alerts = getattr(ctx, 'alerts', []) or []
        for a in alerts:
            msg = (getattr(a, 'message', '') or '').lower()
            sev = (getattr(a, 'severity', '') or '').upper()
            cat = (getattr(a, 'category', '') or '').lower()
            if 'trésorerie' in msg or 'cashflow' in msg or 'cash' in cat:
                if sev == 'HIGH':
                    return 'risque'
                if sev == 'MEDIUM':
                    return 'serré'

        # Lecture des métriques financières
        ac = getattr(ctx, 'alert_ctx', None)
        if ac is None:
            return 'inconnu'

        metrics = getattr(ac, 'metrics', None)
        if metrics is None:
            return 'inconnu'

        encaisse  = float(getattr(metrics, 'encaisse_st', 0) or 0)
        non_enc   = float(getattr(metrics, 'non_encaisse_st', 0) or 0)

        # Taxes restantes
        taxes_pos = getattr(ac, 'taxes_position', None)
        taxes_due = 0.0
        if taxes_pos is not None:
            taxes_due = float(getattr(taxes_pos, 'total_restantes', 0) or 0)

        # Obligations fiscales
        tax_est = getattr(ac, 'tax_estimate', None)
        oblig   = 0.0
        if tax_est is not None:
            oblig = float(getattr(tax_est, 'total_obligations', 0) or 0)

        charges = taxes_due + oblig
        liquide = encaisse + non_enc * 0.5   # AR pondéré à 50%

        if charges <= 0:
            return 'OK'
        ratio = liquide / charges
        if ratio >= 1.2:
            return 'OK'
        if ratio >= 0.8:
            return 'serré'
        return 'risque'

    except Exception:
        return 'inconnu'


# ─── Actions prioritaires ─────────────────────────────────────────────────────

def format_action_summary(top: list[Any], line_index: int = 0) -> str:
    """Retourne la ligne d'action à l'index donné, ou une chaîne vide.

    Paramètres
    ----------
    top        : liste d'objets action (issus de get_top_actions)
    line_index : 0 pour la première ligne, 1 pour la seconde

    L'objet action est supposé avoir un attribut .label, .title, ou .message.
    """
    if not top or line_index >= len(top):
        if line_index == 0:
            return '   ✅ Aucune action prioritaire'
        return '   —'

    action = top[line_index]

    # Chercher le texte dans plusieurs attributs possibles
    text = (
        getattr(action, 'label', None)
        or getattr(action, 'title', None)
        or getattr(action, 'message', None)
        or getattr(action, 'description', None)
        or str(action)
    )

    # Préfixe d'urgence
    priority = (getattr(action, 'priority', None) or
                getattr(action, 'severity', None) or '').upper()
    prefix = {'HIGH': '🔴', 'MEDIUM': '🟡', 'LOW': '⚪'}.get(priority, '▸')

    text = text.strip()
    if len(text) > 55:
        text = text[:54] + '…'

    return f'   {prefix} {text}'


# Alias — certaines versions du code appellent format_decision_line
format_decision_line = format_action_summary


# ─── Autopilot → résumé simple ───────────────────────────────────────────────

def format_autopilot_simple(result: Any) -> str:
    """Retourne un résumé one-liner de l'AutopilotResult pour le menubar.

    Exemple : '🟢 OK — 0 anomalie'
              '🟡 Attention — 1 critique, 2 moyens'
              '🔴 Risque — 3 critiques'
    """
    try:
        emoji = getattr(result, 'score_emoji', '⚪') or '⚪'
        label = getattr(result, 'score_label', 'Inconnu') or 'Inconnu'
        n_h   = int(getattr(result, 'n_high',   0) or 0)
        n_m   = int(getattr(result, 'n_medium', 0) or 0)
        n_l   = int(getattr(result, 'n_low',    0) or 0)
        total = n_h + n_m + n_l

        if total == 0:
            detail = '0 anomalie'
        else:
            parts = []
            if n_h:
                parts.append(f'{n_h} critique{"s" if n_h > 1 else ""}')
            if n_m:
                parts.append(f'{n_m} moyen{"s" if n_m > 1 else ""}')
            if n_l:
                parts.append(f'{n_l} mineur{"s" if n_l > 1 else ""}')
            detail = ', '.join(parts)

        err = getattr(result, 'error', None) or ''
        if err:
            detail = f'erreur : {str(err)[:40]}'

        return f'{emoji} {label} — {detail}'

    except Exception as exc:
        return f'⚪ Autopilot : erreur ({exc})'
