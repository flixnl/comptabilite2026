#!/usr/bin/env python3
"""
dashboard/alerts_view.py — Formatage des alertes et données financières
                            pour le menubar et les notifications.

Usage
─────
    from dashboard.alerts_view import (
        format_alerts_for_menubar,
        format_finances_lines,
        format_echeances_lines,
    )

    alert_lines = format_alerts_for_menubar(ctx.alerts, max_items=4)
    fin_lines   = format_finances_lines(ctx)
    ech_lines   = format_echeances_lines(ctx)
"""
from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from dashboard.context import CockpitContext
    from analytics.alerts import Alert


# ─── Alertes → lignes menubar ─────────────────────────────────────────────────

def format_alerts_for_menubar(
    alerts: list['Alert'],
    max_items: int = 4,
) -> list[str]:
    """Retourne jusqu'à max_items lignes d'alertes pour le menubar.

    Tronque intelligemment : conserve d'abord les HIGH, puis les MEDIUM.
    Chaque ligne est ≤ 60 caractères (troncature avec '…').

    Returns
    -------
    list[str] — lignes prêtes à mettre dans des rumps.MenuItem
    """
    if not alerts:
        return ['✅ Aucune alerte']

    # Séparer par sévérité
    high   = [a for a in alerts if a.severity == 'HIGH']
    medium = [a for a in alerts if a.severity == 'MEDIUM']
    low    = [a for a in alerts if a.severity == 'LOW']

    selected: list['Alert'] = []
    for bucket in (high, medium, low):
        remaining = max_items - len(selected)
        if remaining <= 0:
            break
        selected.extend(bucket[:remaining])

    lines = []
    for a in selected:
        text = f'{a.icon} {a.message}'
        if len(text) > 60:
            text = text[:59] + '…'
        lines.append(text)

    # Résumé si des alertes ont été tronquées
    total = len(alerts)
    shown = len(selected)
    if total > shown:
        remaining = total - shown
        lines.append(f'… +{remaining} autre(s) alerte(s)')

    return lines


# ─── Finances → lignes menubar ────────────────────────────────────────────────

def format_finances_lines(ctx: 'CockpitContext') -> list[str]:
    """Retourne les lignes de la section 💰 Finances du menubar.

    Contenu :
      • Revenus encaissés YTD (ST)
      • AR non encaissé
      • TPS + TVQ collectées
      • Obligations fiscales estimées

    Returns
    -------
    list[str] — lignes textuelles (pas de rumps.MenuItem)
    """
    lines: list[str] = []
    ac = ctx.alert_ctx

    # Revenus
    if ac.metrics:
        m = ac.metrics
        enc = f'{m.encaisse_st:,.0f} $'
        lines.append(f'Revenus encaissés : {enc}')
        if m.non_encaisse_st > 0:
            lines.append(f'Non encaissé : {m.non_encaisse_st:,.0f} $')

    # AR
    if ac.ar and ac.ar.count > 0:
        lines.append(f'AR ({ac.ar.count} fact.) : {ac.ar.total_st:,.0f} $')

    # Taxes — source de vérité unique : TaxesPosition
    # Aucun fallback sur ac.fiscal — les taxes brutes sans déduction des
    # paiements (tps_collectee + tps_a_recevoir) seraient un chiffre faux.
    if ac.taxes_position is not None:
        pos = ac.taxes_position
        lines.append(f'Taxes restantes : {pos.total_restantes:,.2f} $')

    # Obligations fiscales (impôt + cotisations)
    if ac.tax_estimate:
        te = ac.tax_estimate
        lines.append(f'Obligations fiscales : {te.total_obligations:,.0f} $')

    if not lines:
        lines.append('Données financières indisponibles')

    return lines


# ─── Échéances → lignes menubar ───────────────────────────────────────────────

def format_echeances_lines(ctx: 'CockpitContext') -> list[str]:
    """Retourne les lignes de la section 📅 Prochaines échéances du menubar.

    Affiche :
      • Prochain acompte provisionnel (date + montant + statut)
      • AR le plus ancien (si > 30 j)

    Returns
    -------
    list[str]
    """
    lines: list[str] = []
    ac = ctx.alert_ctx

    # Acomptes provisionnels
    if ac.installments and ac.installments.required and ac.next_installment:
        nxt = ac.next_installment
        p   = nxt.period
        status_icon = {'EN RETARD': '🔴', 'IMMINENT': '🟡', 'À VENIR': '📅'}.get(
            nxt.status, '📅')
        days_str = ''
        if nxt.days_left is not None:
            if nxt.status == 'EN RETARD':
                days_str = f' ({abs(nxt.days_left)}j de retard)'
            else:
                days_str = f' ({nxt.days_left}j)'
        lines.append(
            f'{status_icon} Acompte {p.due_date} : {p.amount:,.0f} ${days_str}')
    elif ac.installments and not ac.installments.required:
        lines.append('📅 Pas d\'acomptes requis cette année')

    # AR le plus ancien
    if ac.ar and ac.ar.entries:
        ref = ctx.reference_date
        oldest = None
        oldest_age = 0
        for e in ac.ar.entries:
            age = e.age_days(ref)
            if age is not None and age > oldest_age:
                oldest_age = age
                oldest = e
        if oldest and oldest_age > 30:
            lines.append(
                f'⏳ {oldest.ref} — {oldest_age}j ({oldest.client[:20]})')

    if not lines:
        lines.append('📅 Aucune échéance imminente')

    return lines
