#!/usr/bin/env python3
"""
dashboard/html_writer.py — Renderer HTML depuis CockpitContext.

Produit le même dashboard HTML mobile-first que generate_dashboard.py,
mais lit TOUTES les données depuis un CockpitContext déjà construit.
Aucune lecture Excel directe — zéro openpyxl ici.

Usage
─────
    from dashboard.context import build_cockpit_context
    from dashboard.html_writer import write_html_dashboard

    ctx = build_cockpit_context(ventes_path, depenses_path)
    write_html_dashboard(ctx, '/tmp/dashboard.html')

Mapping CockpitContext → variables HTML
────────────────────────────────────────
    AR total              → ctx.alert_ctx.ar.total_st
    Revenus encaissés     → ctx.alert_ctx.metrics.encaisse_st
    À recevoir (AR)       → ctx.alert_ctx.ar.total_st
    TPS/TVQ dues          → ctx.alert_ctx.taxes_position.total_restantes
    Obligations           → ctx.obligations
    Revenus par mois      → ctx.monthly_revenue
    Dépenses par catég.   → ctx.expenses_by_category
    Bilan 2025            → ctx.fiscal_2025_summary
    Estimation fiscale    → ctx.alert_ctx.tax_estimate

Limitations v2.2
─────────────────
    - impots_retenus / acomptes non exposés dans TaxEstimate : le solde fiscal
      affiché est brut (charges totales sans déduction retenues à la source).
"""
from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from dashboard.context import CockpitContext

# ── Constantes ────────────────────────────────────────────────────────────────

ANNEE_DEFAUT = 2026

MOIS_ORDER = ['JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
              'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE']

MOIS_ABBR = {
    'JANVIER': 'JAN', 'FÉVRIER': 'FÉV', 'MARS': 'MAR', 'AVRIL': 'AVR',
    'MAI': 'MAI', 'JUIN': 'JUIN', 'JUILLET': 'JUIL', 'AOÛT': 'AOÛ',
    'SEPTEMBRE': 'SEP', 'OCTOBRE': 'OCT', 'NOVEMBRE': 'NOV', 'DÉCEMBRE': 'DÉC',
}


# ── Helpers de formatage ──────────────────────────────────────────────────────

def fmt_cad(v):
    """Formate un montant en dollars CAD avec séparateurs."""
    if v is None or (isinstance(v, (int, float)) and v == 0):
        return '—'
    try:
        val = float(v)
        if val == 0:
            return '—'
        formatted = f'{abs(val):,.2f}'.replace(',', '\u202f')
        sign = '−' if val < 0 else ''
        return f'{sign}{formatted}\u00a0$'
    except Exception:
        return str(v)


def color_for(value, positive_is_good=True):
    """Retourne une couleur CSS selon le signe de la valeur."""
    try:
        v = float(value)
    except Exception:
        return '#94a3b8'
    if positive_is_good:
        return '#34d399' if v >= 0 else '#f87171'
    else:
        return '#f87171' if v > 0 else '#34d399'


# ── Graphique SVG ─────────────────────────────────────────────────────────────

def bar_chart_revenus(revenus_ventes: dict, a_recevoir_par_mois: dict | None = None) -> str:
    """Génère un graphique en barres SVG revenus par mois.

    revenus_ventes  : {mois_int: {'freelance': float, 'salaire': float}}
    a_recevoir_par_mois : {mois_int: float} — affiché pâle/translucide
    """
    COL_FREELANCE  = '#6366f1'
    COL_SALAIRE    = '#22d3ee'
    COL_DEPENSES   = '#fb923c'   # inutilisé ici (pas de dépenses mensuelles)
    COL_A_RECEVOIR = '#a78bfa'

    MOIS_ABBR_LOCAL = {
        'JANVIER': 'JAN', 'FÉVRIER': 'FÉV', 'MARS': 'MAR', 'AVRIL': 'AVR',
        'MAI': 'MAI', 'JUIN': 'JUN', 'JUILLET': 'JUL', 'AOÛT': 'AOÛ',
        'SEPTEMBRE': 'SEP', 'OCTOBRE': 'OCT', 'NOVEMBRE': 'NOV', 'DÉCEMBRE': 'DÉC',
    }

    ar_mois = a_recevoir_par_mois or {}
    max_val = 1
    data = []
    annee = None
    if revenus_ventes:
        k = next(iter(revenus_ventes))
        annee = k // 100

    for i, mois in enumerate(MOIS_ORDER):
        if annee:
            mois_int = annee * 100 + (i + 1)
        else:
            mois_int = ANNEE_DEFAUT * 100 + (i + 1)
        rev = revenus_ventes.get(mois_int, {})
        freelance = rev.get('freelance', 0)
        salaire   = rev.get('salaire', 0)
        a_rec     = ar_mois.get(mois_int, 0)
        total_rev = freelance + salaire + a_rec
        max_val   = max(max_val, total_rev)
        data.append((MOIS_ABBR_LOCAL[mois], freelance, salaire, a_rec))

    w, h = 800, 280
    pad_left, pad_right = 60, 20
    pad_top, pad_bottom = 36, 36
    chart_w = w - pad_left - pad_right
    chart_h = h - pad_top - pad_bottom
    bar_group_w = chart_w / 12
    bar_w = bar_group_w * 0.55
    y_base = pad_top + chart_h

    legend = f"""
    <g transform="translate({pad_left}, 14)">
      <rect x="0" y="-9" width="12" height="12" rx="3" fill="{COL_FREELANCE}"/>
      <text x="16" y="1" fill="#cbd5e1" font-size="12" font-family="-apple-system,system-ui,sans-serif">Freelance</text>
      <rect x="100" y="-9" width="12" height="12" rx="3" fill="{COL_SALAIRE}"/>
      <text x="116" y="1" fill="#cbd5e1" font-size="12" font-family="-apple-system,system-ui,sans-serif">Salaire</text>
      <rect x="190" y="-9" width="12" height="12" rx="3" fill="{COL_A_RECEVOIR}" opacity="0.4"/>
      <rect x="190" y="-9" width="12" height="12" rx="3" fill="none" stroke="{COL_A_RECEVOIR}" stroke-width="1" stroke-dasharray="3,2" opacity="0.85"/>
      <text x="206" y="1" fill="#cbd5e1" font-size="12" font-family="-apple-system,system-ui,sans-serif">À recevoir</text>
    </g>"""

    grids = ''
    for i in range(5):
        y   = y_base - (i / 4) * chart_h
        val = (i / 4) * max_val
        grids += f'<line x1="{pad_left}" y1="{y:.0f}" x2="{w - pad_right}" y2="{y:.0f}" stroke="#334155" stroke-width="0.5"/>'
        grids += f'<text x="{pad_left - 8}" y="{y + 4:.0f}" text-anchor="end" fill="#64748b" font-size="10" font-family="-apple-system,system-ui,sans-serif">{val:,.0f}$</text>'

    bars = ''
    for i, (label, freelance, salaire, a_rec) in enumerate(data):
        x_bar = pad_left + i * bar_group_w + bar_group_w * 0.15

        cursor_y = y_base
        if salaire > 0:
            sal_h = (salaire / max_val) * chart_h
            cursor_y -= sal_h
            rx = '3' if freelance == 0 and a_rec == 0 else '0'
            bars += f'<rect x="{x_bar:.1f}" y="{cursor_y:.1f}" width="{bar_w:.1f}" height="{sal_h:.1f}" rx="{rx}" fill="{COL_SALAIRE}" opacity="0.9"/>'
        if freelance > 0:
            free_h = (freelance / max_val) * chart_h
            cursor_y -= free_h
            rx = '3' if a_rec == 0 else '0'
            bars += f'<rect x="{x_bar:.1f}" y="{cursor_y:.1f}" width="{bar_w:.1f}" height="{free_h:.1f}" rx="{rx}" fill="{COL_FREELANCE}" opacity="0.9"/>'
        if a_rec > 0:
            ar_h = (a_rec / max_val) * chart_h
            cursor_y -= ar_h
            bars += f'<rect x="{x_bar:.1f}" y="{cursor_y:.1f}" width="{bar_w:.1f}" height="{ar_h:.1f}" rx="3" fill="{COL_A_RECEVOIR}" opacity="0.30"/>'
            bars += f'<rect x="{x_bar:.1f}" y="{cursor_y:.1f}" width="{bar_w:.1f}" height="{ar_h:.1f}" rx="3" fill="none" stroke="{COL_A_RECEVOIR}" stroke-width="1.2" stroke-dasharray="3,2" opacity="0.80"/>'

        lx = pad_left + i * bar_group_w + bar_group_w * 0.5
        bars += f'<text x="{lx:.1f}" y="{h - 8}" text-anchor="middle" fill="#94a3b8" font-size="10" font-weight="600" font-family="-apple-system,system-ui,sans-serif">{label}</text>'

    return f"""<svg viewBox="0 0 {w} {h}" style="width:100%;height:auto;display:block;margin:8px 0">
      {legend}{grids}{bars}
    </svg>"""


# ── Annexe 2025 ───────────────────────────────────────────────────────────────

def _build_annexe_2025(fiscal_summary: dict, annee: int) -> str:
    """Génère le bloc HTML de l'annexe 2025 depuis ctx.fiscal_2025_summary."""
    if not fiscal_summary or not fiscal_summary.get('total_charges', 0):
        return ''

    d = fiscal_summary
    profit      = (d.get('rev_total') or 0) - (d.get('depenses') or 0)
    solde_color = '#f87171' if (d.get('solde_du') or 0) > 0 else '#34d399'

    return f'''
  <div style="margin-top:32px;border-top:2px solid #374151;padding-top:24px">
    <p class="section-title" style="color:#94a3b8">Annexe — Bilan fiscal 2025 (année complète)</p>
    <div class="kpi-row" style="grid-template-columns:repeat(3,1fr);margin-bottom:16px">
      <div class="kpi">
        <div class="label">Revenus bruts 2025</div>
        <div class="value sm" style="color:#60a5fa">{fmt_cad(d.get('rev_total'))}</div>
        <div style="font-size:11px;color:#6b7280;margin-top:4px">Freelance {fmt_cad(d.get('rev_freelance'))} · Salaire {fmt_cad(d.get('rev_salarie'))}</div>
      </div>
      <div class="kpi">
        <div class="label">Dépenses déductibles 2025</div>
        <div class="value sm" style="color:var(--orange)">{fmt_cad(d.get('depenses'))}</div>
      </div>
      <div class="kpi">
        <div class="label">Profit brut 2025</div>
        <div class="value sm" style="color:#34d399">{fmt_cad(profit)}</div>
      </div>
    </div>
    <div class="card" style="background:#1e2433">
      <table>
        <thead>
          <tr>
            <th>Total charges 2025 (impôt + RRQ + FSS)</th>
            <th>Déjà retenu à la source</th>
            <th>Solde dû en avril {annee}</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td class="num" style="font-weight:600">{fmt_cad(d.get('total_charges'))}</td>
            <td class="num" style="color:#34d399">{fmt_cad(d.get('deja_paye'))}</td>
            <td class="num" style="color:{solde_color};font-weight:700;font-size:16px">{fmt_cad(d.get('solde_du'))}</td>
          </tr>
        </tbody>
      </table>
    </div>
    <p style="font-size:11px;color:#4b5563;margin-top:10px;font-style:italic">
      Source : cache fiscal 2025 (DÉPENSES 2025.xlsm)
    </p>
  </div>'''


# ── Bloc obligations ──────────────────────────────────────────────────────────

def _build_obligations_html(obligations: list) -> str:
    """Génère le HTML des obligations depuis ctx.obligations.

    Format ctx.obligations (depuis _read_obligations) :
      {'label', 'paid', 'total', 'pct', 'due_date', 'source', 'row'}
    """
    if not obligations:
        return ''

    cards = []
    for ob in obligations:
        label    = ob.get('label', '')
        paid     = ob.get('paid', 0) or 0
        total    = ob.get('total', 0) or 0
        pct      = ob.get('pct', 0) or 0
        due_date = ob.get('due_date')
        source   = ob.get('source', 'estimé') or 'estimé'

        bar_pct   = min(100.0, max(0.0, pct * 100))
        bar_color = '#34d399' if bar_pct >= 100 else ('#60a5fa' if bar_pct >= 50 else '#f87171')

        due_str = '—'
        if due_date:
            try:
                due_str = due_date.strftime('%Y-%m-%d') if hasattr(due_date, 'strftime') else str(due_date)[:10]
            except Exception:
                due_str = str(due_date)

        src_color = '#a3e635' if source == 'réel' else '#fb923c'
        src_badge = (f'<span style="font-size:9px;padding:2px 6px;border-radius:10px;'
                     f'background:#1e293b;color:{src_color};border:1px solid {src_color}40;'
                     f'margin-left:6px">{source}</span>')

        progress_bar = ''
        if total > 0:
            progress_bar = f'''
          <div style="margin:8px 0 4px">
            <div style="display:flex;justify-content:space-between;font-size:10px;color:var(--muted);margin-bottom:4px">
              <span>{fmt_cad(paid)} payé / {fmt_cad(total)} total</span>
              <span style="color:{bar_color};font-weight:700">{bar_pct:.1f} %</span>
            </div>
            <div style="height:6px;background:#1e293b;border-radius:3px;overflow:hidden">
              <div style="height:100%;width:{bar_pct:.1f}%;background:{bar_color};border-radius:3px;transition:width 0.4s"></div>
            </div>
          </div>'''

        cards.append(f'''
        <div class="card" style="padding:16px;margin-bottom:10px">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px">
            <div style="font-size:14px;font-weight:700;color:var(--text)">{label}{src_badge}</div>
            <span style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px;
                         padding:3px 8px;border-radius:20px;background:#1e293b;
                         color:#a78bfa;border:1px solid #4c1d95">actif</span>
          </div>
          {progress_bar}
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px">
            <div style="background:#0f172a;border-radius:10px;padding:10px">
              <div style="font-size:10px;color:var(--muted);margin-bottom:3px">Montant payé</div>
              <div style="font-size:16px;font-weight:700;color:#34d399">{fmt_cad(paid)}</div>
            </div>
            <div style="background:#0f172a;border-radius:10px;padding:10px">
              <div style="font-size:10px;color:var(--muted);margin-bottom:3px">Solde total cible</div>
              <div style="font-size:16px;font-weight:700;color:var(--text)">{fmt_cad(total) if total else '—'}</div>
            </div>
            <div style="background:#0f172a;border-radius:10px;padding:10px;grid-column:span 2">
              <div style="font-size:10px;color:var(--muted);margin-bottom:3px">Date limite</div>
              <div style="font-size:13px;font-weight:600;color:#60a5fa">{due_str}</div>
            </div>
          </div>
        </div>''')

    return '''
  <p class="section-title">Obligations financières</p>
''' + ''.join(cards)


# ── Bloc alertes ──────────────────────────────────────────────────────────────

def _build_alerts_html(alerts: list, score_label: str, score_emoji: str,
                       taxes_month: float = 0.0, taxes_label: str = '') -> str:
    """Génère le bloc HTML des alertes depuis ctx.alerts."""
    if not alerts:
        return ''

    sev_colors = {'HIGH': '#f87171', 'MEDIUM': '#fbbf24', 'LOW': '#60a5fa'}
    items = []
    for a in alerts:
        msg = a.message
        # Enrichir l'alerte taxes : mois courant + cumul annuel (pas de remplacement)
        if hasattr(a, 'type') and a.type in ('TAXES_ELEVEES', 'TAXES_MODEREES'):
            if taxes_month is not None and taxes_month > 0 and taxes_label:
                msg = (f'Prochaine remise ({taxes_label.lower()}) : '
                       f'{taxes_month:,.2f}\u202f$ — {a.message}')
        items.append(
            f'<div style="display:flex;align-items:flex-start;gap:10px;padding:10px 14px;'
            f'border-bottom:1px solid #1e293b">'
            f'<span style="font-size:16px;line-height:1.4">{a.icon}</span>'
            f'<div><div style="font-size:13px;color:var(--text);line-height:1.4">{msg}</div>'
            f'<div style="font-size:10px;color:{sev_colors.get(a.severity,"#94a3b8")};font-weight:600;'
            f'text-transform:uppercase;margin-top:3px">{a.severity}</div></div></div>'
        )
    alert_items = ''.join(items)

    return f'''
  <p class="section-title">Alertes de pilotage {score_emoji}</p>
  <div class="card">
    {alert_items}
  </div>'''


# ── Renderer principal ────────────────────────────────────────────────────────

def write_html_dashboard(ctx: 'CockpitContext', output_path: str) -> None:
    """Génère et écrit le dashboard HTML depuis un CockpitContext.

    Paramètres
    ----------
    ctx         : CockpitContext construit par build_cockpit_context()
    output_path : Chemin absolu du fichier HTML à écrire

    Ne lève pas d'exception — les erreurs de rendu dégradent gracieusement.
    """
    now   = datetime.now().strftime('%d/%m/%Y à %H:%M')
    annee = (ctx.reference_date.year if ctx.reference_date else ANNEE_DEFAUT)

    # ── Extraire les valeurs clés depuis ctx ──────────────────────────────────

    ac = ctx.alert_ctx  # AlertContext

    # AR (comptes à recevoir)
    a_recevoir     = float(ac.ar.total_st)     if ac and ac.ar       else 0.0
    a_recevoir_par_mois: dict = ctx.ar_par_mois or {}

    # Revenus encaissés depuis ctx.monthly_revenue
    revenus_ventes: dict = {}
    for entry in ctx.monthly_revenue:
        mois_int = entry.get('mois')
        if mois_int:
            revenus_ventes[mois_int] = {
                'freelance': entry.get('freelance', 0.0),
                'salaire':   entry.get('salaire', 0.0),
            }

    total_rev_free = sum(r.get('freelance', 0) for r in revenus_ventes.values())
    total_rev_sal  = sum(r.get('salaire', 0) for r in revenus_ventes.values())
    total_rev      = round(total_rev_free + total_rev_sal, 2)

    # Dépenses (agrégat annuel — pas de ventilation mensuelle en v2.1)
    total_dep = round(sum(ctx.expenses_by_category.values()), 2)

    total_profit = round(total_rev - total_dep, 2)

    # Taxes du mois courant — montant mensuel à remettre (pas le cumul annuel)
    taxes_mois_courant = ctx.taxes_current_month or 0.0
    taxes_mois_label   = ctx.taxes_current_label or ''

    # ── Estimation fiscale depuis TaxEstimate ─────────────────────────────────
    te             = ac.tax_estimate if ac else None
    solde          = 0.0
    total_charges  = 0.0
    total_impot    = 0.0
    cotisations    = 0.0   # RRQ + RQAP + FSS combinés

    if te:
        total_impot   = round(te.impot_federal + te.impot_qc, 2)
        cotisations   = round(te.cotisations_autonomes, 2)
        total_charges = round(total_impot + cotisations, 2)
        # solde = charges totales (sans déduction retenues — non dispo via ctx v2.1)
        solde         = total_charges

    solde_color = color_for(solde, positive_is_good=False) if solde else '#94a3b8'

    # ── Tableau revenus & dépenses par mois ──────────────────────────────────
    dep_by_m = ctx.expenses_by_month or {}
    rows_mois = []
    for i, mois in enumerate(MOIS_ORDER):
        mois_int   = annee * 100 + (i + 1)
        mois_idx   = i + 1
        rev        = revenus_ventes.get(mois_int, {})
        freelance  = rev.get('freelance', 0)
        salaire    = rev.get('salaire', 0)
        total_r    = freelance + salaire
        dep_m      = dep_by_m.get(mois_idx, 0)
        profit_m   = total_r - dep_m if (total_r > 0 or dep_m > 0) else 0
        has_data   = total_r > 0 or dep_m > 0
        empty_cls  = '' if has_data else ' class="empty-row"'
        rows_mois.append(f'''
        <tr{empty_cls}>
          <td class="mois-cell">{MOIS_ABBR.get(mois, mois[:3])}</td>
          <td class="num">{fmt_cad(freelance) if freelance else '—'}</td>
          <td class="num sal">{fmt_cad(salaire) if salaire else '—'}</td>
          <td class="num dep">{fmt_cad(dep_m) if dep_m else '—'}</td>
          <td class="num">{fmt_cad(profit_m) if has_data else '—'}</td>
        </tr>''')

    # ── Dépenses par catégorie ────────────────────────────────────────────────
    all_cats = dict(sorted(ctx.expenses_by_category.items(), key=lambda x: -x[1]))
    rows_cats = ''.join(
        f'<tr><td class="cat-cell">{cat}</td><td class="num">{fmt_cad(v)}</td></tr>'
        for cat, v in all_cats.items()
    )

    # ── Blocs HTML dynamiques ─────────────────────────────────────────────────
    obligations_html = _build_obligations_html(ctx.obligations)
    alerts_html      = _build_alerts_html(ctx.alerts, ctx.score_label, ctx.score_emoji,
                                            taxes_mois_courant, taxes_mois_label)
    annexe_2025_html = _build_annexe_2025(ctx.fiscal_2025_summary, annee)
    chart_html       = bar_chart_revenus(revenus_ventes, a_recevoir_par_mois)

    # ── Template HTML complet ─────────────────────────────────────────────────
    html = f'''<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"/>
  <meta name="apple-mobile-web-app-capable" content="yes"/>
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent"/>
  <title>Comptabilité {annee}</title>
  <style>
    :root {{
      --bg:      #0f172a;
      --card:    #1e293b;
      --border:  #334155;
      --text:    #f1f5f9;
      --muted:   #94a3b8;
      --accent:  #6366f1;
      --green:   #34d399;
      --red:     #f87171;
      --orange:  #fb923c;
      --yellow:  #fbbf24;
    }}
    * {{ box-sizing:border-box; margin:0; padding:0; }}
    body {{
      font-family: -apple-system, 'SF Pro Text', system-ui, sans-serif;
      background: var(--bg);
      color: var(--text);
      font-size: 15px;
      padding-bottom: env(safe-area-inset-bottom, 24px);
    }}
    .header {{
      background: linear-gradient(135deg, #1e1b4b 0%, #1e293b 100%);
      padding: 52px 20px 24px;
      border-bottom: 1px solid var(--border);
    }}
    .header h1 {{
      font-size: 22px; font-weight: 700;
      letter-spacing: -0.5px; margin-bottom: 4px;
    }}
    .header .sub {{ color: var(--muted); font-size: 13px; }}
    .header .update {{
      font-size: 11px; color: var(--muted); margin-top: 8px;
      display: flex; align-items: center; gap: 6px;
    }}
    .dot {{ width:6px;height:6px;border-radius:50%;background:var(--green);display:inline-block; }}
    .content {{ padding: 16px; max-width: 600px; margin: 0 auto; }}
    .section-title {{
      font-size: 11px; font-weight: 600; text-transform: uppercase;
      letter-spacing: 1px; color: var(--muted); margin: 24px 0 10px;
    }}
    .card {{
      background: var(--card);
      border-radius: 16px;
      border: 1px solid var(--border);
      overflow: hidden;
      margin-bottom: 12px;
    }}
    .kpi-row {{
      display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px;
    }}
    .kpi-row.three {{ grid-template-columns: 1fr 1fr 1fr; }}
    .kpi {{
      background: var(--card); border-radius: 14px;
      border: 1px solid var(--border); padding: 16px 14px;
    }}
    .kpi .label {{ font-size: 11px; color: var(--muted); margin-bottom: 6px; }}
    .kpi .value {{ font-size: 20px; font-weight: 700; letter-spacing: -0.5px; }}
    .kpi .value.sm {{ font-size: 16px; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th {{
      font-size: 10px; font-weight: 600; text-transform: uppercase;
      letter-spacing: 0.5px; color: var(--muted);
      padding: 10px 12px 8px; text-align: right;
      border-bottom: 1px solid var(--border);
    }}
    th:first-child {{ text-align: left; }}
    td {{ padding: 10px 12px; border-bottom: 1px solid #1e293b; font-size: 13px; }}
    td:first-child {{ color: var(--text); font-weight: 500; }}
    .num {{ text-align: right; font-variant-numeric: tabular-nums; color: var(--muted); }}
    .sal {{ color: #a5b4fc; }}
    .dep {{ color: var(--orange); }}
    .cat-cell {{ font-size: 12px; }}
    .mois-cell {{
      font-size: 12px; font-weight: 700;
      text-transform: uppercase; letter-spacing: 0.5px;
    }}
    tr:last-child td {{ border-bottom: none; }}
    .total-row td {{
      font-weight: 700;
      border-top: 1px solid var(--border);
      background: rgba(255,255,255,0.03);
    }}
    .empty-row td {{ opacity: 0.35; }}
    .solde-card {{
      background: linear-gradient(135deg, #1e1b4b 0%, #1e293b 100%);
      border-radius: 16px; border: 1px solid #3730a3;
      padding: 20px; margin-bottom: 12px; text-align: center;
    }}
    .solde-card .label {{ font-size: 12px; color: #a5b4fc; margin-bottom: 8px; }}
    .solde-card .amount {{ font-size: 36px; font-weight: 800; letter-spacing: -1px; }}
    .solde-card .sub {{ font-size: 11px; color: var(--muted); margin-top: 6px; }}
    .footer {{ text-align:center; color:var(--muted); font-size:11px; padding:32px 16px; }}
  </style>
</head>
<body>
  <div class="header">
    <h1>📊 Comptabilité {annee}</h1>
    <div class="sub">Félix Nault-Laberge</div>
    <div class="update">
      <span class="dot"></span>
      Mis à jour le {now}
    </div>
  </div>

  <div class="content">

    {alerts_html}

    <p class="section-title">Vue d'ensemble</p>
    <div class="kpi-row" style="grid-template-columns:repeat(4,1fr)">
      <div class="kpi">
        <div class="label">Revenus encaissés</div>
        <div class="value sm">{fmt_cad(total_rev)}</div>
      </div>
      <div class="kpi">
        <div class="label">À recevoir</div>
        <div class="value sm" style="color:var(--red)">{fmt_cad(a_recevoir)}</div>
      </div>
      <div class="kpi">
        <div class="label">Revenu total</div>
        <div class="value sm" style="color:#60a5fa">{fmt_cad(total_rev + a_recevoir)}</div>
      </div>
      <div class="kpi">
        <div class="label">Dépenses déductibles</div>
        <div class="value sm" style="color:var(--orange)">{fmt_cad(total_dep)}</div>
      </div>
    </div>
    <div class="kpi-row">
      <div class="kpi">
        <div class="label">Profit brut (encaissé)</div>
        <div class="value sm" style="color:{color_for(total_profit)}">{fmt_cad(total_profit)}</div>
      </div>
      <div class="kpi">
        <div class="label">Taxes à remettre{f' — {taxes_mois_label.lower()}' if taxes_mois_label else ''}</div>
        <div class="value sm" style="color:var(--yellow)">{fmt_cad(taxes_mois_courant)}</div>
      </div>
    </div>

    <p class="section-title">Solde fiscal estimé</p>
    <div class="solde-card">
      <div class="label">Charges fiscales estimées en {annee + 1} (impôt + cotisations autonomes)</div>
      <div class="amount" style="color:{solde_color}">{fmt_cad(solde)}</div>
      <div class="sub" style="margin-top:10px;line-height:1.8">
        Impôt fédéral + provincial : {fmt_cad(total_impot)}<br>
        Cotisations autonomes (RRQ + RQAP + FSS) : {fmt_cad(cotisations)}<br>
        <span style="border-top:1px solid #4b5563;display:inline-block;padding-top:4px;margin-top:2px">
          Total charges estimées : {fmt_cad(total_charges)}
        </span>
      </div>
      <div class="sub" style="margin-top:6px;font-style:italic;color:#6b7280">
        Basé sur les revenus réels à ce jour — aucune projection annuelle
      </div>
    </div>

    <p class="section-title">Revenus par mois</p>
    <div class="card" style="padding:16px 12px 8px">
      {chart_html}
    </div>
    <div class="card">
      <table>
        <thead>
          <tr>
            <th></th>
            <th>Freelance</th>
            <th style="color:#a5b4fc">Salaire</th>
            <th style="color:var(--orange)">Dépenses</th>
            <th>Profit</th>
          </tr>
        </thead>
        <tbody>
          {''.join(rows_mois)}
          <tr class="total-row">
            <td>TOTAL</td>
            <td class="num">{fmt_cad(total_rev_free)}</td>
            <td class="num sal">{fmt_cad(total_rev_sal)}</td>
            <td class="num dep">{fmt_cad(total_dep)}</td>
            <td class="num" style="color:{color_for(total_profit)};font-weight:700">{fmt_cad(total_profit)}</td>
          </tr>
        </tbody>
      </table>
    </div>

    <p class="section-title">Dépenses par catégorie (cumulatif)</p>
    <div class="card">
      <table>
        <thead>
          <tr><th>Catégorie</th><th>Total</th></tr>
        </thead>
        <tbody>
          {rows_cats}
          <tr class="total-row">
            <td>TOTAL</td>
            <td class="num" style="color:var(--orange);font-weight:700">{fmt_cad(total_dep)}</td>
          </tr>
        </tbody>
      </table>
    </div>

    {obligations_html}

    {annexe_2025_html}

  </div>

  <div class="footer">Généré automatiquement · Données privées · Source : CockpitContext</div>
</body>
</html>'''

    with open(output_path, 'w', encoding='utf-8') as fh:
        fh.write(html)
