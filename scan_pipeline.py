#!/usr/bin/env python3
"""
dashboard/scan_pipeline.py — Orchestrateur du scan complet (STEP 3.2).

Isole la logique d'orchestration qui était dans menubar_compta._do_scan(),
permettant les tests unitaires et les futurs refactors.

Étapes
──────
1. scan_import.py           → importe les PDFs → Journal
2. sync_ventes_to_depenses.py → sync VENTES → DÉPENSES
3a. write_html_dashboard()  → génère le HTML depuis CockpitContext frais
3b. write_dashboard_v2()    → écrit Dashboard V2 dans DÉPENSES.xlsm
3c. V3 health check         → validation per-obligation cash_ledger
3d. system_snapshot          → capture JSON de l'état complet (audit)
4. git_commit_and_push()    → pousse les fichiers HTML (silent, non-fatal)

Usage
─────
    from dashboard.scan_pipeline import run
    result = run(outils_dir)
    # result = {'ok': bool, 'errors': list[str], 'output': str, 'ctx': CockpitContext|None}

Depuis menubar_compta._do_scan() :
    result = run(SCRIPT_DIR, ctx=self._cockpit)
    # ctx passé pour usage futur ; la pipeline reconstruit toujours un ctx frais.
"""
from __future__ import annotations

import glob as _glob
import os
import subprocess
from datetime import datetime
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from dashboard.context import CockpitContext


# ─── Résolution des chemins source ────────────────────────────────────────────

def _resolve_dirs(outils_dir: str) -> tuple[str, str]:
    """Retourne (ventes_path, depenses_path) depuis outils_dir.

    Retourne les chemins vers les FICHIERS xlsx/xlsm (pas les répertoires)
    pour que build_cockpit_context() n'ait pas à deviner via glob.

    Priorité :
      1. core.paths.resolve_main_paths(2026) → fichiers exacts
      2. Fallback glob conventionnel (Comptabilite/2026/)
    """
    try:
        import sys as _sys
        if outils_dir not in _sys.path:
            _sys.path.insert(0, outils_dir)
        from core.paths import resolve_main_paths
        paths = resolve_main_paths(2026)
        return paths.ventes, paths.depenses
    except Exception:
        pass

    # Fallback : outils_dir = .../Comptabilite/2026/_OUTILS
    base_2026 = os.path.dirname(outils_dir)

    # VENTES_AAAA.xlsx à la racine de 2026/
    v_hits = sorted(_glob.glob(os.path.join(base_2026, 'VENTES_*.xlsx')))
    ventes = v_hits[0] if v_hits else os.path.join(base_2026, 'VENTES')

    # DÉPENSES AAAA.xlsm dans DÉPENSES/
    dep_dir = os.path.join(base_2026, 'DÉPENSES')
    if os.path.isdir(dep_dir):
        d_hits = sorted(_glob.glob(os.path.join(dep_dir, '*PENSES*.xlsm')))
        depenses = d_hits[-1] if d_hits else dep_dir
    else:
        d_hits = _glob.glob(os.path.join(base_2026, '*PENSES*'))
        depenses = d_hits[0] if d_hits else ''

    return ventes, depenses


# ─── run() ────────────────────────────────────────────────────────────────────

def run(
    outils_dir: str,
    ctx: Optional['CockpitContext'] = None,   # accepté pour usage futur
) -> dict:
    """Exécute les 4 étapes du scan complet.

    Paramètres
    ----------
    outils_dir : Chemin vers _OUTILS/ (contient scan_import.py, .venv/, …)
    ctx        : CockpitContext en cache (réservé — la pipeline reconstruit
                 toujours un ctx frais après le sync)

    Retourne
    --------
    {
      'ok':     bool,              # True si 0 erreur
      'errors': list[str],         # messages d'erreur courts (affichage menubar)
      'output': str,               # stdout+stderr de scan_import (pour notif)
      'ctx':    CockpitContext|None  # contexte frais post-sync (peut être None)
    }
    """
    errors:  list[str] = []
    output:  str       = ''
    new_ctx: Optional['CockpitContext'] = None

    log_path = os.path.join(outils_dir, 'scan_import.log')
    venv_py  = os.path.join(outils_dir, '.venv', 'bin', 'python3')

    def _log(msg: str) -> None:
        try:
            with open(log_path, 'a') as _lf:
                _lf.write(f'[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n')
        except Exception:
            pass

    # ── Étape 1 : scan PDFs → Journal ─────────────────────────────────────────
    scan_script = os.path.join(outils_dir, 'scan_import.py')
    try:
        r1     = subprocess.run(
            [venv_py, scan_script],
            capture_output=True, text=True, timeout=300,
        )
        output = r1.stdout + r1.stderr
        if r1.returncode != 0:
            errors.append(
                f'Scan: {output.strip().split(chr(10))[-1][:60]}')
            _log(f'⚠️ Scan erreur:\n{output}')
        else:
            _log('✅ Scan OK')
    except subprocess.TimeoutExpired:
        errors.append('Scan: timeout (>5min)')
        _log('⚠️ Scan timeout')
    except Exception as _e:
        errors.append(f'Scan: {_e}')
        _log(f'⚠️ Scan exception: {_e}')

    # ── Étape 2 : sync VENTES → DÉPENSES ──────────────────────────────────────
    sync_script = os.path.join(outils_dir, 'sync_ventes_to_depenses.py')
    try:
        r2 = subprocess.run(
            [venv_py, sync_script],
            capture_output=True, text=True, timeout=60,
        )
        if r2.returncode != 0:
            errors.append(
                f'Sync: {(r2.stderr or r2.stdout or "").strip().split(chr(10))[-1][:60]}')
            _log(f'⚠️ Sync erreur:\n{r2.stderr or r2.stdout}')
        else:
            _log('✅ Sync OK')
    except Exception as _e:
        errors.append(f'Sync: {_e}')
        _log(f'⚠️ Sync exception: {_e}')

    # ── Étape 3 : dashboard HTML ───────────────────────────────────────────────
    try:
        import sys as _sys
        import importlib as _importlib
        if outils_dir not in _sys.path:
            _sys.path.insert(0, outils_dir)
        # Hot-reload tous les modules modifiables (Pipeline V3)
        import dashboard.context     as _ctx_mod
        import dashboard.html_writer as _hw_mod
        import analytics.alerts      as _alerts_mod
        _importlib.reload(_ctx_mod)           # hot-reload si .py modifié
        _importlib.reload(_hw_mod)            # hot-reload si .py modifié
        _importlib.reload(_alerts_mod)        # hot-reload alertes V3
        # Reload cash_ledger si disponible (Pipeline V3)
        try:
            import fiscal.cash_ledger as _cl_mod
            _importlib.reload(_cl_mod)
        except ImportError:
            pass
        build_cockpit_context  = _ctx_mod.build_cockpit_context
        write_html_dashboard   = _hw_mod.write_html_dashboard
        import pathlib
        import shutil

        ventes_dir, depenses_dir = _resolve_dirs(outils_dir)
        new_ctx = build_cockpit_context(ventes_dir, depenses_dir)

        html_main    = os.path.join(outils_dir, 'Tableau de bord comptable.html')
        html_ghpages = os.path.join(outils_dir, 'dashboard', 'index.html')

        write_html_dashboard(new_ctx, html_main)
        pathlib.Path(html_ghpages).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(html_main, html_ghpages)
        _log('✅ Dashboard OK')
    except Exception as _e:
        errors.append(f'Dashboard: {_e}')
        _log(f'⚠️ Dashboard erreur: {_e}')
        new_ctx = None

    # ── Étape 3b : DASHBOARD V2 dans DÉPENSES.xlsm ──────────────────────────
    if new_ctx is not None:
        try:
            import importlib
            import dashboard.write_dashboard_v2 as _wd2_mod
            importlib.reload(_wd2_mod)           # hot-reload si .py modifié
            write_dashboard_v2 = _wd2_mod.write_dashboard_v2
            _, dep_file = _resolve_dirs(outils_dir)
            write_dashboard_v2(new_ctx, dep_file)
            _log('✅ Dashboard V2 OK')

            # Nav bars sur les feuilles statiques
            try:
                from dashboard.apply_nav_bar import apply_nav_bars
                _nav_done = apply_nav_bars(dep_file, verbose=False)
                _log(f'✅ Nav bars OK ({len(_nav_done)} feuilles)')
            except Exception as _nav_e:
                _log(f'ℹ️ Nav bars non-fatal: {_nav_e}')

        except Exception as _e:
            _log(f'⚠️ Dashboard V2 erreur: {_e}')
            # Non-fatal — l'ancien TABLEAU DE BORD reste fonctionnel

    # ── Étape 3c : Pipeline V3 health check (cash_ledger per-obligation) ─────
    if new_ctx is not None:
        try:
            from fiscal.cash_ledger import load_ledger as _load_cl
            cl_entries = [e for e in _load_cl() if e.is_active]
            if not cl_entries:
                _log('ℹ️ V3 Health: cash_ledger vide (aucune entrée)')
            else:
                cl_total = sum(e.amount for e in cl_entries)
                # Regrouper par obligation_id
                cl_by_obl: dict = {}
                for _e in cl_entries:
                    _oid = getattr(_e, 'obligation_id', '') or '_sans'
                    cl_by_obl[_oid] = cl_by_obl.get(_oid, 0.0) + _e.amount
                obl_list = new_ctx.obligations or []
                # Comparer uniquement les obligations couvertes par le ledger
                mismatches = []
                for _o in obl_list:
                    _oid = _o.get('id', '')
                    if _oid not in cl_by_obl:
                        continue  # obligation hors ledger — normal
                    _obl_paid = _o.get('paid', 0)
                    _delta = abs(cl_by_obl[_oid] - _obl_paid)
                    if _delta > 100:
                        mismatches.append(f'{_oid}: ledger={cl_by_obl[_oid]:.0f}$ vs obl={_obl_paid:.0f}$')
                covered = len(set(cl_by_obl.keys()) & {_o.get('id','') for _o in obl_list})
                uncovered = len(obl_list) - covered
                if mismatches:
                    _log(f'⚠️ V3 Health: {len(mismatches)} divergence(s) per-obligation : '
                         f'{"; ".join(mismatches[:3])} | {uncovered} obligation(s) hors ledger')
                else:
                    _log(f'✅ V3 Health: {covered} obligation(s) OK — '
                         f'ledger={cl_total:.0f}$ | {uncovered} hors ledger (normal)')
        except ImportError:
            pass  # cash_ledger pas encore installé
        except Exception as _e:
            _log(f'ℹ️ V3 Health check non-fatal: {_e}')

    # ── Étape 3d : System snapshot (audit hebdomadaire) ─────────────────────
    try:
        from dashboard.system_snapshot import build_snapshot, write_snapshot
        _snap = build_snapshot(
            ctx=new_ctx,
            outils_dir=outils_dir,
            scan_ok=(len(errors) == 0),
        )
        _snap_path = write_snapshot(_snap, outils_dir)
        _log(f'✅ Snapshot OK → {os.path.basename(_snap_path)}')
    except Exception as _e:
        _log(f'ℹ️ Snapshot non-fatal: {_e}')

    # ── Étape 4 : git push (silent, non-fatal) ────────────────────────────────
    try:
        from dashboard.git_utils import git_commit_and_push
        html_main    = os.path.join(outils_dir, 'Tableau de bord comptable.html')
        html_ghpages = os.path.join(outils_dir, 'dashboard', 'index.html')
        # html_ghpages uniquement : html_main est dans _OUTILS/ (hors du repo dashboard/)
        git_ok, git_msg = git_commit_and_push(
            files    = [html_ghpages],
            message  = f'Scan auto — dashboard mis à jour ({datetime.now():%Y-%m-%d %H:%M})',
            log_path = log_path,
        )
        if git_ok:
            _log('✅ Git push OK')
        else:
            _log(f'ℹ️ Git push ignoré : {git_msg}')
    except Exception as _e:
        _log(f'ℹ️ Git push non exécuté : {_e}')

    return {
        'ok':     len(errors) == 0,
        'errors': errors,
        'output': output,
        'ctx':    new_ctx,
    }
