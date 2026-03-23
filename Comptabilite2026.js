// Comptabilité 2026 — Widget Scriptable
// Affiche les KPIs du dashboard comptable
// URL: https://flixnl.github.io/comptabilite2026/data.json

const DATA_URL = "https://flixnl.github.io/comptabilite2026/data.json"
const SITE_URL = "https://flixnl.github.io/comptabilite2026/"

// Couleurs
const BG      = new Color("#0f172a")
const CARD    = new Color("#1e293b")
const TEXT    = new Color("#f1f5f9")
const MUTED   = new Color("#94a3b8")
const GREEN   = new Color("#34d399")
const RED     = new Color("#f87171")
const ORANGE  = new Color("#fb923c")
const YELLOW  = new Color("#fbbf24")
const ACCENT  = new Color("#a5b4fc")

async function loadData() {
  try {
    let req = new Request(DATA_URL)
    req.timeoutInterval = 10
    return await req.loadJSON()
  } catch (e) {
    return null
  }
}

function fmt(n) {
  if (n == null) return "—"
  let s = Math.abs(n).toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, " ")
  return (n < 0 ? "-" : "") + s + " $"
}

function addKPI(stack, label, value, color) {
  let col = stack.addStack()
  col.layoutVertically()
  col.setPadding(10, 12, 10, 12)
  col.backgroundColor = CARD
  col.cornerRadius = 12

  let lbl = col.addText(label.toUpperCase())
  lbl.font = Font.semiboldSystemFont(9)
  lbl.textColor = MUTED
  lbl.lineLimit = 1

  col.addSpacer(4)

  let val = col.addText(fmt(value))
  val.font = Font.boldSystemFont(16)
  val.textColor = color
  val.lineLimit = 1
  val.minimumScaleFactor = 0.7
}

function addMiniBar(stack, moisData) {
  let maxVal = Math.max(...moisData.map(m => m.freelance + m.salaire), 1)

  let barRow = stack.addStack()
  barRow.layoutHorizontally()
  barRow.spacing = 3

  for (let m of moisData) {
    let col = barRow.addStack()
    col.layoutVertically()
    col.size = new Size(0, 60)

    let total = m.freelance + m.salaire
    let pct = total / maxVal

    col.addSpacer()

    // Barre freelance
    let barF = col.addStack()
    barF.backgroundColor = new Color("#6366f1")
    barF.cornerRadius = 2
    barF.size = new Size(18, Math.max(2, pct * 40))

    // Barre salaire
    if (m.salaire > 0) {
      let barS = col.addStack()
      barS.backgroundColor = YELLOW
      barS.cornerRadius = 2
      barS.size = new Size(18, Math.max(2, (m.salaire / maxVal) * 40))
    }

    col.addSpacer(2)

    let lbl = col.addText(m.mois)
    lbl.font = Font.boldSystemFont(8)
    lbl.textColor = MUTED
    lbl.centerAlignText()
  }
}

async function createWidget(data) {
  let w = new ListWidget()
  w.backgroundColor = BG
  w.setPadding(14, 14, 14, 14)
  w.url = SITE_URL

  if (!data) {
    let err = w.addText("Données indisponibles")
    err.font = Font.mediumSystemFont(14)
    err.textColor = RED
    return w
  }

  // Titre
  let header = w.addStack()
  header.layoutHorizontally()
  header.centerAlignContent()
  let title = header.addText("Comptabilité 2026")
  title.font = Font.boldSystemFont(14)
  title.textColor = TEXT
  header.addSpacer()
  let upd = header.addText(data.updated)
  upd.font = Font.mediumSystemFont(9)
  upd.textColor = MUTED

  w.addSpacer(10)

  // KPI row 1
  let row1 = w.addStack()
  row1.layoutHorizontally()
  row1.spacing = 8
  addKPI(row1, "Revenus", data.revenus, TEXT)
  addKPI(row1, "Dépenses", data.depenses, ORANGE)
  addKPI(row1, "Profit", data.profit, data.profit >= 0 ? GREEN : RED)

  w.addSpacer(8)

  // KPI row 2
  let row2 = w.addStack()
  row2.layoutHorizontally()
  row2.spacing = 8
  addKPI(row2, "Taxes dues", data.taxes_a_remettre, YELLOW)
  addKPI(row2, "Impôt estimé", data.solde_fiscal, data.solde_fiscal > 0 ? RED : GREEN)

  w.addSpacer(8)

  // Mini barres par mois
  if (data.mois && data.mois.length > 0) {
    addMiniBar(w, data.mois)
  }

  return w
}

let data = await loadData()
let widget = await createWidget(data)

if (config.runsInWidget) {
  Script.setWidget(widget)
} else {
  await widget.presentMedium()
}

Script.complete()
