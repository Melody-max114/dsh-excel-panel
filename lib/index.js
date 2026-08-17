import { appendFile, mkdir } from 'node:fs/promises'
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
import { homedir } from 'node:os'
import { dirname, isAbsolute, join, relative, resolve, sep } from 'node:path'

const execFileAsync = promisify(execFile)

const name = 'dsh-excel-panel'
const inject = ['webServer']

function json(res, status, obj) {
  res.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' })
  res.end(JSON.stringify(obj))
}

async function logSave(body) {
  try {
    await appendFile(join(homedir(), '.dsh', 'excel-panel-save.log'), JSON.stringify(body) + '\n', 'utf8')
  } catch {}
}

function readBody(req) {
  return new Promise((resolvePromise) => {
    let raw = ''
    req.on('data', (chunk) => { raw += chunk })
    req.on('end', () => {
      try { resolvePromise(JSON.parse(raw || '{}')) } catch { resolvePromise({}) }
    })
    req.on('error', () => resolvePromise({}))
  })
}

function inside(root, target) {
  const rel = relative(resolve(root), resolve(target))
  return rel === '' || (rel !== '..' && !rel.startsWith('..' + sep) && !isAbsolute(rel))
}

function safeResolve(root, rel) {
  const base = resolve(root)
  const target = resolve(base, rel || '')
  if (!inside(base, target)) return null
  return target
}

async function loadExcelJs() {
  const mod = await import('exceljs')
  return mod.default ?? mod
}


function colToIndex(colStr) {
  let n = 0
  const str = String(colStr).toUpperCase()
  for (let i = 0; i < str.length; i += 1) n = n * 26 + (str.charCodeAt(i) - 64)
  return n - 1
}

async function readXlsx(abs) {
  const ExcelJS = await loadExcelJs()
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(abs)
  const sheets = wb.worksheets.map((ws) => {
    const rows = []
    for (let r = 1; r <= ws.actualRowCount; r += 1) {
      const row = []
      for (let c = 1; c <= ws.actualColumnCount; c += 1) {
        const cell = ws.getCell(r, c)
        const fill = cell.fill
        const bg = fill && fill.type === 'pattern' && fill.pattern === 'solid' && fill.fgColor ? fill.fgColor.argb : undefined
        const font = cell.font || {}
        const fontColor = font.color && (font.color.argb || font.color.rgb) ? (font.color.argb || font.color.rgb) : undefined
        const align = cell.alignment && cell.alignment.horizontal ? cell.alignment.horizontal : undefined
        row.push({
          v: cell.text,
          f: cell.formula !== undefined ? cell.formula : undefined,
          bg: typeof bg === 'string' && bg !== '' ? bg : undefined,
          align: align,
          font: {
            bold: !!font.bold,
            italic: !!font.italic,
            underline: !!font.underline,
            size: typeof font.size === 'number' ? font.size : undefined,
            color: typeof fontColor === 'string' && fontColor !== '' ? fontColor : undefined
          }
        })
      }
      rows.push(row)
    }
    const mergeRanges = Array.isArray(ws.model?.merges) ? ws.model.merges : []
    const merges = mergeRanges.map((range) => {
      const m = String(range).match(/^([A-Z]+)(d+):([A-Z]+)(d+)$/)
      if (!m) return null
      return { r1: Number(m[2]) - 1, c1: colToIndex(m[1]), r2: Number(m[4]) - 1, c2: colToIndex(m[3]) }
    }).filter(Boolean)
    const colWidths = {}
    for (let c = 1; c <= ws.columnCount; c += 1) {
      const w = ws.getColumn(c).width
      if (typeof w === 'number' && Number.isFinite(w) && w > 0) colWidths[c - 1] = w
    }
    const rowHeights = {}
    for (let r = 1; r <= ws.actualRowCount; r += 1) {
      const h = ws.getRow(r).height
      if (typeof h === 'number' && Number.isFinite(h) && h > 0) rowHeights[r - 1] = h
    }
    return { name: ws.name, rows, merges, colWidths, rowHeights }
  })
  const activeSheet = sheets[0]?.name ?? 'Sheet1'
  return { kind: 'xlsx', sheets, activeSheet, sheetName: activeSheet }
}

async function writeXlsx(abs, payload) {
  const ExcelJS = await loadExcelJs()
  const wb = new ExcelJS.Workbook()
  const sheets = Array.isArray(payload?.sheets) ? payload.sheets : []
  for (const sheet of sheets) {
    if (!sheet || typeof sheet !== 'object') continue
    const name = typeof sheet.name === 'string' && sheet.name !== '' ? sheet.name : 'Sheet'
    const ws = wb.addWorksheet(name)
    const rows = Array.isArray(sheet.rows) ? sheet.rows : []
    for (let r = 0; r < rows.length; r += 1) {
      const row = rows[r] || []
      for (let c = 0; c < row.length; c += 1) {
        const src = row[c] || {}
        const cell = ws.getRow(r + 1).getCell(c + 1)
        if (src.f) {
          cell.value = { formula: String(src.f).replace(/^=/, ''), result: src.v }
        } else {
          cell.value = src.v
        }
        const f = src.font || {}
        const font = { size: typeof f.size === 'number' ? f.size : 11, name: f.name || 'Calibri' }
        if (f.bold) font.bold = true
        if (f.italic) font.italic = true
        if (f.underline) font.underline = true
        if (f.color) font.color = { argb: String(f.color) }
        cell.font = font
        if (src.bg) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: String(src.bg) } }
        } else {
          cell.fill = {}
        }
        if (src.align) {
          cell.alignment = { horizontal: src.align, vertical: 'middle' }
        }
      }
    }
    if (sheet.colWidths && typeof sheet.colWidths === 'object') {
      for (const [key, value] of Object.entries(sheet.colWidths)) {
        const c = Number(key) + 1
        if (Number.isInteger(c) && c >= 1 && Number.isFinite(Number(value)) && Number(value) > 0) {
          ws.getColumn(c).width = Number(value)
        }
      }
    }
    if (sheet.rowHeights && typeof sheet.rowHeights === 'object') {
      for (const [key, value] of Object.entries(sheet.rowHeights)) {
        const r = Number(key) + 1
        if (Number.isInteger(r) && r >= 1 && Number.isFinite(Number(value)) && Number(value) > 0) {
          ws.getRow(r).height = Number(value)
        }
      }
    }
    if (Array.isArray(sheet.merges)) {
      for (const m of sheet.merges) {
        if (!m || typeof m.r1 !== 'number') continue
        try { ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1) } catch {}
      }
    }
  }
  await wb.xlsx.writeFile(abs)
}
async function unlockFile(abs) {
  const escaped = String(abs).replace(/'/g, "''")
  const script = `$p = @(Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*${escaped}*' }); $n = $p.Count; foreach ($proc in $p) { Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue }; Write-Output $n`
  try {
    const { stdout } = await execFileAsync('powershell', ['-NoProfile', '-Command', script], { windowsHide: true })
    return Number(String(stdout).trim()) || 0
  } catch (error) {
    return 0
  }
}

function apply(ctx) {
  const webServer = ctx.get('webServer')
  if (!webServer) return
  webServer.register({
    kind: 'prefix',
    path: '/excel-panel',
    handler: async (req, res) => {
      if (req.method !== 'POST') {
        json(res, 405, { ok: false, error: 'method not allowed' })
        return
      }
      const body = await readBody(req)
      const root = typeof body.root === 'string' ? body.root : ''
      const rel = typeof body.path === 'string' ? body.path : ''
      if (!root || !rel) {
        json(res, 400, { ok: false, error: 'root and path are required' })
        return
      }
      const abs = safeResolve(root, rel)
      if (!abs) {
        json(res, 403, { ok: false, error: 'path outside root' })
        return
      }
      const url = new URL(req.url ?? '/', 'http://x').pathname
      try {
        if (url === '/excel-panel/read') {
          const data = await readXlsx(abs)
          json(res, 200, { ok: true, value: data })
          return
        }
        if (url === '/excel-panel/write') {
          await logSave(body)
          await writeXlsx(abs, body.payload)
          json(res, 200, { ok: true, value: { saved: true } })
          return
        }
        if (url === '/excel-panel/log') {
          try {
            await appendFile(join(homedir(), '.dsh', 'excel-panel-operation.log'), JSON.stringify({ time: new Date().toISOString(), root, path: rel, ...(typeof body.op === 'object' ? body.op : { op: body.op }) }) + String.fromCharCode(10), 'utf8')
          } catch {}
          json(res, 200, { ok: true, value: { logged: true } })
          return
        }
        if (url === '/excel-panel/unlock') {
          const killed = await unlockFile(abs)
          json(res, 200, { ok: true, value: { killed } })
          return
        }
        json(res, 404, { ok: false, error: 'not found' })
      } catch (error) {
        try {
          await appendFile(join(homedir(), '.dsh', 'excel-panel-error.log'), JSON.stringify({ time: new Date().toISOString(), root, rel, error: String(error && error.stack || error) }) + '\n', 'utf8')
        } catch {}
        json(res, 500, { ok: false, error: String(error && error.message || error) })
      }
    }
  })
}

export { apply, inject, name, readXlsx, writeXlsx }
