import { appendFile, mkdir } from 'node:fs/promises'
import { homedir } from 'node:os'
import { dirname, isAbsolute, join, relative, resolve, sep } from 'node:path'

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
    return { name: ws.name, rows, merges }
  })
  const activeSheet = sheets[0]?.name ?? 'Sheet1'
  return { kind: 'xlsx', sheets, activeSheet, sheetName: activeSheet }
}

async function writeXlsx(abs, payload) {
  const ExcelJS = await loadExcelJs()
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(abs)
  const sheets = Array.isArray(payload?.sheets) ? payload.sheets : []
  for (const sheet of sheets) {
    if (!sheet || typeof sheet !== 'object') continue
    const name = typeof sheet.name === 'string' && sheet.name !== '' ? sheet.name : null
    const rows = Array.isArray(sheet.rows) ? sheet.rows : []
    let ws = name ? wb.getWorksheet(name) : undefined
    if (ws === undefined && name) ws = wb.addWorksheet(name)
    if (ws === undefined) ws = wb.worksheets[0]
    if (ws === undefined) continue
    const maxRow = Math.max(ws.actualRowCount, rows.length)
    for (let r = 1; r <= maxRow; r += 1) {
      const incoming = rows[r - 1] !== undefined ? rows[r - 1] : []
      for (let c = 1; c <= Math.max(ws.actualColumnCount, incoming.length); c += 1) {
        const cell = ws.getCell(r, c)
        const next = incoming[c - 1]
        const nextText = next && next.v !== undefined && next.v !== null ? String(next.v) : ''
        const nextFormula = next && typeof next.f === 'string' && next.f.trim() !== '' ? next.f.trim().replace(/^=/, '') : ''
        const nextFont = next && typeof next.font === 'object' && next.font ? next.font : null
        const nextBg = next && typeof next.bg === 'string' && next.bg !== '' ? next.bg : null
        const nextAlign = next && typeof next.align === 'string' && ['left', 'center', 'right'].includes(next.align) ? next.align : null
        if (nextFont) {
          const font = Object.assign({}, cell.font || {})
          if (nextFont.bold !== undefined) font.bold = !!nextFont.bold
          if (nextFont.italic !== undefined) font.italic = !!nextFont.italic
          if (nextFont.underline !== undefined) font.underline = !!nextFont.underline
          if (nextFont.size !== undefined) font.size = Number(nextFont.size)
          if (nextFont.color !== undefined && nextFont.color !== '') font.color = { argb: String(nextFont.color) }
          cell.font = font
        }
        if (nextBg) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: String(nextBg) } }
        }
        if (nextAlign) {
          cell.alignment = { horizontal: nextAlign, vertical: 'middle' }
        }
        if (nextFormula !== '') {
          cell.value = { formula: nextFormula, result: nextText === '' ? undefined : nextText }
          continue
        }
        if (cell.text === nextText) continue
        if (nextText === '') { cell.value = null; continue }
        const num = Number(nextText)
        cell.value = Number.isFinite(num) && String(num) === nextText ? num : nextText
      }
    }
  }
  for (const sheet of sheets) {
    if (!sheet || typeof sheet !== 'object') continue
    const name = typeof sheet.name === 'string' && sheet.name !== '' ? sheet.name : null
    const ws = name ? wb.getWorksheet(name) : undefined
    if (!ws) continue
    const existing = Array.isArray(ws.model?.merges) ? ws.model.merges.slice() : []
    for (const range of existing) {
      try { ws.unMergeCells(range) } catch {}
    }
    const merges = Array.isArray(sheet.merges) ? sheet.merges : []
    for (const m of merges) {
      if (!m || typeof m.r1 !== 'number') continue
      try {
        ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1)
      } catch {}
    }
  }
  await wb.xlsx.writeFile(abs)
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
        json(res, 404, { ok: false, error: 'not found' })
      } catch (error) {
        json(res, 500, { ok: false, error: String(error && error.message || error) })
      }
    }
  })
}

export { apply, inject, name }
