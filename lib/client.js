window.__ModuleLoader__.load({
  id: 'dsh-excel-panel',
  factory: (require) => {
    var module = { exports: {} }
    var exports = module.exports
    Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' })
    var React = require('react')
    var h = React.createElement

    var inject = ['betterSidebar']

    function post(path, body) {
      return fetch(path, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body || {})
      }).then(function (r) {
        return r.json().catch(function () {
          return { ok: false, error: 'invalid JSON response', status: r.status }
        }).then(function (body) {
          if (body && typeof body === 'object') body.status = r.status
          return body
        })
      })
    }

    function officeTarget(scope, path) {
      var root = scope && scope.cwd ? scope.cwd : ''
      var rel = path
      if (!root) {
        var idx = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
        if (idx > -1) {
          root = path.slice(0, idx)
          rel = path.slice(idx + 1)
        }
      }
      return { root: root, path: rel }
    }

    function readExcel(scope, path) {
      var target = officeTarget(scope, path)
      return post('/excel-panel/read', { root: target.root, path: target.path, kind: 'xlsx' }).then(function (res) {
        if (!res.ok) {
          var detail = res && res.error
          var message = typeof detail === 'string' && detail ? detail : (detail && detail.message) || ''
          if (!message && detail && typeof detail === 'object') {
            try { message = JSON.stringify(detail) } catch (_) {}
          }
          throw new Error('read failed' + (message ? ': ' + message : '') + (res && res.status ? ' (HTTP ' + res.status + ')' : ''))
        }
        return res.value
      })
    }

    function cellDisplay(v) {
      if (v === undefined || v === null) return ''
      return String(v)
    }

    function colName(n) {
      var s = ''
      while (n > 0) {
        var m = (n - 1) % 26
        s = String.fromCharCode(65 + m) + s
        n = Math.floor((n - 1) / 26)
      }
      return s
    }

    function shiftFormula(formula, dRow, dCol) {
      return String(formula).replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g, function (match, dollarCol, colStr, dollarRow, rowStr) {
        var col = 0
        for (var i = 0; i < colStr.length; i++) col = col * 26 + (colStr.toUpperCase().charCodeAt(i) - 64)
        var row = Number(rowStr)
        var newCol = dollarCol ? col : col + dCol
        var newRow = dollarRow ? row : row + dRow
        if (newCol < 1 || newRow < 1) return '#REF!'
        return (dollarCol || '') + colName(newCol) + (dollarRow || '') + newRow
      })
    }

    function headerStyle() {
      return { border: '1px solid #ddd', padding: '2px 6px', fontWeight: 600, fontSize: 12, background: '#f2f3f5', textAlign: 'center' }
    }

    function ensureExcelStyles() {
      if (typeof document === 'undefined') return
      var id = 'dsh-excel-panel-selection-css'
      if (document.getElementById(id)) return
      var style = document.createElement('style')
      style.id = id
      style.textContent = '@keyframes dshExcelSelectIn{from{box-shadow:inset 0 0 0 0 rgba(22,93,255,0)}to{box-shadow:inset 0 0 0 2px #165dff}}'
      document.head.appendChild(style)
    }



    function cellAddress(r, c) {
      return colName(c + 1) + (r + 1)
    }

    function tokenizeFormula(s) {
      var str = String(s == null ? '' : s).replace(/^=/, '').trim()
      if (!str) return []
      var re = /("(?:[^"]|"")*"|'(?:[^']|'')*'|\$?[A-Za-z]{1,3}\$?\d+|[A-Za-z_][A-Za-z0-9_.]*|\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|<=|>=|<>|!=|[+\-*/^(),:=<>])/g
      return str.match(re) || []
    }

    function parseCellRef(ref) {
      var m = String(ref).match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/)
      if (!m) return null
      var colStr = m[2].toUpperCase()
      var col = 0
      for (var i = 0; i < colStr.length; i++) col = col * 26 + (colStr.charCodeAt(i) - 64)
      var row = Number(m[4]) - 1
      var c = col - 1
      if (row < 0 || c < 0) return null
      return { r: row, c: c, absCol: !!m[1], absRow: !!m[3] }
    }

    function isCellRef(t) {
      return /^\$?[A-Za-z]{1,3}\$?\d+$/.test(t)
    }

    function evalCellValue(cell, rows, stack, r, c) {
      if (!cell) return ''
      if (cell.f) return evaluateFormula('=' + cell.f, rows, r, c, stack)
      var v = cell.v
      if (typeof v === 'string' && v.trim() !== '' && isFinite(Number(v))) return Number(v)
      return v
    }

    function flattenArg(arg) {
      if (arg && arg.range) {
        var out = []
        for (var i = 0; i < arg.values.length; i++) {
          var v = arg.values[i]
          if (v && v.range) out = out.concat(flattenArg(v))
          else out.push(v)
        }
        return out
      }
      return [arg]
    }

    function callFunction(name, args) {
      var vals = []
      for (var i = 0; i < args.length; i++) vals = vals.concat(flattenArg(args[i]))
      if (name === 'IF') return args[0] ? args[1] : args[2]
      if (name === 'SUM') {
        var sum = 0
        for (var j = 0; j < vals.length; j++) {
          var n = Number(vals[j])
          if (vals[j] !== '' && vals[j] != null && isFinite(n)) sum += n
        }
        return sum
      }
      if (name === 'AVERAGE') {
        var nums = []
        for (var k = 0; k < vals.length; k++) {
          var nk = Number(vals[k])
          if (vals[k] !== '' && vals[k] != null && isFinite(nk)) nums.push(nk)
        }
        return nums.length ? nums.reduce(function (a, b) { return a + b }, 0) / nums.length : 0
      }
      if (name === 'COUNT') {
        var count = 0
        for (var q = 0; q < vals.length; q++) {
          var nq = Number(vals[q])
          if (vals[q] !== '' && vals[q] != null && isFinite(nq)) count += 1
        }
        return count
      }
      if (name === 'MAX') {
        var mx = -Infinity
        var found = false
        for (var m = 0; m < vals.length; m++) {
          var nm = Number(vals[m])
          if (vals[m] !== '' && vals[m] != null && isFinite(nm)) { mx = Math.max(mx, nm); found = true }
        }
        return found ? mx : 0
      }
      if (name === 'MIN') {
        var mn = Infinity
        var foundMin = false
        for (var mi = 0; mi < vals.length; mi++) {
          var nmi = Number(vals[mi])
          if (vals[mi] !== '' && vals[mi] != null && isFinite(nmi)) { mn = Math.min(mn, nmi); foundMin = true }
        }
        return foundMin ? mn : 0
      }
      if (name === 'ROUND') {
        var num = Number(args[0])
        var digits = Number(args[1] == null ? 0 : args[1])
        if (!isFinite(num)) return 0
        var p = Math.pow(10, digits)
        return Math.round(num * p) / p
      }
      if (name === 'ABS') return Math.abs(Number(args[0]))
      if (name === 'INT') return Math.floor(Number(args[0]))
      if (name === 'LEN') return String(args[0] == null ? '' : args[0]).length
      if (name === 'UPPER') return String(args[0] == null ? '' : args[0]).toUpperCase()
      if (name === 'LOWER') return String(args[0] == null ? '' : args[0]).toLowerCase()
      return '#NAME?'
    }

    function evaluateFormula(formula, rows, selfR, selfC, stack) {
      var tokens = tokenizeFormula(formula)
      if (!tokens.length) return ''
      var st = stack || []
      var key = selfR + ':' + selfC
      if (selfR != null && selfC != null) {
        if (st.indexOf(key) !== -1) return '#CYCLE!'
        st.push(key)
      }
      var pos = 0
      function peek() { return tokens[pos] }
      function next() { return tokens[pos++] }
      function expect(t) {
        var x = next()
        if (x !== t) throw new Error('expected ' + t + ' got ' + x)
        return x
      }
      function applyBinary(op, a, b) {
        if (op === '+') return a + b
        if (op === '-') return a - b
        if (op === '*') return a * b
        if (op === '/') return b === 0 ? '#DIV/0!' : a / b
        if (op === '^') return Math.pow(a, b)
        if (op === '=') return a == b
        if (op === '<>') return a != b
        if (op === '!=') return a != b
        if (op === '<') return a < b
        if (op === '>') return a > b
        if (op === '<=') return a <= b
        if (op === '>=') return a >= b
        return null
      }
      function parseArgs() {
        var args = []
        if (peek() === ')') { next(); return args }
        while (true) {
          args.push(parseExpr(0))
          var t = next()
          if (t === ')') break
          if (t !== ',') throw new Error('expected , or ) got ' + t)
        }
        return args
      }
      function parsePrimary() {
        var t = next()
        if (t == null) throw new Error('unexpected end')
        if (t === '(') {
          var v = parseExpr(0)
          expect(')')
          return v
        }
        if (t === '-') return -parsePrimary()
        if (t === '+') return parsePrimary()
        if (t.charAt(0) === '"' || t.charAt(0) === "'") {
          return t.slice(1, -1).replace(/""/g, '"').replace(/''/g, "'")
        }
        if (/^\d/.test(t)) return Number(t)
        if (isCellRef(t)) {
          var ref = parseCellRef(t)
          if (ref && peek() === ':') {
            next()
            var t2 = next()
            var ref2 = parseCellRef(t2)
            if (ref && ref2) {
              var minR = Math.min(ref.r, ref2.r), maxR = Math.max(ref.r, ref2.r)
              var minC = Math.min(ref.c, ref2.c), maxC = Math.max(ref.c, ref2.c)
              var vals = []
              for (var rr = minR; rr <= maxR; rr++) {
                for (var cc = minC; cc <= maxC; cc++) {
                  var cell = rows[rr] && rows[rr][cc]
                  vals.push(evalCellValue(cell, rows, st, rr, cc))
                }
              }
              return { range: true, values: vals }
            }
          }
          if (ref) return evalCellValue(rows[ref.r] && rows[ref.r][ref.c], rows, st, ref.r, ref.c)
          return '#REF!'
        }
        if (/^[A-Za-z_]/.test(t)) {
          if (peek() === '(') { next(); return callFunction(t.toUpperCase(), parseArgs()) }
          return '#NAME?'
        }
        throw new Error('unexpected token ' + t)
      }
      var PREC = { '=': 1, '<>': 1, '!=': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '+': 2, '-': 2, '*': 3, '/': 3, '^': 4 }
      function parseExpr(minPrec) {
        var left = parsePrimary()
        while (peek() && PREC[peek()] !== undefined && PREC[peek()] >= minPrec) {
          var op = next()
          var right = parseExpr(PREC[op] + (op === '^' ? 0 : 1))
          left = applyBinary(op, left, right)
        }
        return left
      }
      try {
        var result = parseExpr(0)
        if (pos !== tokens.length) throw new Error('trailing token ' + tokens[pos])
        return result
      } catch (e) {
        return '#ERR?'
      } finally {
        if (selfR != null && selfC != null) st.pop()
      }
    }

    function cloneRows(rows) {
      return rows.map(function (row) {
        return row.map(function (cell) {
          return Object.assign({}, cell)
        })
      })
    }

    function recalcSheet(sheet) {
      var rows = sheet.rows
      for (var r = 0; r < rows.length; r++) {
        var row = rows[r]
        if (!row) continue
        for (var c = 0; c < row.length; c++) {
          var cell = row[c]
          if (cell && cell.f) cell.v = evaluateFormula('=' + cell.f, rows, r, c, [])
        }
      }
      return sheet
    }

    function formatCellValue(v) {
      if (v === undefined || v === null) return ''
      if (typeof v === 'number' && isFinite(v)) return String(v)
      return String(v)
    }



    function argbToHex(argb) {
      if (!argb) return '#000000'
      var v = String(argb).replace(/^#/, '')
      if (v.length === 8) v = v.slice(2)
      if (v.length !== 6) return '#000000'
      return '#' + v
    }

    function hexToArgb(hex) {
      var v = String(hex || '#000000').replace(/^#/, '')
      if (v.length === 6) return 'FF' + v.toUpperCase()
      return String(hex || '#000000').toUpperCase()
    }


    var FONT_COLORS = ['#000000', '#FFFFFF', '#FF0000', '#FFA500', '#FFFF00', '#008000', '#0000FF', '#800080', '#808080', '#FFC0CB', '#A52A2A', '#FFD700', '#00FF00', '#00FFFF', '#000080', '#FF00FF', '#F0F8FF', '#FAEBD7', '#2F4F4F', '#8B0000']
    var BG_COLORS = ['#FFFFFF', '#FFFF00', '#FFA500', '#FF0000', '#00B050', '#00B0F0', '#0000FF', '#7030A0', '#F2F2F2', '#FFC7CE', '#FFEB9C', '#C6EFCE', '#BDD7EE', '#D9D9D9', '#FFF2CC', '#E2EFDA', '#DDEBF7', '#FCE4D6', '#EDEDED', '#F8CBAD']
    function fontToStyle(font) {
      return {
        fontWeight: font && font.bold ? 'bold' : 'inherit',
        fontStyle: font && font.italic ? 'italic' : 'inherit',
        textDecoration: font && font.underline ? 'underline' : 'none',
        fontSize: font && font.size ? font.size + 'px' : 'inherit',
        color: font && font.color ? argbToHex(font.color) : 'inherit'
      }
    }

    function XlsxEditor({ scope, path, title, customData }) {
      var initial = customData && customData.sheets
        ? { status: 'ready', workbook: customData, activeSheet: customData.activeSheet || (customData.sheets[0] && customData.sheets[0].name) || 'Sheet1', saveState: 'idle' }
        : { status: 'loading', workbook: null, activeSheet: '', saveState: 'idle' }
      var state = React.useState(initial)
      var data = state[0]
      var setData = state[1]
      var selected = React.useState({ r: 0, c: 0 })
      var sel = selected[0]
      var setSel = selected[1]
      var anchorState = React.useState(null)
      var anchor = anchorState[0]
      var setAnchor = anchorState[1]
      var rangeState = React.useState(null)
      var range = rangeState[0]
      var setRange = rangeState[1]
      var menuState = React.useState(null)
      var menu = menuState[0]
      var setMenu = menuState[1]
      var rangeDragState = React.useState(null)
      var rangeDrag = rangeDragState[0]
      var setRangeDrag = rangeDragState[1]
      var editingState = React.useState(null)
      var editing = editingState[0]
      var setEditing = editingState[1]
      var barState = React.useState('')
      var barValue = barState[0]
      var setBarValue = barState[1]
      var historyState = React.useState({ past: [], future: [] })
      var history = historyState[0]
      var setHistory = historyState[1]
      var fillDrag = React.useState(null)
      var drag = fillDrag[0]
      var setDrag = fillDrag[1]
      var colorStartRef = React.useRef(null)
      var resizeState = React.useState(null)
      var resize = resizeState[0]
      var setResize = resizeState[1]
      var opLogState = React.useState([])
      var opLog = opLogState[0]
      var colorMenuState = React.useState(null)
      var colorMenu = colorMenuState[0]
      var setColorMenu = colorMenuState[1]
      var customColorsState = React.useState(['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF'])
      var customColors = customColorsState[0]
      var setCustomColors = customColorsState[1]
      var customColorIndexState = React.useState(null)
      var customColorIndex = customColorIndexState[0]
      var setCustomColorIndex = customColorIndexState[1]
      var customColorInputRef = React.useRef(null)
      var pendingCustomColorState = React.useState(null)
      var pendingCustomColor = pendingCustomColorState[0]
      var setPendingCustomColor = pendingCustomColorState[1]
      var setOpLog = opLogState[1]
      var prevWorkbookRef = React.useRef(null)
      var lastActionRef = React.useRef(null)
      var rangeRef = React.useRef(null)
      ensureExcelStyles()

      React.useEffect(function () {
        if (data.status === 'ready') return
        var cancelled = false
        readExcel(scope, path).then(function (wb) {
          try {
            var firstSheet = wb && wb.sheets && wb.sheets[0]
            var sample = firstSheet && firstSheet.rows && firstSheet.rows[0] && firstSheet.rows[0][0]
            console.log('[dsh-excel-panel] loaded align sample:', sample && sample.align, 'sheets:', wb && wb.sheets && wb.sheets.length)
          } catch (e) {}
          if (cancelled) return
          setData({ status: 'ready', workbook: wb, activeSheet: wb.activeSheet || (wb.sheets[0] && wb.sheets[0].name) || 'Sheet1', saveState: 'idle' })
        }).catch(function (e) {
          if (!cancelled) setData({ status: 'error', workbook: null, activeSheet: '', saveState: 'idle', message: String(e.message || e) })
        })
        return function () { cancelled = true }
      }, [scope.cwd, path, data.status])

      function getActiveSheetData() {
        var wb = data.workbook
        return wb.sheets.find(function (s) { return s.name === data.activeSheet }) || wb.sheets[0] || { name: 'Sheet1', rows: [] }
      }

      function updateCell(r, c, text) {
        logOp('edit', { cell: cellAddress(r, c), value: text })
        var wb = data.workbook
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var rows = cloneRows(sheet.rows || [])
          while (rows.length <= r) rows.push([])
          while (rows[r].length <= c) rows[r].push({ v: '', f: undefined })
          var prev = rows[r][c]
          var raw = String(text == null ? '' : text)
          var isFormula = raw.charAt(0) === '='
          var formula = isFormula ? raw.slice(1) : undefined
          var value = isFormula ? evaluateFormula(raw, rows, r, c, []) : raw
          rows[r][c] = { v: formatCellValue(value), f: formula, bg: prev && prev.bg, font: prev && prev.font, align: prev && prev.align }
          recalcSheet({ name: sheet.name, rows: rows })
          return { name: sheet.name, rows: rows, merges: sheet.merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, saveState: 'dirty' })
      }

      function selectCell(r, c) {
        var sheet = getActiveSheetData()
        var cell = (sheet.rows[r] && sheet.rows[r][c]) || { v: '' }
        setSel({ r: r, c: c })
        setBarValue(cell.f ? '=' + cell.f : cellDisplay(cell.v))
      }

      function clearSelection() {
        setSel({ r: -1, c: -1 })
        applyRange(null)
        setAnchor(null)
        setRangeDrag(null)
        setEditing(null)
        setBarValue('')
      }

      function applyRange(nextRange) {
        setRange(nextRange)
        rangeRef.current = nextRange
      }

      function moveSelection(dr, dc) {
        if (!sel || sel.r === undefined) return
        var nr = Math.max(0, sel.r + dr)
        var nc = Math.max(0, sel.c + dc)
        selectCell(nr, nc)
        applyRange(null)
        setAnchor({ r: nr, c: nc })
      }

      function startEdit(r, c) {
        var sheetData = getActiveSheetData()
        var cell = (sheetData.rows[r] && sheetData.rows[r][c]) || { v: '' }
        selectCell(r, c)
        setEditing({ r: r, c: c, value: cell.f ? '=' + cell.f : cellDisplay(cell.v) })
        setBarValue(cell.f ? '=' + cell.f : cellDisplay(cell.v))
      }

      function onBarChange(text) {
        setBarValue(text)
        if (sel && sel.r !== undefined && sel.r >= 0) {
          updateCell(sel.r, sel.c, text)
          setEditing({ r: sel.r, c: sel.c, value: text })
        }
      }



      function findMergeAt(sheet, r, c) {
        var merges = sheet && sheet.merges || []
        for (var i = 0; i < merges.length; i++) {
          var m = merges[i]
          if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m
        }
        return null
      }

      function mergeRange(selRange) {
        logOp('merge', { range: selRange })
        if (!selRange || (selRange.r1 === selRange.r2 && selRange.c1 === selRange.c2)) return
        var wb = data.workbook
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var merges = (sheet.merges || []).filter(function (m) {
            return !(m.r1 <= selRange.r2 && m.r2 >= selRange.r1 && m.c1 <= selRange.c2 && m.c2 >= selRange.c1)
          })
          merges.push({ r1: selRange.r1, c1: selRange.c1, r2: selRange.r2, c2: selRange.c2 })
          return { name: sheet.name, rows: sheet.rows, merges: merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, saveState: 'dirty' })
        applyRange(null)
      }

      function unmergeAt(r, c) {
        logOp('unmerge', { cell: cellAddress(r, c) })
        var wb = data.workbook
        var hit = null
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var merges = (sheet.merges || []).filter(function (m) {
            var inside = r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2
            if (inside) hit = m
            return !inside
          })
          return { name: sheet.name, rows: sheet.rows, merges: merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        if (!hit) return
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, saveState: 'dirty' })
      }

      function handleCellMouseDown(e, r, c) {
        if (e.button !== 0) return
        e.preventDefault()
        setEditing(null)
        if (e.shiftKey && anchor) {
          var newRange = { r1: Math.min(anchor.r, r), c1: Math.min(anchor.c, c), r2: Math.max(anchor.r, r), c2: Math.max(anchor.c, c) }
          applyRange(newRange)
          selectCell(r, c)
        } else {
          setAnchor({ r: r, c: c })
          applyRange(null)
          selectCell(r, c)
          setRangeDrag({ anchor: { r: r, c: c }, current: { r: r, c: c } })
        }
      }


      function closeMenu() {
        setMenu(null)
      }

      function openMenu(e, r, c) {
        e.preventDefault()
        e.stopPropagation()
        setMenu({ x: e.clientX, y: e.clientY, r: r, c: c })
      }

      function copyCell(r, c) {
        var sheetData = getActiveSheetData()
        var cell = (sheetData.rows[r] && sheetData.rows[r][c]) || { v: '' }
        var text = cell.f ? '=' + cell.f : cellDisplay(cell.v)
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).catch(function () {})
        }
        closeMenu()
      }

      function pasteCell(r, c) {
        if (navigator.clipboard && navigator.clipboard.readText) {
          navigator.clipboard.readText().then(function (text) {
            updateCell(r, c, text)
            closeMenu()
          }).catch(function () { closeMenu() })
        } else {
          closeMenu()
        }
      }

      function clearCell(r, c) {
        updateCell(r, c, '')
        closeMenu()
      }

      function insertRow(r) {
        logOp('insertRow', { row: r + 1 })
        var wb = data.workbook
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var rows = cloneRows(sheet.rows || [])
          while (rows.length < r) rows.push([])
          rows.splice(r, 0, [])
          var merges = (sheet.merges || []).map(function (m) {
            var nm = { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 }
            if (nm.r1 >= r) nm.r1 += 1
            if (nm.r2 >= r) nm.r2 += 1
            return nm
          })
          return { name: sheet.name, rows: rows, merges: merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, saveState: 'dirty' })
        closeMenu()
      }

      function deleteRow(r) {
        logOp('deleteRow', { row: r + 1 })
        var wb = data.workbook
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var rows = cloneRows(sheet.rows || [])
          if (r < 0) return sheet
          while (rows.length <= r) rows.push([])
          rows.splice(r, 1)
          var merges = (sheet.merges || []).filter(function (m) {
            return !(r >= m.r1 && r <= m.r2)
          }).map(function (m) {
            var nm = { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 }
            if (nm.r1 > r) nm.r1 -= 1
            if (nm.r2 > r) nm.r2 -= 1
            return nm
          })
          return { name: sheet.name, rows: rows, merges: merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, saveState: 'dirty' })
        closeMenu()
      }

      function insertCol(c, r) {
        logOp('insertCol', { col: c + 1 })
        var wb = data.workbook
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var rows = cloneRows(sheet.rows || [])
          while (rows.length <= (r || 0)) rows.push([])
          rows = rows.map(function (row) {
            var nr = row.slice()
            while (nr.length < c) nr.push({ v: '', f: undefined })
            nr.splice(c, 0, { v: '', f: undefined })
            return nr
          })
          var merges = (sheet.merges || []).map(function (m) {
            var nm = { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 }
            if (nm.c1 >= c) nm.c1 += 1
            if (nm.c2 >= c) nm.c2 += 1
            return nm
          })
          return { name: sheet.name, rows: rows, merges: merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, saveState: 'dirty' })
        closeMenu()
      }

      function deleteCol(c, r) {
        logOp('deleteCol', { col: c + 1 })
        var wb = data.workbook
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var rows = cloneRows(sheet.rows || [])
          while (rows.length <= (r || 0)) rows.push([])
          rows = rows.map(function (row) {
            var nr = row.slice()
            while (nr.length <= c) nr.push({ v: '', f: undefined })
            if (c >= 0 && c < nr.length) nr.splice(c, 1)
            return nr
          })
          var merges = (sheet.merges || []).filter(function (m) {
            return !(c >= m.c1 && c <= m.c2)
          }).map(function (m) {
            var nm = { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 }
            if (nm.c1 > c) nm.c1 -= 1
            if (nm.c2 > c) nm.c2 -= 1
            return nm
          })
          return { name: sheet.name, rows: rows, merges: merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, saveState: 'dirty' })
        closeMenu()
      }

      React.useEffect(function () {
        function onWindowMouseDown() { setMenu(null); setColorMenu(null) }
        window.addEventListener('mousedown', onWindowMouseDown)
        return function () { window.removeEventListener('mousedown', onWindowMouseDown) }
      }, [])

      React.useEffect(function () {
        if (!rangeDrag) return
        function onMove(e) {
          var el = document.elementFromPoint(e.clientX, e.clientY)
          var node = el && el.closest ? el.closest('[data-cell]') : null
          if (!node) return
          var r = Number(node.getAttribute('data-r'))
          var c = Number(node.getAttribute('data-c'))
          if (Number.isFinite(r) && Number.isFinite(c)) {
            var newRange = { r1: Math.min(rangeDrag.anchor.r, r), c1: Math.min(rangeDrag.anchor.c, c), r2: Math.max(rangeDrag.anchor.r, r), c2: Math.max(rangeDrag.anchor.c, c) }
            applyRange(newRange)
            setRangeDrag(function (prev) { return prev ? { anchor: prev.anchor, current: { r: r, c: c } } : prev })
          }
        }
        function onUp() {
          setRangeDrag(null)
        }
        window.addEventListener('mousemove', onMove)
        window.addEventListener('mouseup', onUp)
        return function () {
          window.removeEventListener('mousemove', onMove)
          window.removeEventListener('mouseup', onUp)
        }
      }, [rangeDrag])

      function updateSheetSize(type, index, value) {
        logOp('resize', { type: type, index: index, value: Math.round(value) })
        setData(function (prev) {
          if (!prev || !prev.workbook) return prev
          var sheets = prev.workbook.sheets.map(function (sheet) {
            if (sheet.name !== prev.activeSheet) return sheet
            var next = Object.assign({}, sheet)
            if (type === 'col') {
              var cw = Object.assign({}, next.colWidths || {})
              cw[index] = Math.max(20, value)
              next.colWidths = cw
            } else {
              var rh = Object.assign({}, next.rowHeights || {})
              rh[index] = Math.max(15, value)
              next.rowHeights = rh
            }
            return next
          })
          return { ...prev, workbook: { ...prev.workbook, sheets: sheets }, saveState: 'dirty' }
        })
      }

      function startColResize(e, c) {
        e.preventDefault()
        e.stopPropagation()
        var sheetData = getActiveSheetData()
        var w = (sheetData.colWidths && sheetData.colWidths[c]) || 70
        setResize({ type: 'col', index: c, startX: e.clientX, startSize: w })
      }

      function startRowResize(e, r) {
        e.preventDefault()
        e.stopPropagation()
        var sheetData = getActiveSheetData()
        var h = (sheetData.rowHeights && sheetData.rowHeights[r]) || 26
        setResize({ type: 'row', index: r, startY: e.clientY, startSize: h })
      }

      function logOp(action, extra) {
        var entry = Object.assign({ action: action, time: new Date().toLocaleTimeString() }, extra || {})
        setOpLog(function (prev) { return [entry].concat(prev).slice(0, 50) })
        lastActionRef.current = entry
      }

      // Full-snapshot logging disabled temporarily to avoid freezing on every keystroke.

      React.useEffect(function () {
        if (!resize) return
        function onMove(e) {
          if (resize.type === 'col') {
            var dx = e.clientX - resize.startX
            updateSheetSize('col', resize.index, resize.startSize + dx / 8)
          } else {
            var dy = e.clientY - resize.startY
            updateSheetSize('row', resize.index, resize.startSize + dy * 0.75)
          }
        }
        function onUp() { setResize(null) }
        window.addEventListener('mousemove', onMove)
        window.addEventListener('mouseup', onUp)
        return function () {
          window.removeEventListener('mousemove', onMove)
          window.removeEventListener('mouseup', onUp)
        }
      }, [resize])

      function applyStyleToWorkbook(wb, r, c, patch) {
        var rangeToApply = rangeRef.current && rangeRef.current.r1 !== undefined ? rangeRef.current : { r1: r, c1: c, r2: r, c2: c }
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var rows = cloneRows(sheet.rows || [])
          for (var rr = rangeToApply.r1; rr <= rangeToApply.r2; rr++) {
            for (var cc = rangeToApply.c1; cc <= rangeToApply.c2; cc++) {
              while (rows.length <= rr) rows.push([])
              while (rows[rr].length <= cc) rows[rr].push({ v: '', f: undefined, font: undefined })
              var prev = rows[rr][cc] || { v: '', f: undefined }
              var nextCell = { v: prev.v, f: prev.f, bg: prev.bg, font: prev.font, align: prev.align }
              if (Object.prototype.hasOwnProperty.call(patch, 'bg')) nextCell.bg = patch.bg
              if (Object.prototype.hasOwnProperty.call(patch, 'align')) nextCell.align = patch.align
              var fontPatch = {}
              for (var key in patch) {
                if (key !== 'bg' && key !== 'align') fontPatch[key] = patch[key]
              }
              if (Object.keys(fontPatch).length > 0) nextCell.font = Object.assign({}, prev.font || {}, fontPatch)
              rows[rr][cc] = nextCell
            }
          }
          return { name: sheet.name, rows: rows, merges: sheet.merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        return { ...wb, sheets: sheets }
      }

      function updateStyle(r, c, patch) {
        logOp('style', { cell: cellAddress(r, c), patch: patch })
        var nextWb = applyStyleToWorkbook(data.workbook, r, c, patch)
        setHistory(function (h) { return { past: h.past.concat([data.workbook]), future: [] } })
        setData({ ...data, workbook: nextWb, saveState: 'dirty' })
      }

      function updateStyleNoHistory(r, c, patch) {
        setData(function (prev) {
          if (!prev || !prev.workbook) return prev
          return { ...prev, workbook: applyStyleToWorkbook(prev.workbook, r, c, patch), saveState: 'dirty' }
        })
      }

      function fillRange(start, end) {
        logOp('fill', { from: cellAddress(start.r, start.c), to: cellAddress(end.r, end.c) })
        var wb = data.workbook
        var sheetData = getActiveSheetData()
        var srcRow = sheetData.rows[start.r] || []
        var srcCell = srcRow[start.c] || { v: '' }
        var minR = Math.min(start.r, end.r), maxR = Math.max(start.r, end.r)
        var minC = Math.min(start.c, end.c), maxC = Math.max(start.c, end.c)
        var sheets = wb.sheets.map(function (sheet) {
          if (sheet.name !== data.activeSheet) return sheet
          var rows = cloneRows(sheet.rows || [])
          for (var r = minR; r <= maxR; r++) {
            for (var c = minC; c <= maxC; c++) {
              if (r === start.r && c === start.c) continue
              while (rows.length <= r) rows.push([])
              while (rows[r].length <= c) rows[r].push({ v: '', f: undefined })
              if (srcCell.f) {
                var shifted = shiftFormula(srcCell.f, r - start.r, c - start.c)
                rows[r][c] = { v: evaluateFormula('=' + shifted, rows, r, c, []), f: shifted, bg: rows[r][c] && rows[r][c].bg, font: srcCell.font, align: srcCell.align }
              } else {
                rows[r][c] = { v: srcCell.v, f: undefined, bg: rows[r][c] && rows[r][c].bg, font: srcCell.font, align: srcCell.align }
              }
            }
          }
          recalcSheet({ name: sheet.name, rows: rows })
          return { name: sheet.name, rows: rows, merges: sheet.merges, colWidths: sheet.colWidths, rowHeights: sheet.rowHeights }
        })
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, saveState: 'dirty' })
      }

      React.useEffect(function () {
        if (!drag) return
        var onMove = function (e) {
          var el = document.elementFromPoint(e.clientX, e.clientY)
          if (!el) return
          var node = el.closest ? el.closest('[data-cell]') : null
          if (!node) return
          var r = Number(node.getAttribute('data-r'))
          var c = Number(node.getAttribute('data-c'))
          if (Number.isFinite(r) && Number.isFinite(c)) {
            setDrag(function (prev) { return prev ? { start: prev.start, end: { r: r, c: c } } : prev })
          }
        }
        var onUp = function () {
          if (drag) fillRange(drag.start, drag.end)
          setDrag(null)
        }
        window.addEventListener('mousemove', onMove)
        window.addEventListener('mouseup', onUp)
        return function () {
          window.removeEventListener('mousemove', onMove)
          window.removeEventListener('mouseup', onUp)
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
      }, [drag])

      function undo() {
        if (!history.past.length) return
        var past = history.past.slice()
        var prev = past.pop()
        setHistory({ past: past, future: [data.workbook].concat(history.future) })
        setData({ ...data, workbook: prev, saveState: 'dirty' })
        setEditing(null)
        setBarValue('')
      }

      function redo() {
        if (!history.future.length) return
        var future = history.future.slice()
        var next = future.shift()
        setHistory({ past: history.past.concat([data.workbook]), future: future })
        setData({ ...data, workbook: next, saveState: 'dirty' })
        setEditing(null)
        setBarValue('')
      }

      React.useEffect(function () {
        function onKeyDown(e) {
          var mod = e.ctrlKey || e.metaKey
          if (!mod) {
            if (editing) return
            if (e.key === 'Enter') {
              e.preventDefault()
              if (sel && sel.r !== undefined && sel.r >= 0) startEdit(sel.r, sel.c)
              return
            }
            if (e.key === 'Tab') {
              e.preventDefault()
              moveSelection(0, 1)
              return
            }
            return
          }
          var key = e.key.toLowerCase()
          if (key === 'z' && e.shiftKey) {
            e.preventDefault()
            redo()
          } else if (key === 'y') {
            e.preventDefault()
            redo()
          } else if (key === 'z') {
            e.preventDefault()
            undo()
          }
        }
        window.addEventListener('keydown', onKeyDown)
        return function () { window.removeEventListener('keydown', onKeyDown) }
      }, [data, history, sel, editing, range])

      function refresh() {
        readExcel(scope, path).then(function (wb) {
          setData(function (prev) {
            if (prev.status !== 'ready') return prev
            var active = prev.activeSheet && wb.sheets.some(function (s) { return s.name === prev.activeSheet })
              ? prev.activeSheet
              : (wb.activeSheet || (wb.sheets[0] && wb.sheets[0].name) || 'Sheet1')
            var next = { ...prev, workbook: wb, activeSheet: active, saveState: prev.saveState === 'dirty' ? prev.saveState : 'idle' }
            return JSON.stringify(prev.workbook) === JSON.stringify(wb) ? prev : next
          })
          setEditing(null)
          setBarValue('')
        }).catch(function () {})
      }

      // 自动刷新已关闭：需要时请手动点击“刷新”

      function addSheet() {
        logOp('addSheet', { name: 'Sheet' })
        var wb = data.workbook
        var used = {}
        wb.sheets.forEach(function (sheet) { used[sheet.name] = true })
        var base = 'Sheet'
        var i = 1
        while (used[base + i]) i += 1
        var name = base + i
        var sheets = wb.sheets.concat([{ name: name, rows: [], merges: [] }])
        setHistory(function (h) { return { past: h.past.concat([wb]), future: [] } })
        setData({ ...data, workbook: { ...wb, sheets: sheets }, activeSheet: name, saveState: 'dirty' })
        setEditing(null)
        setSel({ r: 0, c: 0 })
        setBarValue('')
        applyRange(null)
        setAnchor(null)
        setRangeDrag(null)
      }

      function save() {
        setData({ ...data, saveState: 'saving' })
        var target = officeTarget(scope, path)
        post('/excel-panel/write', { root: target.root, path: target.path, kind: 'xlsx', payload: { sheets: data.workbook.sheets } }).then(function (res) {
          if (!res.ok) { var errDetail = res && res.error; var errMsg = typeof errDetail === 'string' ? errDetail : (errDetail && errDetail.message) || 'save failed'; throw new Error(errMsg) }
          setData({ ...data, saveState: 'saved' })
        }).catch(function (e) {
          console.error('[dsh-excel-panel] save failed:', e)
          var rawMsg = String(e && e.message || e)
          var friendly = /EBUSY|resource busy|locked/i.test(rawMsg)
            ? '文件被其他程序占用，请关闭 Excel/WPS 后重试'
            : rawMsg
          setData({ ...data, saveState: 'error', message: friendly })
        })
      }


      function unlockAndResave() {
        var target = officeTarget(scope, path)
        post('/excel-panel/unlock', { root: target.root, path: target.path }).then(function (res) {
          if (!res.ok) { var errDetail = res && res.error; var errMsg = typeof errDetail === 'string' ? errDetail : (errDetail && errDetail.message) || 'unlock failed'; throw new Error(errMsg) }
          save()
        }).catch(function (e) {
          console.error('[dsh-excel-panel] unlock failed:', e)
          setData({ ...data, saveState: 'error', message: '关闭占用程序失败：' + String(e && e.message || e) })
        })
      }
      // 自动保存已关闭：请手动点击“保存”

      if (data.status === 'loading') return h('div', { style: { padding: 16, color: '#888' } }, '加载中…')
      if (data.status === 'error') return h('div', { style: { padding: 16, color: '#c00' } }, '加载失败：' + (data.message || ''))
      var wb = data.workbook
      var sheet = wb.sheets.find(function (s) { return s.name === data.activeSheet }) || wb.sheets[0] || { name: 'Sheet1', rows: [] }
      var rows = sheet.rows || []
      var dataColCount = rows.reduce(function (m, row) { return Math.max(m, row.length) }, 0)
      var selColSpan = Math.max(sel && sel.c !== undefined ? sel.c + 1 : 0, range ? range.c2 + 1 : 0)
      var selRowSpan = Math.max(sel && sel.r !== undefined && sel.r >= 0 ? sel.r + 1 : 0, range ? range.r2 + 1 : 0)
      var colCount = Math.max(1, dataColCount, selColSpan)
      var rowCount = Math.max(1, rows.length, selRowSpan)
      var selCell = (rows[sel.r] && rows[sel.r][sel.c]) || { v: '' }
      var selLabel = (sel && sel.r !== undefined && sel.r >= 0) ? cellAddress(sel.r, sel.c) : ''
      var selFormula = selCell.f ? '=' + selCell.f : cellDisplay(selCell.v)
      var currentFontSize = (function () {
        if (!sel || sel.r === undefined || sel.r < 0) return ''
        var rng = rangeRef.current
        if (!rng) return (selCell.font && selCell.font.size) || 11
        var sizes = {}
        for (var rr = rng.r1; rr <= rng.r2; rr++) {
          for (var cc = rng.c1; cc <= rng.c2; cc++) {
            var c = (rows[rr] && rows[rr][cc]) || {}
            sizes[(c.font && c.font.size) || 11] = true
          }
        }
        var keys = Object.keys(sizes)
        return keys.length === 1 ? Number(keys[0]) : ''
      })()
      function styleActiveInSelection(prop) {
        if (!sel || sel.r === undefined || sel.r < 0) return false
        var rng = rangeRef.current
        if (!rng) return !!(selCell.font && selCell.font[prop])
        var any = false, all = true, count = 0
        for (var rr = rng.r1; rr <= rng.r2; rr++) {
          for (var cc = rng.c1; cc <= rng.c2; cc++) {
            var c = (rows[rr] && rows[rr][cc]) || {}
            count++
            if (c.font && c.font[prop]) any = true; else all = false
          }
        }
        return count > 0 && all
      }
      var isBoldActive = styleActiveInSelection('bold')
      var isItalicActive = styleActiveInSelection('italic')
      var isUnderlineActive = styleActiveInSelection('underline')
      return h('div', { style: { display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }, onMouseDown: function (e) { var t = e.target; if (t && t.closest && (t.closest('[data-cell]') || t.closest('button,input,select'))) return; clearSelection() } },
        h('div', { style: { display: 'flex', alignItems: 'center', gap: 8, padding: '4px 8px', borderBottom: '1px solid #ddd', fontSize: 12 } },
          h('span', { style: { fontWeight: 600 } }, title + ' v0.5.0'),
          h('span', { style: { flex: 1 } }),
          h('button', { onClick: refresh, style: { padding: '3px 10px', cursor: 'pointer', marginRight: 4 } }, '刷新'),
          h('button', { onClick: save, disabled: data.saveState === 'saving' || data.saveState === 'idle', style: { padding: '3px 12px', cursor: 'pointer' } },
            data.saveState === 'saving' ? '保存中…' : data.saveState === 'saved' ? '已保存' : data.saveState === 'error' ? '保存失败' : '保存')
        ),
        h('div', { style: { display: 'flex', alignItems: 'center', gap: 8, padding: '2px 8px', borderBottom: '1px solid #ddd', fontSize: 12, background: '#fafafa' } },
          h('span', { style: { width: 52, color: '#333', fontWeight: 600 } }, selLabel || ''),
          h('span', { style: { color: '#888' } }, 'fx'),
          h('input', {
            style: { flex: 1, border: '1px solid #ddd', borderRadius: 4, padding: '3px 8px', font: 'inherit', outline: 'none' },
            value: barValue,
            onFocus: function () { setBarValue(selFormula) },
            onChange: function (e) { onBarChange(e.target.value) }
          })
        ),
        h('div', { style: { display: 'flex', alignItems: 'center', gap: 6, padding: '3px 8px', borderBottom: '1px solid #ddd', background: '#fff', fontSize: 12 } },
          h('select', {
            value: currentFontSize,
            onChange: function (e) { if (sel && sel.r !== undefined && sel.r >= 0) updateStyle(sel.r, sel.c, { size: Number(e.target.value) }) },
            style: { padding: '2px 4px', border: '1px solid #ccc', borderRadius: 4 }
          }, [h('option', { key: 'auto', value: '' }, '字号'), 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map(function (n) { return typeof n === 'number' ? h('option', { key: n, value: n }, n) : n })),
          h('button', {
            onClick: function () { if (sel && sel.r !== undefined && sel.r >= 0) updateStyle(sel.r, sel.c, { bold: !isBoldActive }) },
            title: '加粗',
            style: { padding: '2px 8px', cursor: 'pointer', fontWeight: 'bold', border: isBoldActive ? '1px solid #165dff' : '1px solid #ccc', borderRadius: 4, background: isBoldActive ? '#eef2ff' : '#fff' }
          }, 'B'),
          h('button', {
            onClick: function () { if (sel && sel.r !== undefined && sel.r >= 0) updateStyle(sel.r, sel.c, { italic: !isItalicActive }) },
            title: '斜体',
            style: { padding: '2px 8px', cursor: 'pointer', fontStyle: 'italic', border: isItalicActive ? '1px solid #165dff' : '1px solid #ccc', borderRadius: 4, background: isItalicActive ? '#eef2ff' : '#fff' }
          }, 'I'),
          h('button', {
            onClick: function () { if (sel && sel.r !== undefined && sel.r >= 0) updateStyle(sel.r, sel.c, { underline: !isUnderlineActive }) },
            title: '下划线',
            style: { padding: '2px 8px', cursor: 'pointer', textDecoration: 'underline', border: isUnderlineActive ? '1px solid #165dff' : '1px solid #ccc', borderRadius: 4, background: isUnderlineActive ? '#eef2ff' : '#fff' }
          }, 'U'),
          h('button', {
            onClick: function () { if (sel && sel.r !== undefined && sel.r >= 0) updateStyle(sel.r, sel.c, { align: 'left' }) },
            title: '左对齐',
            style: { padding: '2px 8px', cursor: 'pointer', border: selCell.align === 'left' ? '1px solid #165dff' : '1px solid #ccc', borderRadius: 4, background: selCell.align === 'left' ? '#eef2ff' : '#fff' }
          }, '⇤'),
          h('button', {
            onClick: function () { if (sel && sel.r !== undefined && sel.r >= 0) updateStyle(sel.r, sel.c, { align: 'center' }) },
            title: '居中',
            style: { padding: '2px 8px', cursor: 'pointer', border: selCell.align === 'center' ? '1px solid #165dff' : '1px solid #ccc', borderRadius: 4, background: selCell.align === 'center' ? '#eef2ff' : '#fff' }
          }, '≡'),
          h('button', {
            onClick: function () { if (sel && sel.r !== undefined && sel.r >= 0) updateStyle(sel.r, sel.c, { align: 'right' }) },
            title: '右对齐',
            style: { padding: '2px 8px', cursor: 'pointer', border: selCell.align === 'right' ? '1px solid #165dff' : '1px solid #ccc', borderRadius: 4, background: selCell.align === 'right' ? '#eef2ff' : '#fff' }
          }, '⇥'),
          h('button', {
            onClick: function (e) { e.stopPropagation(); setColorMenu({ type: 'font', x: e.clientX, y: e.clientY }) },
            title: '字体颜色',
            style: { padding: '2px 8px', cursor: 'pointer', border: '1px solid #ccc', borderRadius: 4, background: '#fff', color: '#333', textDecoration: 'underline', textUnderlineOffset: 2 }
          }, 'A'),
          h('button', {
            onClick: function (e) { e.stopPropagation(); setColorMenu({ type: 'bg', x: e.clientX, y: e.clientY }) },
            title: '背景颜色',
            style: { padding: '2px 8px', cursor: 'pointer', border: '1px solid #ccc', borderRadius: 4, background: '#fff', color: '#333' }
          }, '底色'),
        ),

        h('div', { style: { flex: 1, overflow: 'auto', padding: 8 } },
          h('table', { style: { borderCollapse: 'collapse', fontSize: 13 } },
            h('tbody', null,
              h('tr', null,
                h('td', { style: headerStyle() }, ''),
                Array.from({ length: colCount }, function (_, c) {
                  var colW = (sheet.colWidths && sheet.colWidths[c]) || 70
                  return h('td', {
                    key: c,
                    style: Object.assign(headerStyle(), { width: colW, minWidth: colW, position: 'relative' }),
                    children: [
                      colName(c + 1),
                      h('div', {
                        onMouseDown: function (e) { startColResize(e, c) },
                        title: '拖动调整列宽',
                        style: { position: 'absolute', right: -3, top: 0, bottom: 0, width: 6, cursor: 'col-resize', zIndex: 3 }
                      })
                    ]
                  })
                })
              ),
              Array.from({ length: rowCount }, function (_, r) {
                var row = rows[r] || []
                return h('tr', { key: r },
                  h('td', {
                    style: Object.assign(headerStyle(), { height: (sheet.rowHeights && sheet.rowHeights[r]) || 26, minHeight: (sheet.rowHeights && sheet.rowHeights[r]) || 26, position: 'relative' }),
                    children: [
                      r + 1,
                      h('div', {
                        onMouseDown: function (e) { startRowResize(e, r) },
                        title: '拖动调整行高',
                        style: { position: 'absolute', left: 0, right: 0, bottom: -3, height: 6, cursor: 'row-resize', zIndex: 3 }
                      })
                    ]
                  }),
                  Array.from({ length: colCount }, function (_, c) {
                    var merge = findMergeAt(sheet, r, c)
                    if (merge && (r !== merge.r1 || c !== merge.c1)) return null
                    var cell = row[c] || { v: '' }
                    var bg = typeof cell.bg === 'string' && cell.bg.length >= 6 ? '#' + cell.bg.slice(2) : undefined
                    var isSelected = sel.r === r && sel.c === c
                    var inRange = range
                      ? (r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2)
                      : (drag && r >= Math.min(drag.start.r, drag.end.r) && r <= Math.max(drag.start.r, drag.end.r) && c >= Math.min(drag.start.c, drag.end.c) && c <= Math.max(drag.start.c, drag.end.c))
                    var isFormula = !!cell.f
                    var displayValue = (editing && editing.r === r && editing.c === c) ? editing.value : cellDisplay(cell.v)
                    var fontStyle = fontToStyle(cell.font)
                    var tdStyle = {
                      border: '1px solid #ddd',
                      padding: 0,
                      position: 'relative',
                      textAlign: cell.align || 'left',
                      height: (sheet.rowHeights && sheet.rowHeights[r]) || 26,
                      userSelect: 'none',
                      WebkitUserSelect: 'none',
                      background: bg || undefined,
                      boxShadow: isSelected ? (inRange ? 'inset 0 0 0 1000px rgba(22,93,255,0.15), inset 0 0 0 2px #165dff' : 'inset 0 0 0 2px #165dff') : (inRange ? 'inset 0 0 0 1000px rgba(22,93,255,0.25), inset 0 0 0 1px rgba(22,93,255,0.55)' : undefined),
                      animation: isSelected ? 'dshExcelSelectIn .15s ease-out' : undefined
                    }
                    var rowSpan = merge ? merge.r2 - merge.r1 + 1 : undefined
                    var colSpan = merge ? merge.c2 - merge.c1 + 1 : undefined
                    return h('td', {
                      key: c,
                      'data-cell': '',
                      'data-r': r,
                      'data-c': c,
                      rowSpan: rowSpan,
                      colSpan: colSpan,
                      onMouseDown: function (e) { handleCellMouseDown(e, r, c) },
                      onContextMenu: function (e) { openMenu(e, r, c) },
                      onDoubleClick: function () { startEdit(r, c) },
                      style: tdStyle,
                      children: [
                        editing && editing.r === r && editing.c === c
                          ? h('input', {
                              style: Object.assign({ width: '100%', minWidth: 70, border: 'none', padding: '2px 6px', background: 'transparent', font: 'inherit', outline: 'none', textAlign: cell.align || 'left' }, fontStyle, { color: fontStyle.color !== 'inherit' ? fontStyle.color : (isFormula ? '#165dff' : 'inherit'), fontWeight: fontStyle.fontWeight !== 'inherit' ? fontStyle.fontWeight : (isFormula ? 600 : 'inherit') }),
                              value: displayValue,
                              autoFocus: true,
                              onChange: function (e) {
                                updateCell(r, c, e.target.value)
                                setEditing({ r: r, c: c, value: e.target.value })
                                if (sel.r === r && sel.c === c) setBarValue(e.target.value)
                              },
                              onBlur: function () { setEditing(null) },
                              onKeyDown: function (e) {
                                if (e.key === 'Enter') {
                                  e.preventDefault()
                                  setEditing(null)
                                  moveSelection(1, 0)
                                } else if (e.key === 'Tab') {
                                  e.preventDefault()
                                  setEditing(null)
                                  moveSelection(0, 1)
                                } else if (e.key === 'Escape') {
                                  e.preventDefault()
                                  setEditing(null)
                                }
                              },
                              onMouseDown: function (e) { e.stopPropagation() }
                            })
                          : h('div', {
                              style: Object.assign({ width: '100%', minWidth: 70, padding: '2px 6px', boxSizing: 'border-box', cursor: 'cell', textAlign: cell.align || 'left' }, fontStyle, { color: fontStyle.color !== 'inherit' ? fontStyle.color : (isFormula ? '#165dff' : 'inherit'), fontWeight: fontStyle.fontWeight !== 'inherit' ? fontStyle.fontWeight : (isFormula ? 600 : 'inherit') }),
                              onDoubleClick: function () {
                                startEdit(r, c)
                              }
                            }, displayValue),
                        isSelected ? h('div', {
                          style: { position: 'absolute', right: 0, bottom: 0, width: 8, height: 8, background: '#165dff', cursor: 'crosshair', zIndex: 5 },
                          onMouseDown: function (e) { e.preventDefault(); e.stopPropagation(); setDrag({ start: { r: r, c: c }, end: { r: r, c: c } }) },
                          title: '拖动填充公式/值'
                        }) : null
                      ]
                    })
                  })
                )
              })
            )
          )
        ),

        h('div', { style: { maxHeight: 80, overflowY: 'auto', borderTop: '1px solid #eee', padding: '4px 8px', fontSize: 11, color: '#666', background: '#fcfcfc' } },
          opLog.length === 0 ? '暂无操作记录' : opLog.map(function (entry, i) {
            return h('div', { key: i }, '[' + entry.time + '] ' + entry.action + (entry.cell ? ' ' + entry.cell : '') + (entry.value !== undefined ? ' = ' + entry.value : ''))
          })
        ),
        h('div', { style: { display: 'flex', alignItems: 'center', gap: 4, padding: '4px 8px', borderTop: '1px solid #ddd', background: '#f8f9fa', fontSize: 12 } },
          wb.sheets.map(function (s) {
            return h('button', {
              key: s.name,
              onClick: function () { setData({ ...data, activeSheet: s.name, saveState: data.saveState }); setEditing(null); setSel({ r: 0, c: 0 }); setBarValue(''); applyRange(null); setAnchor(null); setRangeDrag(null) },
              style: { padding: '3px 10px', cursor: 'pointer', background: s.name === data.activeSheet ? '#fff' : 'transparent', color: s.name === data.activeSheet ? '#165dff' : '#666', border: s.name === data.activeSheet ? '1px solid #165dff' : '1px solid transparent', borderBottom: s.name === data.activeSheet ? '2px solid #165dff' : '2px solid transparent', borderRadius: '4px 4px 0 0', fontWeight: s.name === data.activeSheet ? 600 : 400 }
            }, s.name)
          }),
          h('button', {
            onClick: addSheet,
            title: '新建工作表',
            style: { padding: '2px 10px', cursor: 'pointer', border: '1px solid #ccc', borderRadius: 4, background: '#fff', color: '#333', fontSize: 14, lineHeight: 1.2 }
          }, '+')
        ),
        data.saveState === 'error' ? h('div', { style: { padding: '4px 8px', color: '#c00', fontSize: 12, display: 'flex', alignItems: 'center', gap: 8 } }, [data.message, /占用/.test(data.message || '') ? h('button', { onClick: unlockAndResave, style: { padding: '2px 10px', cursor: 'pointer', color: '#c00', border: '1px solid #c00', borderRadius: 4, background: '#fff' } }, '关闭占用程序并重试') : null]) : null,
        colorMenu ? h('div', {
          style: { position: 'fixed', zIndex: 9999, top: colorMenu.y, left: colorMenu.x, background: '#fff', border: '1px solid #ccc', borderRadius: 6, boxShadow: '0 2px 8px rgba(0,0,0,0.15)', padding: 6, minWidth: 160 },
          onMouseDown: function (e) { e.stopPropagation() },
          children: [
            h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(5, 20px)', gap: 4 } },
              (colorMenu.type === 'font' ? FONT_COLORS : BG_COLORS).map(function (color) {
                return h('button', {
                  key: color,
                  onClick: function () {
                    if (sel && sel.r !== undefined && sel.r >= 0) {
                      updateStyle(sel.r, sel.c, colorMenu.type === 'font' ? { color: hexToArgb(color) } : { bg: hexToArgb(color) })
                    }
                    setColorMenu(null)
                  },
                  style: { width: 20, height: 20, border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer', background: color, padding: 0 }
                })
              })
            ),
            h('div', { style: { display: 'flex', alignItems: 'center', gap: 4, marginTop: 6, borderTop: '1px solid #eee', paddingTop: 4 } },
              h('span', { style: { fontSize: 11, color: '#888', marginRight: 2 } }, '自定义'),
              customColors.map(function (color, i) {
                return h('button', {
                  key: i,
                  onClick: function () {
                    if (sel && sel.r !== undefined && sel.r >= 0) {
                      updateStyle(sel.r, sel.c, colorMenu.type === 'font' ? { color: hexToArgb(color) } : { bg: hexToArgb(color) })
                    }
                    setColorMenu(null)
                  },
                  onContextMenu: function (e) {
                    e.preventDefault()
                    e.stopPropagation()
                    setCustomColorIndex(i)
                    setPendingCustomColor(null)
                    setTimeout(function () { if (customColorInputRef.current) customColorInputRef.current.click() }, 0)
                  },
                  title: '左键使用，右键设置自定义颜色',
                  style: { width: 18, height: 18, border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer', background: color, padding: 0 }
                })
              }),
              pendingCustomColor !== null && customColorIndex !== null ? h('button', {
                onClick: function () {
                  var picked = pendingCustomColor
                  setCustomColors(function (prev) { var next = prev.slice(); next[customColorIndex] = picked; return next })
                  if (sel && sel.r !== undefined && sel.r >= 0) {
                    updateStyle(sel.r, sel.c, colorMenu.type === 'font' ? { color: hexToArgb(picked) } : { bg: hexToArgb(picked) })
                  }
                  setPendingCustomColor(null)
                  setCustomColorIndex(null)
                  setColorMenu(null)
                },
                title: '确认自定义颜色',
                style: { width: 20, height: 20, border: '1px solid #165dff', borderRadius: 3, cursor: 'pointer', background: '#eef2ff', color: '#165dff', fontWeight: 'bold', padding: 0, fontSize: 14, lineHeight: '18px' }
              }, '✓') : null,
              h('input', {
                ref: customColorInputRef,
                type: 'color',
                style: { position: 'absolute', width: 0, height: 0, opacity: 0, border: 'none', padding: 0 },
                onChange: function (e) { setPendingCustomColor(e.target.value) }
              })
            )
          ]
        }) : null,
        menu ? h('div', {
          style: { position: 'fixed', zIndex: 9999, top: menu.y, left: menu.x, background: '#fff', border: '1px solid #ccc', borderRadius: 6, boxShadow: '0 2px 8px rgba(0,0,0,0.15)', padding: '4px 0', fontSize: 12, minWidth: 150 },
          onMouseDown: function (e) { e.stopPropagation() },
          children: [
            h('div', { key: 'copy', onMouseDown: function (e) { e.stopPropagation(); copyCell(menu.r, menu.c) }, style: { padding: '5px 14px', cursor: 'pointer' } }, '复制'),
            h('div', { key: 'paste', onMouseDown: function (e) { e.stopPropagation(); pasteCell(menu.r, menu.c) }, style: { padding: '5px 14px', cursor: 'pointer' } }, '粘贴'),
            h('div', { key: 'clear', onMouseDown: function (e) { e.stopPropagation(); clearCell(menu.r, menu.c) }, style: { padding: '5px 14px', cursor: 'pointer' } }, '清除内容'),
            h('div', { key: 'd1', style: { height: 1, background: '#eee', margin: '4px 0' } }),
            h('div', { key: 'merge', onMouseDown: function (e) { e.stopPropagation(); if (range) mergeRange(range); else closeMenu() }, style: { padding: '5px 14px', cursor: 'pointer' } }, '合并单元格'),
            h('div', { key: 'unmerge', onMouseDown: function (e) { e.stopPropagation(); unmergeAt(menu.r, menu.c) }, style: { padding: '5px 14px', cursor: 'pointer' } }, '取消合并'),
            h('div', { key: 'd2', style: { height: 1, background: '#eee', margin: '4px 0' } }),
            h('div', { key: 'insrow', onMouseDown: function (e) { e.stopPropagation(); insertRow(menu.r) }, style: { padding: '5px 14px', cursor: 'pointer' } }, '在上方插入行'),
            h('div', { key: 'insrow2', onMouseDown: function (e) { e.stopPropagation(); insertRow(menu.r + 1) }, style: { padding: '5px 14px', cursor: 'pointer' } }, '在下方插入行'),
            h('div', { key: 'delrow', onMouseDown: function (e) { e.stopPropagation(); deleteRow(menu.r) }, style: { padding: '5px 14px', cursor: 'pointer', color: '#c00' } }, '删除行'),
            h('div', { key: 'inscol', onMouseDown: function (e) { e.stopPropagation(); insertCol(menu.c, menu.r) }, style: { padding: '5px 14px', cursor: 'pointer' } }, '在左侧插入列'),
            h('div', { key: 'inscol2', onMouseDown: function (e) { e.stopPropagation(); insertCol(menu.c + 1, menu.r) }, style: { padding: '5px 14px', cursor: 'pointer' } }, '在右侧插入列'),
            h('div', { key: 'delcol', onMouseDown: function (e) { e.stopPropagation(); deleteCol(menu.c, menu.r) }, style: { padding: '5px 14px', cursor: 'pointer', color: '#c00' } }, '删除列')
          ]
        }) : null
      )
    }


    function apply(ctx) {
      console.log('[dsh-excel-panel] v0.5.0 loaded')
      var registered = false
      function tryRegister() {
        var betterSidebar = ctx.betterSidebar || (ctx.get ? ctx.get('betterSidebar') : null)
        if (!betterSidebar || typeof betterSidebar.registerFileViewer !== 'function' || registered) return false
        try {
          var xlsxViewer = {
            id: 'excel-panel-xlsx-v5',
            title: function () { return '可编辑 Excel' },
            icon: function () { return null },
            exts: ['xlsx'],
            priority: 20,
            fetchStrategy: 'custom',
            load: function (path, scope) { return readExcel(scope, path) },
            component: function (props) { return h(XlsxEditor, props) }
          }
          if (ctx.effect) {
            ctx.effect(function () { return betterSidebar.registerFileViewer(xlsxViewer) }, 'dsh-excel-panel: xlsx viewer')
          } else {
            betterSidebar.registerFileViewer(xlsxViewer)
          }
          registered = true
          console.log('[dsh-excel-panel] xlsx viewer registered')
          return true
        } catch (e) {
          console.warn('[dsh-excel-panel] register xlsx viewer failed:', e)
          return false
        }
      }
      if (!tryRegister()) {
        var tries = 0
        var timer = setInterval(function () {
          tries += 1
          if (tryRegister() || tries > 40) clearInterval(timer)
        }, 500)
      }
    }

    exports.apply = apply
    exports.inject = inject
    return module.exports
  }
})

