// 操控模块 · MVP（2026-11-11）
// 目的：把「操作软件」变成有证据的动作——先选通道，再执行，最后验证。
// 用法：
//   node tools/control/control.mjs list
//   node tools/control/control.mjs probe word
//   node tools/control/control.mjs docx2pdf <in.docx> <out.pdf>
//   node tools/control/control.mjs xlsx2csv <in.xlsx> <out.csv>
//   node tools/control/control.mjs frame <video> <sec> <out.jpg>
//   node tools/control/control.mjs log
import { execFileSync, spawn } from 'node:child_process'
import { existsSync, statSync, appendFileSync, mkdirSync, readFileSync, writeFileSync, unlinkSync } from 'node:fs'
import { join, basename, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const WS = process.cwd()
const SELF_DIR = dirname(fileURLToPath(import.meta.url))
// 工具目录：装成包时是 <包>/bin、<包>/ps；在工作区里是 tools/control/bin
const BIN = existsSync(join(SELF_DIR, 'bin')) ? join(SELF_DIR, 'bin') : join(WS, 'tools', 'control', 'bin')
const PS = existsSync(join(SELF_DIR, 'ps')) ? join(SELF_DIR, 'ps') : join(WS, 'tools', 'control', 'ps')
const LOG = join(WS, '数据', '操控日志.jsonl')
mkdirSync(join(WS, '数据'), { recursive: true })

const now = () => new Date().toISOString()
const mb = (n) => (n / 1024 / 1024).toFixed(2) + 'MB'

// ── 通道梯度：越靠前越稳，UI 自动化放最后 ──
const REGISTRY = {
  word: {
    kind: 'com',
    app: 'Microsoft Word',
    probe: () => psOK('(New-Object -ComObject Word.Application).Version; (New-Object -ComObject Word.Application).Quit()'),
    actions: {
      docx2pdf: { desc: 'docx → pdf（Word COM，不弹窗）', minBytes: 5000, run: (a, b) => ps(`docx2pdf.ps1`, ['-In', a, '-Out', b]) },
    },
  },
  excel: {
    kind: 'com',
    app: 'Microsoft Excel',
    probe: () => psOK('(New-Object -ComObject Excel.Application).Version; (New-Object -ComObject Excel.Application).Quit()'),
    actions: {
      xlsx2csv: { desc: 'xlsx → csv（Excel COM，不弹窗）', minBytes: 1, run: (a, b) => ps('xlsx2csv.ps1', ['-In', a, '-Out', b]) },
    },
  },
  ffmpeg: {
    kind: 'cli',
    app: 'ffmpeg',
    probe: () => {
      const v = execFileSync('ffmpeg', ['-version'], { encoding: 'utf8' }).split('\n')[0]
      return { ok: true, evidence: v.slice(0, 60) }
    },
    actions: {
      frame: {
        desc: '抽一帧（视频 → 图片）',
        run: (video, sec, out) => {
          execFileSync('ffmpeg', ['-y', '-hide_banner', '-loglevel', 'error', '-ss', String(sec), '-i', video, '-frames:v', '1', out])
          return out
        },
      },
      probe: {
        desc: '读时长/分辨率',
        run: (video) => {
          const r = execFileSync('ffmpeg', ['-i', video], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] })
          return r
        },
      },
    },
  },
}

function ps(script, args) {
  const out = execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(PS, script), ...args], { encoding: 'utf8' })
  const s = String(out).trim()
  if (/^ERR/.test(s)) throw new Error(s)
  return s
}
function psOK(cmd) {
  const out = execFileSync('powershell', ['-NoProfile', '-Command', cmd], { encoding: 'utf8' })
  return { ok: true, evidence: String(out).trim().slice(0, 40) }
}

function log(entry) {
  appendFileSync(LOG, JSON.stringify(entry) + '\n', 'utf8')
}

function verifyFile(path, { minBytes = 1000 } = {}) {
  if (!existsSync(path)) return { ok: false, why: '输出不存在' }
  const st = statSync(path)
  if (st.size < minBytes) return { ok: false, why: `输出太小 ${st.size}B` }
  return { ok: true, evidence: `${basename(path)} ${mb(st.size)}` }
}


// ── UIA 优先（2026-11-11）：原生窗口读控件树，按名字精确找；读不到再退 OCR ──
function uiaFind(pid, text) {
  const exe = join(BIN, 'uia.exe')
  if (!existsSync(exe)) return { items: [], truncated: false }
  try {
    const out = execFileSync(exe, ['find', String(pid), text, '6', '400'], { encoding: 'utf8', timeout: 20000 }).trim()
    const line = out.split('\n').filter(Boolean).pop() || ''
    const j = JSON.parse(line.slice(0, line.lastIndexOf('}') + 1))
    return j.ok ? { items: j.items || [], truncated: !!j.truncated } : { items: [], truncated: false }
  } catch { return { items: [], truncated: false } }
}

// 打分：完全相等 > 开头匹配 > 包含；同级取名字最短、面积最小（更像"那个按钮"）
function scoreItem(it, text) {
  const n = it.name.replace(/\s/g, '')
  const t = text.replace(/\s/g, '')
  let s = 0
  if (n === t) s += 100
  else if (n.startsWith(t)) s += 60
  else if (n.includes(t)) s += 30
  s -= Math.min(20, Math.abs(n.length - t.length))
  s -= Math.min(10, (it.w * it.h) / 20000)
  return s
}

// ── 窗口作用域守卫（2026-11-11）：锁定窗口后，落在窗口外的坐标一律拒绝 ──
function lockGuard(points) {
  const LOCK = join(WS, '数据', 'window-lock.json')
  if (!existsSync(LOCK)) return { ok: true }
  let lock
  try { lock = JSON.parse(readFileSync(LOCK, 'utf8')) } catch { return { ok: true } }
  const exe = join(BIN, 'window.exe')
  let rect
  try { rect = JSON.parse(execFileSync(exe, ['rectpid', String(lock.pid)], { encoding: 'utf8' })) } catch { rect = { ok: false } }
  if (!rect.ok) return { ok: false, reason: `锁定的窗口「${lock.title}」已不存在` }
  for (const [x, y] of points) {
    if (x < rect.x || y < rect.y || x > rect.x + rect.w || y > rect.y + rect.h) {
      return { ok: false, reason: `(${x},${y}) 不在锁定的窗口「${lock.title}」[${rect.x},${rect.y} ${rect.w}x${rect.h}] 内` }
    }
  }
  return { ok: true, rect, title: lock.title }
}
function run(app, action, args) {
  const t0 = Date.now()
  const a = REGISTRY[app]?.actions?.[action]
  if (!a) {
    console.error(`没有这个动作：${app}.${action}`)
    process.exit(2)
  }
  const outPath = args[args.length - 1]
  const entry = { ts: now(), app, action, args, ok: false }
  try {
    const r = a.run(...args)
    const v = verifyFile(outPath, { minBytes: a.minBytes ?? 1000 })
    entry.ok = v.ok
    entry.evidence = v.ok ? v.evidence : v.why
    if (!v.ok) throw new Error('验证失败：' + v.why)
    console.log(`OK ${app}.${action} → ${v.evidence}（${Date.now() - t0}ms）`)
  } catch (e) {
    entry.error = String(e.message).slice(0, 200)
    console.log(`FAIL ${app}.${action}: ${entry.error}`)
    process.exitCode = 1
  } finally {
    entry.ms = Date.now() - t0
    log(entry)
  }
}

const [cmd, ...rest] = process.argv.slice(2)
if (cmd === 'list') {
  for (const [app, v] of Object.entries(REGISTRY)) {
    console.log(`${app.padEnd(8)} ${v.kind.padEnd(5)} ${v.app}`)
    for (const [act, d] of Object.entries(v.actions)) console.log(`   ${act.padEnd(10)} ${d.desc}`)
  }
} else if (cmd === 'probe') {
  const app = rest[0]
  const v = REGISTRY[app]
  if (!v) { console.log('没有这个 app'); process.exit(2) }
  try {
    const r = v.probe()
    console.log(`${app}: ${r.ok ? '可用' : '不可用'} | ${r.evidence}`)
  } catch (e) {
    console.log(`${app}: 不可用 | ${String(e.message).slice(0, 80)}`)
  }
} else if (cmd === 'log') {
  const lines = existsSync(LOG) ? readFileSync(LOG, 'utf8').split('\n').filter(Boolean).slice(-10) : []
  for (const l of lines) console.log(l)
  if (!lines.length) console.log('（还没有日志）')
} else if (cmd === 'click' || cmd === 'drag' || cmd === 'move') {
  // 鼠标类：click x y / drag x1 y1 x2 y2 / move x y
  // 选项：--visible 截图留证、--steps N 插值步数、--dry 只打印计划
  const nums = rest.filter((a) => /^-?\d+$/.test(a)).map(Number)
  const visible = rest.includes('--visible')
  const dry = rest.includes('--dry')
  const si = rest.indexOf('--steps')
  const steps = si >= 0 ? Number(rest[si + 1]) : 20
  if (!dry) {
    const pts = cmd === 'drag' ? [[nums[0], nums[1]], [nums[2], nums[3]]] : [[nums[0], nums[1]]]
    if (pts.some(([x, y]) => !Number.isFinite(x) || !Number.isFinite(y))) { console.log('坐标不完整'); process.exit(2) }
    const g = lockGuard(pts)
    if (!g.ok) {
      console.log('拒绝执行（窗口锁）：' + g.reason)
      log({ ts: now(), app: 'mouse', action: cmd, args: nums, ok: false, error: 'window-lock: ' + g.reason })
      process.exit(1)
    }
  }  const shotDir = join(WS, '数据', '操控截图', `${cmd}-${Date.now()}`)
  const exe = join(BIN, 'mouse.exe')
  const t0 = Date.now()
  const entry = { ts: now(), app: 'mouse', action: cmd, args: nums, ok: false }
  try {
    let out
    if (existsSync(exe)) {
      // 编译版：快（~1s），优先用
      const exeArgs = [cmd, ...nums.map(String), '--steps', String(steps)]
      if (visible && !dry) exeArgs.push('--shot', shotDir)
      if (visible && !dry && cmd === 'drag') exeArgs.push('--mid')
      if (dry) { exeArgs.length = 0; exeArgs.push(cmd, ...nums.map(String), '--steps', String(steps)) }
      out = execFileSync(exe, exeArgs, { encoding: 'utf8' }).trim()
    } else {
      const psArgs = ['-Action', cmd, '-Steps', String(steps), '-X1', String(nums[0] ?? 0), '-Y1', String(nums[1] ?? 0)]
      if (cmd === 'drag') { psArgs.push('-X2', String(nums[2] ?? 0), '-Y2', String(nums[3] ?? 0)) }
      if (visible && !dry) psArgs.push('-ShotDir', shotDir)
      if (visible && !dry && cmd === 'drag') psArgs.push('-MidShot')
      if (dry) psArgs.push('-DryRun')
      out = execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(PS, 'mouse.ps1'), ...psArgs], { encoding: 'utf8' }).trim()
    }
    entry.ok = /(^|\n)(OK|DRY)/.test(out)
    entry.evidence = out
    if (visible && !dry) {
      entry.shotDir = shotDir
      entry.changed = /changed=(true|True)/.test(out)
    }
    console.log(out, `（${Date.now() - t0}ms）`)
  } catch (e) {
    entry.error = String(e.message).slice(0, 200)
    console.log('FAIL', cmd, entry.error)
    process.exitCode = 1
  } finally {
    entry.ms = Date.now() - t0
    log(entry)
  }
} else if (cmd === 'ocr' || cmd === 'findtext' || cmd === 'clicktext') {
  const img = rest[0]
  const text = cmd === 'ocr' ? '' : rest[1]
  const t0 = Date.now()
  // 小字识别差：先用 ffmpeg 放大 2 倍再 OCR（PS 里没有 System.Drawing，交给 ffmpeg）
  const up = join(WS, '数据', '操控截图', `ocr-up-${Date.now()}.png`)
  let useImg = img
  let scale = 1
  try {
    execFileSync('ffmpeg', ['-y', '-hide_banner', '-loglevel', 'error', '-i', img, '-vf', 'scale=iw*2:ih*2:flags=lanczos', up])
    useImg = up
    scale = 2
  } catch { /* 放大失败就用原图 */ }
  const psArgs = ['-Image', useImg, '-Scale', String(scale)]
  if (text) psArgs.push('-Find', text)
  let out = ''
  let served = false
  // 常驻 OCR 服务：没跑就自动拉起（2026-11-11），服务端空闲 10 分钟自己退
  try {
    const svcStart = join(WS, 'tools', 'control', 'ocr-svc.mjs')
    execFileSync(process.execPath, [svcStart, 'start'], { encoding: 'utf8', timeout: 40000 })
  } catch { /* 起不来就走一次性 */ }
  // 常驻 OCR 服务在跑就直接问它（~0.9s）；不在就跑一次性（~7s）
  if (existsSync(join(WS, '数据', 'ocr-server.pid'))) {
    try {
      const svc = join(WS, 'tools', 'control', 'ocr-svc.mjs')
      const r = execFileSync(process.execPath, [svc, img, text || ''], { encoding: 'utf8', timeout: 30000 }).trim()
      const line = r.split('\n')[0]
      const cut = line.lastIndexOf('}')
      out = JSON.parse(line.slice(0, cut + 1)) ? line.slice(0, cut + 1) : ''
      served = true
    } catch { served = false }
  }
  try {
    if (!served) out = execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(PS, 'ocr.ps1'), ...psArgs], { encoding: 'utf8' }).trim()
    const j = JSON.parse(out)
    if (cmd === 'ocr') {
      console.log(`OCR ${j.count} 行（${Date.now() - t0}ms）`)
      for (const l of (j.lines || []).slice(0, 12)) console.log(`  (${l.x},${l.y}) ${l.text}`)
    } else if (!j.ok) {
      console.log('没找到文字:', j.error)
      process.exitCode = 1
    } else if (cmd === 'findtext') {
      console.log(`找到「${j.text}」→ 中心 (${j.cx},${j.cy})（${Date.now() - t0}ms）`)
    } else {
      const exe = join(BIN, 'mouse.exe')
      const shotDir = join(WS, '数据', '操控截图', `clicktext-${Date.now()}`)
      const exeArgs = ['click', String(j.cx), String(j.cy)]
      if (rest.includes('--visible')) exeArgs.push('--shot', shotDir)
      const out2 = execFileSync(exe, exeArgs, { encoding: 'utf8' }).trim()
      console.log(`点「${j.text}」(${j.cx},${j.cy}) → ${out2}（${Date.now() - t0}ms）`)
    }
    log({ ts: now(), app: 'ocr', action: cmd, args: rest, ok: true, evidence: out.slice(0, 200), ms: Date.now() - t0 })
  } catch (e) {
    console.log('OCR 失败:', String(e.message).slice(0, 160))
    log({ ts: now(), app: 'ocr', action: cmd, args: rest, ok: false, error: String(e.message).slice(0, 160), ms: Date.now() - t0 })
    process.exitCode = 1
  }
} else if (cmd === 'diffimg') {
  const [a, b, x, y, w, h] = rest
  const args = [a, b]
  if (x !== undefined) args.push(x, y, w, h)
  const out = execFileSync(join(BIN, 'imgdiff.exe'), args, { encoding: 'utf8' }).trim()
  console.log(out)
} else if (cmd === 'topdf') {
  // LibreOffice 批量转 PDF：topdf <outdir> <file...>
  const [outdir, ...rawFiles] = rest
  if (!outdir || !rawFiles.length) { console.log('用法: topdf <outdir> <file...>'); process.exit(2) }
  const SOFFICE = 'C:/Program Files/LibreOffice/program/soffice.exe'
  const t0 = Date.now()
  // 2026-11-11 加的守卫：①输入必须存在 ②同名文件会互相覆盖（sample.docx 和 sample.xlsx 都出 sample.pdf）
  const missing = rawFiles.filter((f) => !existsSync(f))
  const files = rawFiles.filter((f) => existsSync(f))
  const byBase = {}
  for (const f of files) { const b = basename(f).replace(/\.[^.]+$/, '').toLowerCase(); byBase[b] = (byBase[b] || 0) + 1 }
  const dupBases = Object.keys(byBase).filter((k) => byBase[k] > 1)
  if (missing.length) console.log('缺少输入（跳过）:', missing.join(', '))
  if (!files.length) { console.log('FAIL topdf: 没有可转的文件'); process.exitCode = 1 }
  else {
    try {
      const plan = files.map((f) => {
        const base = basename(f).replace(/\.[^.]+$/, '')
        const ext = basename(f).split('.').pop()
        const collides = byBase[base.toLowerCase()] > 1
        return { f, base, ext, outName: collides ? `${base}-${ext}.pdf` : `${base}.pdf` }
      })
      if (dupBases.length) {
        console.log('同名文件，逐个转换并改名为 base-ext.pdf:', dupBases.join(', '))
        for (const p of plan) {
          const tmpdir = join(outdir, `_tmp_${p.base}_${p.ext}`)
          execFileSync(SOFFICE, ['--headless', '--norestore', '-env:UserInstallation=file:///D:/tmp/lo-profile',
            '--convert-to', 'pdf', '--outdir', tmpdir, p.f], { encoding: 'utf8', timeout: 600000 })
          const produced = join(tmpdir, `${p.base}.pdf`)
          if (existsSync(produced)) execFileSync('cmd', ['/c', 'move', '/Y', produced, join(outdir, p.outName)])
        }
      } else {
        execFileSync(SOFFICE, ['--headless', '--norestore', '-env:UserInstallation=file:///D:/tmp/lo-profile',
          '--convert-to', 'pdf', '--outdir', outdir, ...files], { encoding: 'utf8', timeout: 600000 })
      }
      const results = plan.map((p) => {
        const pdf = join(outdir, p.outName)
        const v = verifyFile(pdf, { minBytes: 1000 })
        return { file: basename(p.f), pdf: p.outName, ok: v.ok, evidence: v.ok ? v.evidence : v.why }
      })
      const okN = results.filter((r) => r.ok).length
      for (const r of results) console.log(`  ${r.ok ? 'OK ' : 'FAIL'} ${r.file} → ${r.pdf} ${r.evidence}`)
      const ok = okN === rawFiles.length
      console.log(`${ok ? 'OK' : 'PARTIAL'} topdf ${okN}/${rawFiles.length}（${Date.now() - t0}ms）`)
      log({ ts: now(), app: 'office', action: 'topdf', args: rawFiles, outdir, ok, evidence: `${okN}/${rawFiles.length}`, ms: Date.now() - t0 })
      if (!ok) process.exitCode = 1
    } catch (e) {
      console.log('FAIL topdf:', String(e.message).slice(0, 160))
      log({ ts: now(), app: 'office', action: 'topdf', args: rawFiles, ok: false, error: String(e.message).slice(0, 160), ms: Date.now() - t0 })
      process.exitCode = 1
    }
  }
} else if (cmd === 'window') {
  // 窗口作用域：list / lock <关键词> / unlock / status（2026-11-11）
  const LOCK = join(WS, '数据', 'window-lock.json')
  const winExe = join(BIN, 'window.exe')
  const winList = () => JSON.parse(execFileSync(winExe, ['list'], { encoding: 'utf8' }))
  const winRect = (pid) => { try { return JSON.parse(execFileSync(winExe, ['rectpid', String(pid)], { encoding: 'utf8' })) } catch { return { ok: false } } }
  const sub = rest[0]
  if (sub === 'list') {
    const j = winList()
    console.log(`窗口 ${j.count} 个：`)
    for (const w of j.windows.slice(0, 15)) console.log(`  [${w.pid}] ${w.x},${w.y} ${w.w}x${w.h}  ${w.title.slice(0, 40)}`)
  } else if (sub === 'lock') {
    const kw = rest[1] || ''
    const j = winList()
    const hit = j.windows.find((w) => w.title.toLowerCase().includes(kw.toLowerCase()))
    if (!hit) { console.log('没有匹配的窗口:', kw); process.exitCode = 1 }
    else if (hit.x <= -30000 || hit.y <= -30000) {
      console.log(`「${hit.title}」是最小化窗口（${hit.x},${hit.y} ${hit.w}x${hit.h}），先把它恢复出来再锁`)
      process.exitCode = 1
    } else {
      writeFileSync(LOCK, JSON.stringify({ title: hit.title, pid: hit.pid, x: hit.x, y: hit.y, w: hit.w, h: hit.h, ts: now() }, null, 1), 'utf8')
      console.log(`已锁定「${hit.title}」 pid=${hit.pid} 区域 ${hit.x},${hit.y} ${hit.w}x${hit.h}`)
    }
  } else if (sub === 'unlock') {
    try { unlinkSync(LOCK) } catch {}
    console.log('已解除窗口锁')
  } else if (sub === 'status') {
    if (!existsSync(LOCK)) console.log('未锁定（全屏可操作）')
    else {
      const l = JSON.parse(readFileSync(LOCK, 'utf8'))
      const r = winRect(l.pid)
      console.log(r.ok ? `锁定中：「${l.title}」 pid=${l.pid} 当前区域 ${r.x},${r.y} ${r.w}x${r.h}` : `锁定中：「${l.title}」 pid=${l.pid}（窗口已不在）`)
    }
  } else console.log('用法: window list | lock <关键词> | unlock | status')
} else if (cmd === 'grid') {
  // 屏幕点阵：cell 64px，列 A..Z,AA.. 行 1..N，锚点=虚拟屏幕左上角
  const CELL = 64
  const gridInfo = () => {
    const cache = join(WS, '数据', 'grid.json')
    if (existsSync(cache)) return JSON.parse(readFileSync(cache, 'utf8'))
    const out = execFileSync(join(BIN, 'grid.exe'), ['info'], { encoding: 'utf8' }).trim()
    const m = out.match(/virtual=(-?\d+),(-?\d+),(\d+)x(\d+)/)
    const info = { x: Number(m[1]), y: Number(m[2]), w: Number(m[3]), h: Number(m[4]), cell: CELL }
    writeFileSync(cache, JSON.stringify(info), 'utf8')
    return info
  }
  const colName = (i) => { let s = ''; i++; while (i > 0) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26) } return s }
  const parseCell = (id) => {
    // 支持亚格偏移：M8、M8+12-6（相对格心）、M8@10,20（相对格左上角）
    const s = String(id).toUpperCase().trim()
    const at = s.match(/^([A-Z]+\d+)@(-?\d+),(-?\d+)$/)
    const off = s.match(/^([A-Z]+\d+)([+-]\d+)([+-]\d+)$/)
    const plain = s.match(/^([A-Z]+\d+)$/)
    const core = (at && at[1]) || (off && off[1]) || (plain && plain[1])
    if (!core) return null
    const m = core.match(/^([A-Z]+)(\d+)$/)
    let col = 0
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
    return {
      col: col - 1, row: Number(m[2]),
      dx: at ? Number(at[2]) : (off ? Number(off[2]) : 0),
      dy: at ? Number(at[3]) : (off ? Number(off[3]) : 0),
      mode: at ? 'topleft' : (off ? 'center' : 'center'),
    }
  }
  const cellRect = (info, col, row) => {
    const x = info.x + col * CELL, y = info.y + (row - 1) * CELL
    return { x, y, w: CELL, h: CELL, cx: x + CELL / 2, cy: y + CELL / 2 }
  }
  // 把 M8 / M8+12-6 / M8@10,20 解析成绝对坐标
  const resolveCell = (id) => {
    const c = parseCell(id)
    if (!c) return null
    const r = cellRect(gridInfo(), c.col, c.row)
    const bx = c.mode === 'topleft' ? r.x : r.cx
    const by = c.mode === 'topleft' ? r.y : r.cy
    return { rect: r, x: bx + c.dx, y: by + c.dy, col: c.col, row: c.row }
  }
  const sub = rest[0]
  if (sub === 'info') {
    const i = gridInfo()
    console.log(`虚拟屏幕 ${i.x},${i.y} ${i.w}x${i.h} | 格子 ${i.cell}px | 列数 ${Math.ceil(i.w / i.cell)} 行数 ${Math.ceil(i.h / i.cell)}`)
  } else if (sub === 'cell') {
    const r = resolveCell(rest[1] || '')
    if (!r) { console.log('格式要像 B7 / B7+12-6 / B7@10,20'); process.exitCode = 2 }
    else { console.log(`${String(rest[1]).toUpperCase()} → 格 x=${r.rect.x} y=${r.rect.y} ${r.rect.w}x${r.rect.h} | 落点=(${r.x},${r.y})`) }
  } else if (sub === 'at') {
    const i = gridInfo(); const x = Number(rest[1]), y = Number(rest[2])
    const col = Math.floor((x - i.x) / CELL), row = Math.floor((y - i.y) / CELL) + 1
    const r = cellRect(i, col, row)
    console.log(`(${x},${y}) → ${colName(col)}${row} 格子起点=(${r.x},${r.y})`)
  } else if (sub === 'overlay') {
    if (rest[1] === 'on') {
      const p = spawn(join(BIN, 'grid.exe'), ['overlay', 'on'], { detached: true, stdio: 'ignore' })
      p.unref()
      console.log(`覆盖层已启动（pid ${p.pid}）；关掉：taskkill /PID ${p.pid} /F`)
    } else {
      try { execFileSync('cmd', ['/c', 'taskkill', '/IM', 'grid.exe', '/F'], { stdio: 'ignore' }) } catch {}
      console.log('覆盖层已关闭')
    }
  } else {
    console.log('用法: grid info | grid cell B7 | grid at 800 500 | grid overlay on|off')
  }
} else if (cmd === 'dragcell') {
  // 按格子拖动：dragcell A1 B5 [--visible] [--steps N]
  const CELL = 64
  const info = (() => {
    const cache = join(WS, '数据', 'grid.json')
    if (existsSync(cache)) return JSON.parse(readFileSync(cache, 'utf8'))
    const out = execFileSync(join(BIN, 'grid.exe'), ['info'], { encoding: 'utf8' }).trim()
    const m = out.match(/virtual=(-?\d+),(-?\d+),(\d+)x(\d+)/)
    const i = { x: Number(m[1]), y: Number(m[2]), w: Number(m[3]), h: Number(m[4]), cell: CELL }
    writeFileSync(cache, JSON.stringify(i), 'utf8')
    return i
  })()
  const parse = (id) => {
    const m = String(id).toUpperCase().match(/^([A-Z]+)(\d+)$/)
    if (!m) return null
    let col = 0
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
    return { cx: info.x + (col - 1) * CELL + CELL / 2, cy: info.y + (Number(m[2]) - 1) * CELL + CELL / 2 }
  }
  const a = parse(rest[0] || ''), b = parse(rest[1] || '')
  if (!a || !b) { console.log('用法: dragcell A1 B5'); process.exitCode = 2 }
  else {
    const exe = join(BIN, 'mouse.exe')
    const shotDir = join(WS, '数据', '操控截图', `dragcell-${Date.now()}`)
    const args = ['drag', String(Math.round(a.cx)), String(Math.round(a.cy)), String(Math.round(b.cx)), String(Math.round(b.cy)), '--steps', '20']
    if (rest.includes('--visible')) args.push('--shot', shotDir, '--mid')
    const g2 = lockGuard([[Math.round(a.cx), Math.round(a.cy)], [Math.round(b.cx), Math.round(b.cy)]])
    if (!g2.ok) { console.log('拒绝执行（窗口锁）：' + g2.reason); process.exit(1) }
    const t0 = Date.now()
    const out = execFileSync(exe, args, { encoding: 'utf8' }).trim()
    console.log(out, `（${Date.now() - t0}ms）`)
    log({ ts: now(), app: 'grid', action: 'dragcell', args: rest, ok: /OK/.test(out), evidence: out.slice(0, 160), ms: Date.now() - t0 })
  }
} else if (cmd === 'type') {
  // 中文输入：剪贴板 + Ctrl+V（SendKeys 直接发中文会丢）
  // 用法：type "文字" [--at x y] [--verify x y w h]
  const text = rest.find((a) => !a.startsWith('--') && !/^-?\d+$/.test(a)) || ''
  const atIdx = rest.indexOf('--at')
  const verIdx = rest.indexOf('--verify')
  const t0 = Date.now()
  try {
    if (atIdx >= 0) {
      const x = Number(rest[atIdx + 1]), y = Number(rest[atIdx + 2])
      const g3 = lockGuard([[x, y]])
      if (!g3.ok) { console.log('拒绝执行（窗口锁）：' + g3.reason); process.exit(1) }
      execFileSync(join(BIN, 'mouse.exe'), ['click', String(x), String(y)], { encoding: 'utf8' })
      await new Promise((r) => setTimeout(r, 200))
    }
    const out = execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(PS, 'type.ps1'), '-Text', text], { encoding: 'utf8' }).trim()
    console.log(out)
    let verified = null
    if (verIdx >= 0) {
      // 读回：截图该区域并 OCR，确认文字真的进去了
      const [vx, vy, vw, vh] = rest.slice(verIdx + 1, verIdx + 5).map(Number)
      const shot = join(WS, '数据', '操控截图', `type-verify-${Date.now()}.png`)
      execFileSync(join(BIN, 'mouse.exe'), ['move', String(vx + Math.round(vw / 2)), String(vy + Math.round(vh / 2))])
      execFileSync('powershell', ['-NoProfile', '-Command', `Add-Type -AssemblyName System.Drawing; Add-Type -AssemblyName System.Windows.Forms; $b=[System.Windows.Forms.SystemInformation]::VirtualScreen; $bmp=New-Object System.Drawing.Bitmap $b.Width,$b.Height; $g=[System.Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen($b.Left,$b.Top,0,0,$bmp.Size); $bmp.Save('${shot.replace(/'/g, "''")}'); $g.Dispose(); $bmp.Dispose()`], { encoding: 'utf8' })
      const crop = shot.replace(/\.png$/, '-crop.png')
      execFileSync('ffmpeg', ['-y', '-hide_banner', '-loglevel', 'error', '-i', shot, '-vf', `crop=${vw}:${vh}:${vx}:${vy}`, crop])
      const r = execFileSync(process.execPath, [join(WS, 'tools', 'control', 'ocr-svc.mjs'), crop, ''], { encoding: 'utf8' }).trim()
      const line = r.split('\n')[0]
      const j = JSON.parse(line.slice(0, line.lastIndexOf('}') + 1))
      verified = j.ok && (j.lines || []).some((l) => l.text.replace(/\s/g, '').includes(text.replace(/\s/g, '').slice(0, 6)))
      console.log(`读回${verified ? '成功' : '没看到文字'}：` + (j.lines || []).map((l) => l.text).join(' | ').slice(0, 80))
    }
    log({ ts: now(), app: 'mouse', action: 'type', args: [text.slice(0, 40)], ok: true, evidence: out, verified, ms: Date.now() - t0 })
  } catch (e) {
    console.log('FAIL type:', String(e.message).slice(0, 160))
    log({ ts: now(), app: 'mouse', action: 'type', args: [text.slice(0, 40)], ok: false, error: String(e.message).slice(0, 160), ms: Date.now() - t0 })
    process.exitCode = 1
  }

} else if (cmd === 'anchor') {
  // 命名锚点：anchor set <名字> <格子如 Q9 | 坐标 800,600> / list / get <名字>
  const P = join(WS, '数据', '锚点.json')
  const load = () => (existsSync(P) ? JSON.parse(readFileSync(P, 'utf8')) : {})
  const save = (o) => writeFileSync(P, JSON.stringify(o, null, 1), 'utf8')
  const sub = rest[0]
  if (sub === 'set') {
    const name = rest[1], val = rest[2] || ''
    if (!name || !val) { console.log('用法: anchor set <名字> <格子如Q9 | 坐标800,600>'); process.exit(2) }
    let rec
    const m = val.match(/^([A-Za-z]+)(\d+)(?:([+-]\d+)([+-]\d+))?$/)
    if (m) {
      const info = JSON.parse(readFileSync(join(WS, '数据', 'grid.json'), 'utf8'))
      let col = 0; for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
      const dx = m[3] ? Number(m[3]) : 0, dy = m[4] ? Number(m[4]) : 0
      rec = { cell: val.toUpperCase(), x: info.x + (col - 1) * 64 + 32 + dx, y: info.y + (Number(m[2]) - 1) * 64 + 32 + dy, ts: now() }
    } else {
      const mm = val.match(/^(-?\d+),(-?\d+)$/)
      if (!mm) { console.log('坐标要写成 x,y'); process.exit(2) }
      rec = { x: Number(mm[1]), y: Number(mm[2]), ts: now() }
    }
    const o = load(); o[name] = rec; save(o)
    console.log(`锚点「${name}」→ (${rec.x},${rec.y})` + (rec.cell ? ` 格子 ${rec.cell}` : ''))
  } else if (sub === 'list') {
    const o = load(); const ks = Object.keys(o)
    if (!ks.length) console.log('（还没有锚点）')
    else for (const k of ks) console.log(`  ${k.padEnd(12)} (${o[k].x},${o[k].y})` + (o[k].cell ? `  ${o[k].cell}` : ''))
  } else if (sub === 'get') {
    const o = load(); const r = o[rest[1]]
    console.log(r ? `(${r.x},${r.y})` : '没有这个锚点')
  } else console.log('用法: anchor set <名字> <格子|坐标> | list | get <名字>')
} else if (cmd === 'clickname') {
  // 找元素并点击（2026-11-11 改版）
  // 顺序：UIA 控件树（快、准）→ 不行再截图 OCR（慢、会认错字）
  // 有同名多个时列出候选，用 --index N 选第几个
  const args = rest.filter((a) => !a.startsWith('--'))
  const text = args[0] || ''
  if (!text) { console.log('用法: clickname "某段文字" [--index N]'); process.exit(2) }
  const idxAt = rest.indexOf('--index')
  const wantIdx = idxAt >= 0 ? Number(rest[idxAt + 1]) : 0
  const lockFile = join(WS, '数据', 'window-lock.json')
  const lockPid = existsSync(lockFile) ? (JSON.parse(readFileSync(lockFile, 'utf8')).pid) : null

  // ① UIA（有锁用锁的 pid，没锁用前台窗口）
  {
    const u = uiaFind(lockPid || 'fg', text)
    const items = u.items.sort((a, b) => scoreItem(b, text) - scoreItem(a, text))
    if (items.length) {
      if (items.length > 1) {
        console.log(`UIA 命中 ${items.length} 个，候选：`)
        items.slice(0, 8).forEach((it, n) => console.log(`  ${n + 1}. 「${it.name}」${it.type} (${it.x},${it.y} ${it.w}x${it.h})`))
      }
      const pick = items[wantIdx > 0 ? Math.min(wantIdx - 1, items.length - 1) : 0]
      const cx = Math.round(pick.x + pick.w / 2), cy = Math.round(pick.y + pick.h / 2)
      const g = lockGuard([[cx, cy]])
      if (!g.ok) { console.log('拒绝执行（窗口锁）：' + g.reason); process.exit(1) }
      const exe = join(BIN, 'mouse.exe')
      const out = execFileSync(exe, ['click', String(cx), String(cy), '--shot', join(WS, '数据', '操控截图', `clickname-${Date.now()}`)], { encoding: 'utf8' }).trim()
      console.log(`[UIA] 「${pick.name}」${pick.type} → 点 (${cx},${cy}) → ${out}`)
      log({ ts: now(), app: 'uia', action: 'clickname', args: [text], ok: true, evidence: out.slice(0, 160), candidates: items.length })
      process.exit(0)
    }
    console.log('[UIA] 没读到控件' + (u.truncated ? '（扫描被截断，可能没扫全）' : '') + '，退 OCR')
  }

  const shot = join(WS, '数据', '操控截图', `clickname-${Date.now()}.png`)
  mkdirSync(join(WS, '数据', '操控截图'), { recursive: true })
  let ox = 0, oy = 0
  let region = null
  if (existsSync(lockFile)) {
    try {
      const l = JSON.parse(readFileSync(lockFile, 'utf8'))
      const r = JSON.parse(execFileSync(join(BIN, 'window.exe'), ['rectpid', String(l.pid)], { encoding: 'utf8' }))
      if (r.ok && r.x > -30000 && r.y > -30000 && r.w > 200 && r.h > 150) { region = r; ox = r.x; oy = r.y }
    } catch {}
  }
  const shotArgs = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(PS, 'shot.ps1'), '-Out', shot]
  if (region) shotArgs.push('-X', String(region.x), '-Y', String(region.y), '-W', String(region.w), '-H', String(region.h))
  execFileSync('powershell', shotArgs, { encoding: 'utf8' })
  const svc = join(WS, 'tools', 'control', 'ocr-svc.mjs')
  const raw = execFileSync(process.execPath, [svc, shot, text], { encoding: 'utf8', timeout: 60000 }).trim()
  const line = raw.split('\n')[0]
  const j = JSON.parse(line.slice(0, line.lastIndexOf('}') + 1))
  if (!j.ok) { console.log('OCR 没找到「' + text + '」'); process.exit(1) }
  const cx = j.cx + ox, cy = j.cy + oy
  const g = lockGuard([[cx, cy]])
  if (!g.ok) { console.log('拒绝执行（窗口锁）：' + g.reason); process.exit(1) }
  const exe = join(BIN, 'mouse.exe')
  const out = execFileSync(exe, ['click', String(cx), String(cy), '--shot', join(WS, '数据', '操控截图', `clickname-${Date.now()}`)], { encoding: 'utf8' }).trim()
  console.log(`[OCR] 找到「${j.text}」→ 点 (${cx},${cy})` + (region ? `（只截窗口 ${region.w}x${region.h}）` : '（整屏）') + ` → ${out}`)
  log({ ts: now(), app: 'ocr', action: 'clickname', args: [text], ok: true, evidence: out.slice(0, 160), region: region ? `${region.w}x${region.h}` : 'full' })
} else if (cmd === 'uia') {
  // uia list|find <pid> [text] —— 直接看某个窗口的控件树
  const sub = rest[0] || 'list'
  const pid = Number(rest[1] || 0)
  if (!pid) { console.log('用法: uia list|find <pid> [文字]'); process.exit(2) }
  const exe = join(BIN, 'uia.exe')
  const out = execFileSync(exe, [sub, String(pid), rest[2] || '', '6', '400'], { encoding: 'utf8', timeout: 25000 }).trim()
  const line = out.split('\n').filter(Boolean).pop() || ''
  const j = JSON.parse(line.slice(0, line.lastIndexOf('}') + 1))
  if (!j.ok) { console.log('UIA 失败:', j.error); process.exit(1) }
  console.log(`扫描 ${j.scanned} 个节点，命中 ${j.hits}`)
  for (const it of j.items.slice(0, 30)) console.log(`  「${it.name}」${it.type} (${it.x},${it.y} ${it.w}x${it.h})`)
} else if (cmd) {
  // 直接动作：docx2pdf <in> <out>
  const APP = { docx2pdf: 'word', xlsx2csv: 'excel', frame: 'ffmpeg' }[cmd]
  if (!APP) { console.error('未知命令:', cmd); process.exit(2) }
  run(APP, cmd, rest)
} else {
  console.log('用法: list | probe <app> | docx2pdf <in> <out> | xlsx2csv <in> <out> | frame <video> <sec> <out> | log')
}
