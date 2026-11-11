// OCR 常驻服务客户端（2026-11-11）
// 用法：node tools/control/ocr-svc.mjs start | stop | status | <图片> [要找的字]
import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync, writeFileSync, unlinkSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { connect } from 'node:net'

const WS = process.cwd()
const PORT = 41983
const PIDFILE = join(WS, '数据', 'ocr-server.pid')
const SELF_DIR = dirname(fileURLToPath(import.meta.url))
const PS = existsSync(join(SELF_DIR, 'ps', 'ocr-server.ps1')) ? join(SELF_DIR, 'ps', 'ocr-server.ps1') : join(WS, 'tools', 'control', 'ps', 'ocr-server.ps1')
const TMP = join(WS, '数据', '操控截图')

function request(payload, timeoutMs = 30000) {
  return new Promise((resolve, reject) => {
    const s = connect(PORT, '127.0.0.1')
    let buf = ''
    const to = setTimeout(() => { s.destroy(); reject(new Error('timeout')) }, timeoutMs)
    s.on('connect', () => s.write(payload + '\n'))
    s.on('data', (d) => { buf += d.toString('utf8'); if (buf.includes('\n')) { clearTimeout(to); s.end(); resolve(buf.trim()) } })
    s.on('error', (e) => { clearTimeout(to); reject(e) })
  })
}

function sizeOf(img) {
  try {
    const out = execFileSync('ffmpeg', ['-hide_banner', '-i', img], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] })
    const m = /(\d{2,5})x(\d{2,5})/.exec(out); return m ? [Number(m[1]), Number(m[2])] : [0, 0]
  } catch (e) {
    const m = /(\d{2,5})x(\d{2,5})/.exec(String(e.stderr || '')); return m ? [Number(m[1]), Number(m[2])] : [0, 0]
  }
}

function parseLines(json) {
  const cut = json.lastIndexOf('}')
  const j = JSON.parse(json.slice(0, cut + 1))
  return j.ok ? (j.lines || []) : []
}

// 大图切块：2x2 带重叠，每块放大 2 倍再 OCR，坐标合并回原图
async function ocrTiled(img, w, h) {
  // 服务端单线程、请求会排队：块数越少越快。宽屏用左右两块（2x1）
  const cols = 2, rows = 2, overlap = 60
  const tw = Math.ceil(w / cols) + overlap, th = Math.ceil(h / rows) + overlap
  const tasks = []
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const x = Math.max(0, c * Math.ceil(w / cols) - overlap / 2)
      const y = Math.max(0, r * Math.ceil(h / rows) - overlap / 2)
      const cw = Math.min(tw, w - x), ch = Math.min(th, h - y)
      tasks.push((async () => {
        const up = join(TMP, `ocr-tile-${r}-${c}-${Date.now()}.png`)
        try {
          execFileSync('ffmpeg', ['-y', '-hide_banner', '-loglevel', 'error', '-i', img,
            '-vf', `crop=${cw}:${ch}:${x}:${y},scale=iw*2:ih*2:flags=lanczos`, up])
          const lines = parseLines(await request(`${up}\t\t2`))
          return lines.map((l) => ({ ...l, x: l.x + x, y: l.y + y }))
        } catch { return [] }
      })())
    }
  }
  const results = await Promise.all(tasks)
  return results.flat()
}
// 自愈：服务没跑就先拉起来，再发请求（2026-11-11）
async function ensureService() {
  try { await request('quit', 1500); return true } catch {}
  try {
    execFileSync(process.execPath, [process.argv[1], 'start'], { encoding: 'utf8', timeout: 60000, stdio: 'pipe' })
    return true
  } catch (e) {
    console.error('OCR 服务拉起失败:', String(e.stdout || e.message).slice(0, 120))
    return false
  }
}

async function ocrSingle(img, w, h) {
  let useImg = img, scale = '1'
  const up = join(TMP, `ocr-up-${Date.now()}.png`)
  // 像素数在上限内就整体放大 2 倍（小字准）；过大说明是整屏，交给切块
  if (w > 0 && h > 0 && w * h <= 2000000) {
    try { execFileSync('ffmpeg', ['-y', '-hide_banner', '-loglevel', 'error', '-i', img, '-vf', 'scale=iw*2:ih*2:flags=lanczos', up]); useImg = up; scale = '2' } catch {}
  }
  return parseLines(await request(`${useImg}\t\t${scale}`))
}

const [cmd, ...rest] = process.argv.slice(2)

if (cmd === 'start') {
  if (existsSync(PIDFILE)) {
    const pid = readFileSync(PIDFILE, 'utf8').trim()
    let alive = false
    if (pid) { try { process.kill(Number(pid), 0); alive = true } catch {} }
    if (alive) { console.log('已经在跑（pid ' + pid + '）'); process.exit(0) }
    try { unlinkSync(PIDFILE) } catch {}
  }
  const cmdLine = `Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','${PS}','-Port','${PORT}' -WindowStyle Hidden -PassThru | Select-Object -ExpandProperty Id`
  const pidOut = execFileSync('powershell', ['-NoProfile', '-Command', cmdLine], { encoding: 'utf8' }).trim()
  writeFileSync(PIDFILE, pidOut, 'utf8')
  for (let i = 0; i < 60; i++) {
    await new Promise((r) => setTimeout(r, 500))
    try { await request('quit', 2000); console.log('OCR 服务已就绪（pid ' + pidOut + '）'); process.exit(0) } catch {}
  }
  console.log('启动超时')
  process.exit(1)
} else if (cmd === 'stop') {
  try { await request('quit', 3000) } catch {}
  if (existsSync(PIDFILE)) {
    const pid = readFileSync(PIDFILE, 'utf8').trim()
    try { execFileSync('cmd', ['/c', 'taskkill', '/PID', pid, '/F'], { stdio: 'ignore' }) } catch {}
    try { unlinkSync(PIDFILE) } catch {}
  }
  console.log('已停')
} else if (cmd === 'status') {
  try { const r = await request('quit', 2000); console.log('服务在跑:', r) } catch (e) { console.log('没在跑:', e.message) }
} else if (cmd) {
  const img = cmd
  const find = rest[0] || ''
  const t0 = Date.now()
  if (!(await ensureService())) process.exit(1)
  const [w, h] = sizeOf(img)
  let lines = []
  try {
    lines = (w > 0 && h > 0 && w * h > 2000000) ? await ocrTiled(img, w, h) : await ocrSingle(img, w, h)
  } catch (e) {
    console.log('OCR 失败:', e.message)
    process.exit(1)
  }
  if (find) {
    const target = find.replace(/\s/g, '')
    const hits = lines.filter((l) => l.text.replace(/\s/g, '').includes(target))
    if (!hits.length) { console.log(JSON.stringify({ ok: false, error: 'text not found', count: lines.length })); process.exit(0) }
    const best = hits.sort((a, b) => Math.abs(a.text.replace(/\s/g, '').length - target.length) - Math.abs(b.text.replace(/\s/g, '').length - target.length))[0]
    const cx = Math.round(best.x + best.w / 2), cy = Math.round(best.y + best.h / 2)
    console.log(JSON.stringify({ ok: true, text: best.text, x: best.x, y: best.y, cx, cy, count: lines.length }))
  } else {
    console.log(JSON.stringify({ ok: true, count: lines.length, lines }))
  }
  console.log(`（${Date.now() - t0}ms）`)
} else {
  console.log('用法: start | stop | status | <图片> [要找的字]')
}
