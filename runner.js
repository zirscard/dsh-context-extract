/**
 * Runtime bridge for dsh-context-extract.
 *
 * Locates the DSH-bundled Python runtime and runs `core/extract.py`, returning
 * its normalized JSON. No external npm dependencies. A missing Python is a hard
 * error with a remediation hint; unsupported file types (PDF) are reported as a
 * structured error, not a crash.
 *
 * @module dsh-context-extract/runner
 */
import { spawn } from 'node:child_process'
import { existsSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path'

/** Resolve $DSH_HOME like the host does. */
function dshHome() {
  const override = process.env.DSH_HOME
  return (override !== undefined && override.length > 0 ? override : join(homedir(), '.dsh'))
}

function firstExisting(candidates) {
  for (const candidate of candidates) {
    if (candidate === undefined || candidate.length === 0) continue
    try {
      if (existsSync(candidate) && statSync(candidate).isFile()) return candidate
    } catch {
      /* ignore unreadable candidates */
    }
  }
  return undefined
}

/** Locate the bundled Python (or a PATH python3 as last resort). */
export function findPython() {
  const runtime = join(dshHome(), 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python')
  const found = firstExisting([
    process.env.DSH_CONTEXT_EXTRACT_PYTHON,
    join(runtime, 'bin', 'python3'),
    join(runtime, 'bin', 'python'),
    join(runtime, 'python.exe'),
    join(runtime, 'bin', 'python.exe'),
  ])
  if (found !== undefined) return found
  return 'python3'
}

/** Run a child process, capture stdout/stderr, enforce a timeout. */
function exec(command, args, options = {}) {
  return new Promise((resolvePromise) => {
    const child = spawn(command, args, {
      cwd: options.cwd,
      env: { ...process.env, ...(options.env ?? {}) },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let settled = false
    const timeoutMs = options.timeoutMs ?? 300_000
    const timer = setTimeout(() => {
      if (settled) return
      settled = true
      child.kill('SIGKILL')
      resolvePromise({ code: -1, stdout, stderr: `${stderr}\n[timeout after ${timeoutMs}ms]`, timedOut: true })
    }, timeoutMs)
    child.stdout.on('data', (chunk) => { stdout += String(chunk) })
    child.stderr.on('data', (chunk) => { stderr += String(chunk) })
    child.on('error', (error) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolvePromise({ code: -1, stdout, stderr: `${stderr}\n${String(error?.message ?? error)}`, spawnError: error })
    })
    child.on('close', (code) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolvePromise({ code: code ?? -1, stdout, stderr })
    })
  })
}

function timestamp() {
  const d = new Date()
  const pad = (n) => String(n).padStart(2, '0')
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`
}

/**
 * Extract structured context from a raw file.
 *
 * @param {object} options
 * @param {string} options.input - absolute input path
 * @param {string[]} [options.targets] - optional line-item filter words
 * @param {string} [options.outputDir] - report directory
 * @param {string} options.engineRoot - plugin root (contains core/extract.py)
 */
export async function runExtract(options) {
  const input = resolve(options.input)
  if (!existsSync(input)) return { ok: false, error: `input not found: ${input}` }

  const python = findPython()
  const engine = join(options.engineRoot, 'core', 'extract.py')
  if (!existsSync(engine)) return { ok: false, error: `extract engine missing: ${engine}` }

  const ext = extname(input).toLowerCase()
  const reportDir = options.outputDir ?? join(dirname(input), '.context-extract')
  const stem = basename(input, ext)
  const stamp = timestamp()
  const markdownPath = join(reportDir, `${stem}-extract-${stamp}.md`)
  const jsonPath = join(reportDir, `${stem}-extract-${stamp}.json`)

  const args = [engine, '--input', input, '--out-json', jsonPath, '--out-md', markdownPath]
  if (Array.isArray(options.targets) && options.targets.length > 0) {
    args.push('--targets', options.targets.join(','))
  }

  const result = await exec(python, args, { cwd: dirname(engine), timeoutMs: 300_000 })
  if (result.code !== 0) {
    const detail = (result.stderr || result.stdout || '').trim().split('\n').slice(-8).join('\n')
    return {
      ok: false,
      error: `extract engine failed (exit ${result.code})${detail.length > 0 ? `: ${detail}` : ''}`,
      hint: 'check that the DSH Python runtime has openpyxl / python-docx (load_workspace_dependencies reports it)',
    }
  }
  let payload
  try {
    payload = JSON.parse(result.stdout)
  } catch {
    return { ok: false, error: 'extract engine returned invalid JSON', stdoutTail: result.stdout.slice(-2000), stderrTail: result.stderr.slice(-2000) }
  }
  if (payload && typeof payload === 'object' && payload.ok) {
    payload.report = { markdownPath, jsonPath }
  }
  return payload
}
