/**
 * Host half of dsh-context-extract.
 *
 * Registers one agent tool, `context_extract`, plus a short system-prompt
 * section that tells the agent when and how to use it. The tool deterministically
 * compresses a raw financial/contract/ledger file into high-density structured
 * Markdown + JSON (key line items, tables, clause split, entities, 勾稽 alerts),
 * so downstream M1-M4 modules get clean, token-lean context instead of the raw
 * file. All heavy lifting lives in the bundled Python engine (core/extract.py,
 * openpyxl + python-docx); runner.js locates the runtime and normalizes its JSON.
 *
 * Keeps ZERO runtime npm dependencies (mirrors dsh-office-audit / dsh-workpaper-gate).
 *
 * @module dsh-context-extract
 */
import { fileURLToPath } from 'node:url'
import { dirname, isAbsolute, resolve } from 'node:path'
import { runExtract } from './runner.js'

/** Cordis plugin name. */
export const name = 'dsh-context-extract'

/** Required host services: tool registry + prompt assembly. */
export const inject = ['tools', 'systemPrompt']

const HERE = dirname(fileURLToPath(import.meta.url))

/* ------------------------------------------------------------------ */
/* defineTool (structure-compatible reimplementation of the SDK's)     */
/* ------------------------------------------------------------------ */

function compileValue(spec) {
  const node = {}
  if (typeof spec.description === 'string' && spec.description.length > 0) node.description = spec.description
  const type = spec.type
  if (type === undefined || type === 'json') return node
  if (type === 'object') {
    node.type = 'object'
    node.additionalProperties = spec.additionalProperties
    if (spec.properties !== undefined) node.properties = compilePropertyMap(spec.properties).properties
    return node
  }
  if (type === 'array') {
    node.type = 'array'
    if (spec.items !== undefined) node.items = compileValue(spec.items)
    return node
  }
  node.type = type
  if (spec.enum !== undefined) node.enum = [...spec.enum]
  if (spec.const !== undefined) node.const = spec.const
  return node
}

function compilePropertyMap(spec) {
  const properties = {}
  const required = []
  for (const [key, entry] of Object.entries(spec)) {
    const { required: isRequired, ...valueSpec } = entry
    properties[key] = compileValue(valueSpec)
    if (isRequired === true) required.push(key)
  }
  return required.length > 0 ? { properties, required } : { properties }
}

function matchesScalarType(value, type) {
  switch (type) {
    case 'string': return typeof value === 'string'
    case 'number': return typeof value === 'number'
    case 'integer': return typeof value === 'number' && Number.isInteger(value)
    case 'boolean': return typeof value === 'boolean'
    case 'null': return value === null
    default: return true
  }
}

function validateValue(schema, value, path) {
  if (typeof schema.type !== 'string' || schema.type.length === 0) return []
  if (schema.type === 'object') {
    if (typeof value !== 'object' || value === null || Array.isArray(value)) return [`${path} must be an object`]
    const violations = []
    for (const key of schema.required ?? []) if (!(key in value)) violations.push(`${path}.${key} is required`)
    if (schema.additionalProperties === false) {
      const known = new Set(Object.keys(schema.properties ?? {}))
      for (const key of Object.keys(value)) if (!known.has(key)) violations.push(`${path}.${key} is not a declared property`)
    }
    for (const [key, child] of Object.entries(schema.properties ?? {})) {
      if (key in value) violations.push(...validateValue(child, value[key], `${path}.${key}`))
    }
    return violations
  }
  if (schema.type === 'array') {
    if (!Array.isArray(value)) return [`${path} must be an array`]
    const violations = []
    if (schema.items !== undefined) value.forEach((item, i) => violations.push(...validateValue(schema.items, item, `${path}[${i}]`)))
    return violations
  }
  if (!matchesScalarType(value, schema.type)) return [`${path} must be ${schema.type}`]
  if (schema.enum !== undefined && !schema.enum.some((v) => v === value)) return [`${path} must be one of ${schema.enum.map(String).join(', ')}`]
  if (schema.const !== undefined && schema.const !== value) return [`${path} must be ${String(schema.const)}`]
  return []
}

function defineTool(options) {
  const compiled = compilePropertyMap(options.parameters)
  const parameters = { type: 'object', properties: compiled.properties }
  if (compiled.required !== undefined) parameters.required = compiled.required
  return {
    name: options.name,
    description: options.description,
    parameters,
    output: {
      schema: {},
      render(args, value) {
        return options.output.render(args, value)
      },
    },
    async execute(args, exec) {
      const violations = validateValue(parameters, args, 'arguments')
      if (violations.length > 0) throw new Error(`Error: invalid arguments: ${violations.join('; ')}`)
      return options.execute(args, exec)
    },
  }
}

/* ------------------------------------------------------------------ */
/* rendering                                                           */
/* ------------------------------------------------------------------ */

function renderExtract(_args, value) {
  if (value === null || typeof value !== 'object') return [{ type: 'text', text: String(value) }]
  if (value.error !== undefined || value.ok === false) return [{ type: 'text', text: `萃取失败：${value.error}` }]
  // The engine already produced the dense compressed markdown — hand it straight
  // to the agent so it can consume the clean context without the raw file.
  const lines = []
  if (typeof value.markdown === 'string' && value.markdown.length > 0) {
    lines.push(value.markdown)
  }
  const meta = []
  if (typeof value.token_estimate === 'number') meta.push(`≈${value.token_estimate} tokens`)
  if (Array.isArray(value.alerts) && value.alerts.length > 0) meta.push(`勾稽提示 ${value.alerts.length} 条`)
  if (value.report?.markdownPath) meta.push(`完整报告：${value.report.markdownPath}`)
  if (meta.length > 0) lines.push('', meta.join(' · '))
  return [{ type: 'text', text: lines.join('\n') }]
}

/* ------------------------------------------------------------------ */
/* plugin apply                                                        */
/* ------------------------------------------------------------------ */

const PROMPT_SECTION = `## 上下文萃取与压缩器（context_extract）

当用户上传或指向几十页财报、长篇幅协议、工商底稿等超长原始文件，或要求“压缩 / 萃取 / 结构化”上下文时，先用 context_extract 工具把原始文件压成高密度结构化 Markdown / JSON，再基于结果处理，不要直接整份读原文件。

- 参数 path：文件路径（.xlsx / .xlsm / .csv / .docx / .txt / .md；PDF 需先用 LibreOffice Kit 转成文本或 xlsx）。
- 参数 targets：可选，只抽取指定指标（如 ["货币资金","净利润"]），忽略说明性长段落。
- 工具自动：识别文件类型（财务报表 / 法律合同 / 业务台账）、抽取关键主体 / 日期 / 金额、对财务报表做勾稽校验（资产总计 vs 负债+权益、货币资金 vs 期末现金），不一致以 <alert> 标出。
- 输出：密集 Markdown 表格 + JSON（保留原单位，如人民币千元），并给 token 估算。

萃取后再按这四条压缩，喂给下游 M1–M4：
1. 禁止解释：不要输出“根据图片可知”“以下是我提取的结果”等过渡语。
2. 按需抽取：指定了指标就只给带年份的交叉数据。
3. 表格优先：财务数据、交易节点、金额一律 Markdown 表格或 JSON，保留单位。
4. 条款编号化：法律协议按“条款号 + 核心义务/权利”精简为一句话，严禁照抄长段落。`

function resolveInputPath(args, exec) {
  let p = args.path
  if (typeof p !== 'string' || p.trim().length === 0) throw new Error('Error: invalid_input: path is required')
  if (!isAbsolute(p)) {
    const cwd = exec?.agent?.session?.header?.cwd
    if (typeof cwd !== 'string' || cwd.length === 0) throw new Error('Error: invalid_input: relative path given but the session has no working directory')
    p = resolve(cwd, p)
  }
  return p
}

export function apply(ctx) {
  const disposers = []

  disposers.push(ctx.tools.register(defineTool({
    name: 'context_extract',
    description: 'Compress a raw financial/contract/ledger file (.xlsx/.xlsm/.csv/.docx/.txt/.md) into high-density structured Markdown + JSON for downstream modules: detects file type (financial statement / contract / ledger), extracts key line items, entities (companies/dates/amounts), splits contracts into numbered clauses, and runs light cross-checks (balance-sheet assets vs liabilities+equity, cash tie-out) emitting <alert> tags on mismatch. Returns dense markdown plus report paths and a token estimate. Read-only; never modifies the source file. PDF is not supported — convert it first.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Input file path (.xlsx/.xlsm/.csv/.docx/.txt/.md). Absolute, or relative to the session working directory.',
      },
      targets: {
        type: 'array',
        items: { type: 'string' },
        description: 'Optional line-item filter words, e.g. ["货币资金","净利润"]. When set, only matching key items are extracted.',
      },
      outputDir: {
        type: 'string',
        description: 'Directory for the report files (default: <input dir>/.context-extract).',
      },
    },
    output: { render: renderExtract },
    async execute(args, exec) {
      const input = resolveInputPath(args, exec)
      const outputDir = typeof args.outputDir === 'string' && args.outputDir.trim().length > 0
        ? (isAbsolute(args.outputDir) ? args.outputDir : resolve(dirname(input), args.outputDir))
        : undefined
      return runExtract({
        input,
        targets: Array.isArray(args.targets) ? args.targets : [],
        outputDir,
        engineRoot: HERE,
      })
    },
  })))

  const disposeSection = ctx.systemPrompt.section({
    name: 'context-extract',
    order: 186,
    text: PROMPT_SECTION,
  })
  ctx.effect(() => disposeSection, 'dsh-context-extract: prompt section')

  ctx.effect(() => () => {
    for (const dispose of disposers.splice(0)) dispose()
  }, 'dsh-context-extract: tools')
}
