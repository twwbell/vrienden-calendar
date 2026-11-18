import { execFileSync } from 'node:child_process'
import { readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const N8N_DIR = join(dirname(fileURLToPath(import.meta.url)), '..')
export const WORKFLOWS_DIR = join(N8N_DIR, 'workflows')
export const PROMPT_FILE = join(N8N_DIR, 'prompts', 'system.md')
export const CONFIG_FILE = join(N8N_DIR, 'config', 'config.json')
export const CONTAINER = process.env.N8N_CONTAINER || 'n8n-n8n-1'

/** The main workflow; push publishes it so its schedule runs. */
export const MAIN_WORKFLOW_ID = 'vriendenCalDaily'

/** Values the workflows read from the Config node. */
export const CONFIG_KEYS = ['calendarId', 'birthdaySheetId', 'birthdaySheetGid', 'gistId', 'gistFile', 'errorMailFrom', 'errorMailTo']

/** Credential types the workflows use; config.credentials maps each to an n8n credential ID. */
export const CREDENTIAL_TYPES = ['googleApi', 'googlePalmApi', 'githubApi', 'smtp']

/** Marker in the files where push puts the system prompt and pull takes it out again. */
export const PROMPT_PLACEHOLDER = '(filled from n8n/prompts/system.md by push.mjs)'

export function workflowFiles() {
  return readdirSync(WORKFLOWS_DIR)
    .filter((file) => file.endsWith('.json'))
    .sort()
    .map((file) => join(WORKFLOWS_DIR, file))
}

export function readJson(path) {
  return JSON.parse(readFileSync(path, 'utf8'))
}

/** Returns { config, problems }. Problems are empty values; push refuses to publish while there are any. */
export function readConfig() {
  let config
  try {
    config = readJson(CONFIG_FILE)
  } catch {
    throw new Error(`Missing or invalid ${CONFIG_FILE}. Copy config/example.json and fill it in.`)
  }
  const problems = []
  for (const key of CONFIG_KEYS) {
    if (config[key] === undefined || String(config[key]).trim() === '') problems.push(key)
  }
  for (const type of CREDENTIAL_TYPES) {
    if (!config.credentials?.[type]) problems.push(`credentials.${type}`)
  }
  return { config, problems }
}

/**
 * The version of a workflow file that goes into n8n: the Config node gets the
 * config values, the agent gets the system prompt and every node that uses a
 * credential type gets the ID from config.credentials.
 */
export function forN8n(workflow, config) {
  const result = structuredClone(workflow)
  result.active = false
  const values = Object.fromEntries(CONFIG_KEYS.map((key) => [key, config[key] ?? '']))
  for (const node of result.nodes) {
    if (node.name === 'Config') node.parameters.jsonOutput = JSON.stringify(values, null, 2)
    if (node.parameters?.options?.systemMessage === PROMPT_PLACEHOLDER) {
      node.parameters.options.systemMessage = readFileSync(PROMPT_FILE, 'utf8').trim()
    }
    for (const type of Object.keys(node.credentials ?? {})) {
      // An empty id keeps the credential type on the node, so pull does not lose it.
      node.credentials[type] = { id: config.credentials?.[type] || '', name: type }
    }
  }
  return result
}

/** The reverse of forN8n, for pull: no config values, prompt or credential IDs in the files. */
export function forFiles(workflow) {
  const nodes = workflow.nodes.map((node) => {
    const copy = structuredClone(node)
    if (copy.name === 'Config') copy.parameters.jsonOutput = '{}'
    if (copy.parameters?.options?.systemMessage !== undefined) copy.parameters.options.systemMessage = PROMPT_PLACEHOLDER
    for (const type of Object.keys(copy.credentials ?? {})) copy.credentials[type] = { id: '', name: type }
    return copy
  })
  return {
    id: workflow.id,
    name: workflow.name,
    active: false,
    nodes,
    connections: workflow.connections,
    settings: workflow.settings,
  }
}

/** Runs the n8n CLI in the container and returns its output without the startup noise. */
export function n8n(args) {
  const output = execFileSync('docker', ['exec', CONTAINER, 'n8n', ...args], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  return output
    .split('\n')
    .filter((line) => !/Postgres 16|Deprecation warning|migration lock/.test(line))
    .join('\n')
    .trim()
}

export function docker(args) {
  return execFileSync('docker', args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] })
}
