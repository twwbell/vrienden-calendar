#!/usr/bin/env node
// Imports the workflow files into n8n and publishes the daily workflow.
//   node n8n/scripts/push.mjs
//
// With empty values in config/config.json the workflows are imported but not
// published, so nothing runs on a schedule until the config is complete.
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { CONTAINER, MAIN_WORKFLOW_ID, docker, forN8n, n8n, readConfig, readJson, workflowFiles } from './lib.mjs'

const { config, problems } = readConfig()
const files = workflowFiles()

const localDir = mkdtempSync(join(tmpdir(), 'vrienden-n8n-'))
const remoteDir = `/tmp/vrienden-import-${process.pid}`

try {
  for (const file of files) {
    const result = forN8n(readJson(file), config)
    writeFileSync(join(localDir, basename(file)), JSON.stringify(result, null, 2))
    console.log(`${result.id}  ${result.name}`)
  }
  docker(['cp', localDir, `${CONTAINER}:${remoteDir}`])
  console.log(n8n(['import:workflow', '--separate', `--input=${remoteDir}`]))

  if (problems.length > 0) {
    console.warn(`\nNot published: config/config.json has no value for: ${problems.join(', ')}`)
    process.exitCode = 1
  } else {
    console.log(n8n(['publish:workflow', `--id=${MAIN_WORKFLOW_ID}`]))
    console.log('\nImported and published. Restart n8n (docker restart n8n-n8n-1) if the schedule does not show as active in the editor.')
  }
} finally {
  rmSync(localDir, { recursive: true, force: true })
  try {
    docker(['exec', CONTAINER, 'rm', '-rf', remoteDir])
  } catch {
    console.warn(`Could not remove ${remoteDir} in the container.`)
  }
}
