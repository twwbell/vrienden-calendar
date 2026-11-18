#!/usr/bin/env node
// Copies the workflows from n8n back into the files, after a change in the editor.
//   node n8n/scripts/pull.mjs
//
// Config values, the system prompt and credential IDs are left out of the
// files; they live in config/config.json and prompts/system.md. A prompt
// edited in the editor is therefore NOT pulled: edit prompts/system.md.
import { writeFileSync } from 'node:fs'
import { CONTAINER, docker, forFiles, n8n, readJson, workflowFiles } from './lib.mjs'

const remoteFile = `/tmp/vrienden-export-${process.pid}.json`

for (const file of workflowFiles()) {
  const { id } = readJson(file)
  try {
    n8n(['export:workflow', `--id=${id}`, `--output=${remoteFile}`])
    const exported = JSON.parse(docker(['exec', CONTAINER, 'cat', remoteFile]))
    const workflow = Array.isArray(exported) ? exported[0] : exported
    writeFileSync(file, JSON.stringify(forFiles(workflow), null, 2) + '\n')
    console.log(`pulled ${workflow.name}`)
  } finally {
    try {
      docker(['exec', CONTAINER, 'rm', '-f', remoteFile])
    } catch {
      // nothing to clean up
    }
  }
}
