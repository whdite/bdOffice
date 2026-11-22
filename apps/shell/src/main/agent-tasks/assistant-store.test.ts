import assert from 'node:assert/strict'
import { randomUUID } from 'node:crypto'
import test from 'node:test'
import { sessionBudget, type AgentSessionSnapshot } from '@genoffice/agent-core'
import type { TaskTargetVersion } from '../../shared/assistant-task-api'
import type { AssistantTaskStore } from './assistant-store'
import { TaskStore } from './store'

const slides: TaskTargetVersion = { path: '/fixtures/slides.pptx', exists: true, complete: true }
const word: TaskTargetVersion = { path: '/fixtures/word.docx', exists: true, complete: true }
const snapshot: AgentSessionSnapshot = {
  version: 1,
  messages: [],
  phase: 'running',
  budget: sessionBudget(false),
  updatedAt: new Date().toISOString(),
}

function accept(store: AssistantTaskStore, sessionId: string, document: TaskTargetVersion) {
  return store.accept({
    requestId: randomUUID(),
    sessionId,
    entry: 'file',
    input: { text: 'Edit the document' },
    requiredResult: 'answer',
    document,
  }).task.id
}

const edit = (document: TaskTargetVersion) => ({
  digest: 'apply-ops',
  label: 'apply_ops',
  recovery: 'verify_effect' as const,
  targets: [document],
})

test('an interrupted editor from a previous boot leaves unrelated tasks runnable', (context) => {
  const ledger = new TaskStore(':memory:')
  context.after(() => ledger.close())
  const { assistant } = ledger
  assistant.startBoot('old-boot')
  const interrupted = accept(assistant, 'slides-session', slides)
  const lease = assistant.claim(interrupted, 'old-boot')
  assert.ok(lease)
  assistant.begin(lease, snapshot)
  assistant.prepareStep(lease, edit(slides))
  // Reproduce a resource already marked quarantined in an installed profile.
  assistant.finish(lease, { state: 'failed', reason: 'executor_lost' })
  assistant.startBoot('new-boot')
  const previous = assistant.get(interrupted)
  assert.equal(previous.reason, 'outcome_unknown')
  assert.equal(previous.steps[0].effect, 'unknown')

  const sameSession = accept(assistant, 'slides-session', slides)
  assert.equal(assistant.claim(sameSession, 'new-boot'), null)
  assert.equal(assistant.get(sameSession).reason, 'task_effect_requires_review')

  const independent = accept(assistant, 'word-session', word)
  const wordLease = assistant.claim(independent, 'new-boot')
  assert.ok(wordLease)
  assistant.begin(wordLease, snapshot)
  const wordStep = assistant.prepareStep(wordLease, edit(word))
  const sameDocument = accept(assistant, 'other-session', slides)
  assert.equal(assistant.claim(sameDocument, 'new-boot'), null)
  assert.equal(assistant.get(sameDocument).reason, 'resource_busy')
  assistant.finishStep(
    wordLease,
    wordStep.stepId,
    {
      result: { output: 'Done', summary: 'No edit needed' },
      targets: [word],
      effect: 'not_applied',
    },
    snapshot,
  )
  assistant.finish(wordLease, { state: 'completed' })

  const slidesLease = assistant.claim(sameDocument, 'new-boot')
  assert.ok(slidesLease)
  assistant.begin(slidesLease, snapshot)
  assert.throws(
    () => assistant.prepareStep(slidesLease, edit(slides)),
    /task_effect_requires_review/,
  )
  assistant.control(
    {
      taskId: interrupted,
      requestId: randomUUID(),
      expectedRevision: previous.revision,
      action: 'reconcile',
      reviewed: true,
    },
    [slides],
  )
  assert.ok(assistant.prepareStep(slidesLease, edit(slides)).stepId)
})

test('restart preserves the global review requirement for unscoped effects', (context) => {
  const ledger = new TaskStore(':memory:')
  context.after(() => ledger.close())
  const { assistant } = ledger
  assistant.startBoot('old-boot')
  const interrupted = accept(assistant, 'slides-session', slides)
  const lease = assistant.claim(interrupted, 'old-boot')
  assert.ok(lease)
  assistant.begin(lease, snapshot)
  assistant.prepareStep(lease, { ...edit(slides), recovery: 'user_reconcile', targets: [] })

  assistant.startBoot('new-boot')
  const independent = accept(assistant, 'word-session', word)
  assert.equal(assistant.get(interrupted).reason, 'outcome_unknown')
  assert.equal(assistant.claim(independent, 'new-boot'), null)
  assert.equal(assistant.get(independent).reason, 'task_effect_requires_review')
})
