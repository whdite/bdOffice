import { explicitlyRequestsSave } from '@genoffice/agent-core'
import { randomUUID } from 'node:crypto'
import type { AssistantTaskService } from '../agent-tasks/service'
import type { WorkspaceAgentRuntime } from '../workspace-agent/service'
import type { AgentInput } from '../../shared/workspace-agent-api'

/** Both desktop IPC and the authenticated mobile adapter enter the existing M3 scheduler here. */
export async function submitAssistantMessage(
  tasks: AssistantTaskService,
  runtime: WorkspaceAgentRuntime,
  sessionId: string,
  input: AgentInput,
  waitForCompletion = true,
): Promise<{ taskId: string; duplicate: boolean }> {
  const question =
    !input.images?.length &&
    !input.instructions &&
    tasks.store
      .forSession(sessionId)
      .find((item) => item.state === 'waiting_user' && item.confirmation?.kind === 'clarification')
  if (question && question.confirmation) {
    const receipt = tasks.answer(
      {
        taskId: question.id,
        confirmationId: question.confirmation.id,
        requestId: input.requestId ?? randomUUID(),
        expectedRevision: question.revision,
        answer: input.text,
      },
      false,
    )
    if (!receipt.duplicate) await runtime.recordTaskReply(sessionId, input.text)
    tasks.start()
    return { taskId: receipt.task.id, duplicate: receipt.duplicate }
  }
  const document = await runtime.assistantTaskDocument(sessionId, input, tasks.bootId)
  const receipt = tasks.submit({
    requestId: input.requestId ?? randomUUID(),
    sessionId,
    entry: 'workspace',
    input,
    document,
    requiredResult: explicitlyRequestsSave(input.text)
      ? 'saved_artifact'
      : (input.requiredResult ?? 'answer'),
  })
  if (waitForCompletion) {
    const result = await tasks.wait(receipt.task.id)
    if (result.state === 'failed' && !result.steps.length && result.reason)
      throw new Error(result.reason)
  }
  return { taskId: receipt.task.id, duplicate: receipt.duplicate }
}
