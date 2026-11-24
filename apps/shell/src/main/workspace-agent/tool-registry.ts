import type { AgentToolCall, AgentToolDef, ToolExecution } from '@genoffice/agent-core'
import type {
  HarnessGuardHook,
  HarnessPostExecuteHook,
  HarnessPreExecuteHook,
  HarnessToolContext,
  HarnessToolExecutor,
  HarnessToolRegistration,
} from './harness-types'

const TOOL_NAME_RE = /^[a-zA-Z0-9_-]+(?:\.[a-zA-Z0-9_-]+)*$/

import { validateToolInput } from '../../shared/tool-input-schema'
export { validateToolInput } from '../../shared/tool-input-schema'

function errorExecution(name: string, message: string): ToolExecution {
  return { output: message, summary: name, isError: true, effect: 'not_applied' }
}

function assertNotCancelled(context: HarnessToolContext): void {
  if (context.signal?.aborted) throw new Error('Tool execution was cancelled.')
}

/**
 * Harness-compatible tool registry.
 *
 * The provider sees only `AgentToolDef`s. Every execution passes through the
 * same pre -> guard -> executor -> post pipeline, which keeps permissions,
 * telemetry and future confirmation UI out of individual Office tools.
 */
export class HarnessToolRegistry {
  private readonly registrations = new Map<string, HarnessToolRegistration>()
  private readonly preHooks: HarnessPreExecuteHook[] = []
  private readonly guardHooks: HarnessGuardHook[] = []
  private readonly postHooks: HarnessPostExecuteHook[] = []

  constructor(initial: readonly HarnessToolRegistration[] = []) {
    for (const registration of initial) this.register(registration.definition, registration.execute)
  }

  register(definition: AgentToolDef, execute: HarnessToolExecutor): this {
    if (!definition || typeof definition.name !== 'string' || !TOOL_NAME_RE.test(definition.name)) {
      throw new Error(`Invalid Harness tool name: ${String(definition?.name ?? '')}`)
    }
    if (this.registrations.has(definition.name))
      throw new Error(`Duplicate Harness tool name: ${definition.name}`)
    this.registrations.set(definition.name, {
      definition: { ...definition, inputSchema: { ...definition.inputSchema } },
      execute,
    })
    return this
  }

  /** Register an existing definition only when a plugin did not provide it. */
  registerIfAbsent(definition: AgentToolDef, execute: HarnessToolExecutor): this {
    if (!this.registrations.has(definition.name)) this.register(definition, execute)
    return this
  }

  unregister(name: string): boolean {
    return this.registrations.delete(name)
  }

  has(name: string): boolean {
    return this.registrations.has(name)
  }

  get(name: string): HarnessToolRegistration | undefined {
    return this.registrations.get(name)
  }

  definitions(): AgentToolDef[] {
    return [...this.registrations.values()].map(({ definition }) => ({
      ...definition,
      inputSchema: { ...definition.inputSchema },
    }))
  }

  registrationsSnapshot(): HarnessToolRegistration[] {
    return [...this.registrations.values()]
  }

  /** Clone definitions, executors and hooks for a per-session registry. */
  clone(): HarnessToolRegistry {
    const copy = new HarnessToolRegistry(this.registrationsSnapshot())
    copy.preHooks.push(...this.preHooks)
    copy.guardHooks.push(...this.guardHooks)
    copy.postHooks.push(...this.postHooks)
    return copy
  }

  addPreExecuteHook(hook: HarnessPreExecuteHook): () => void {
    this.preHooks.push(hook)
    return () => removeHook(this.preHooks, hook)
  }

  addGuard(hook: HarnessGuardHook): () => void {
    this.guardHooks.push(hook)
    return () => removeHook(this.guardHooks, hook)
  }

  addPostExecuteHook(hook: HarnessPostExecuteHook): () => void {
    this.postHooks.push(hook)
    return () => removeHook(this.postHooks, hook)
  }

  // Names used by the Harness SDK's middleware terminology; keep both forms
  // so plugins can be ported without an adapter-specific shim.
  usePreExecute(hook: HarnessPreExecuteHook): () => void {
    return this.addPreExecuteHook(hook)
  }

  useGuard(hook: HarnessGuardHook): () => void {
    return this.addGuard(hook)
  }

  usePostExecute(hook: HarnessPostExecuteHook): () => void {
    return this.addPostExecuteHook(hook)
  }

  async execute(call: AgentToolCall, context: HarnessToolContext = {}): Promise<ToolExecution> {
    let normalized = call
    let execution: ToolExecution
    let dispatched = false
    try {
      assertNotCancelled(context)
      for (const hook of this.preHooks) {
        const next = await hook(normalized, context)
        if (next) normalized = next
        assertNotCancelled(context)
      }

      // Resolve the registration after pre-execute hooks. A Harness adapter may
      // normalize an alias (or otherwise rewrite the call) and the rewritten
      // name must select the corresponding executor and policy path.
      const registration = this.registrations.get(normalized.name)
      if (!registration) {
        execution = errorExecution(normalized.name, `Unknown tool: ${normalized.name}`)
      } else {
        const invalid = validateToolInput(registration.definition.inputSchema, normalized.input)
        if (invalid) {
          execution = errorExecution(normalized.name, `Invalid tool input: ${invalid}`)
        } else {
          let denied: string | false | undefined
          for (const hook of this.guardHooks) {
            const result = await hook(normalized, context)
            assertNotCancelled(context)
            if (result === false) {
              denied = false
              break
            }
            if (typeof result === 'string' && result.trim()) {
              denied = result
              break
            }
          }
          if (denied === false || typeof denied === 'string') {
            execution = errorExecution(
              normalized.name,
              denied || `Tool execution was denied: ${normalized.name}`,
            )
          } else {
            dispatched = true
            execution = await registration.execute(normalized, context)
          }
        }
      }
    } catch (error) {
      execution = {
        ...errorExecution(normalized.name, error instanceof Error ? error.message : String(error)),
        effect: dispatched ? 'unknown' : 'not_applied',
      }
    }

    // Post hooks are deliberately best-effort but always receive a terminal
    // execution result, including guard denials and executor failures. This
    // lets audit/telemetry plugins observe every attempted tool invocation.
    for (const hook of this.postHooks) {
      try {
        const next = await hook(normalized, execution, context)
        if (next) execution = next
      } catch (error) {
        // A post-execute audit failure cannot undo a completed write. Preserve
        // mutation metadata so run checkpoints and rollback still account for it.
        execution = {
          ...execution,
          isError: true,
          output: `${execution.output}\nPost-execute hook failed: ${error instanceof Error ? error.message : String(error)}`,
        }
      }
    }
    return execution
  }
}

function removeHook<T>(list: T[], hook: T): void {
  const index = list.indexOf(hook)
  if (index >= 0) list.splice(index, 1)
}

export { TOOL_NAME_RE }
