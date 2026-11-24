import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ToolExecution } from '@genoffice/agent-core'
import { cliErrorMessage, type CliRunner } from '../mcp/cli-runner'

/** The CLI owns its DOM in a child process. Workspace sessions keep the authoritative staged bytes. */
export class WordFileExecutor {
  constructor(private readonly runner: CliRunner) {}

  private async temporary<T>(run: (dir: string) => Promise<T>): Promise<T> {
    const dir = await mkdtemp(join(tmpdir(), 'bigdog-word-'))
    try {
      return await run(dir)
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  }

  create(content: string, format: 'md' | 'html', signal?: AbortSignal): Promise<Uint8Array> {
    return this.temporary(async (dir) => {
      const src = join(dir, 'source.' + format),
        out = join(dir, 'created.docx')
      await writeFile(src, content)
      const run = await this.runner.run(['create', '--type', 'docx', '--from', src, '--out', out], {
        cwd: dir,
        signal,
        env: { GENOFFICE_ALLOWED_ROOTS: dir },
      })
      if (!run.ok) throw new Error(cliErrorMessage(run))
      return readFile(out)
    })
  }

  execute(bytes: Uint8Array, name: string, input: Record<string, unknown>, signal?: AbortSignal) {
    return this.temporary(async (dir) => {
      const file = join(dir, 'source.docx'),
        out = join(dir, 'edited.docx')
      await writeFile(file, bytes)
      const run = await this.runner.run(['docs', 'tool', file, '--ops', '-', '--out', out], {
        cwd: dir,
        input: JSON.stringify({ name, input }),
        signal,
        env: { GENOFFICE_ALLOWED_ROOTS: dir },
      })
      if (!run.ok || run.json?.status !== 'ok') throw new Error(cliErrorMessage(run))
      const result = run.json.detail?.result as ToolExecution | undefined
      if (!result || typeof result.output !== 'string')
        throw new Error('Missing Word tool receipt.')
      if (result.isError) return { result: { ...result, effect: 'not_applied' as const } }
      return {
        result,
        bytes: result.effect === 'applied' || result.mutated ? await readFile(out) : undefined,
      }
    })
  }
}
