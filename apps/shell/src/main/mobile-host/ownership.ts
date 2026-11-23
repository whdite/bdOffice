import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'
import type { HostServerApi } from '@bigdogoffice/remote-core/host'
interface RecordEntry {
  globalId: string
  origin?: string
  epoch?: number
  state: 'active' | 'frozen' | 'transferred'
}
/** Persisted fencing survives restart. Missing/expired leases cannot authorize a write. */
export class RemoteOwnership {
  private records: Record<string, RecordEntry>
  private deadlines = new Map<string, number>()
  private api?: HostServerApi
  private pending = new Map<string, Promise<void>>()
  constructor(
    private readonly path: string,
    private readonly previousOrigin?: string,
  ) {
    this.records = existsSync(path) ? JSON.parse(readFileSync(path, 'utf8')) : {}
  }
  configure(api: HostServerApi) {
    if (this.previousOrigin) {
      const legacy = Object.values(this.records).filter((record) => !record.origin)
      for (const record of legacy) record.origin = this.previousOrigin
      if (legacy.length) this.save()
    }
    this.api = api
    this.deadlines.clear()
    this.pending.clear()
  }
  disconnect(api: HostServerApi) {
    if (this.api !== api) return
    this.api = undefined
    this.deadlines.clear()
    this.pending.clear()
  }
  private save() {
    mkdirSync(dirname(this.path), { recursive: true })
    writeFileSync(this.path + '.pending', JSON.stringify(this.records))
    renameSync(this.path + '.pending', this.path)
  }
  allowed(id: string) {
    return (
      !this.records[id] ||
      (this.records[id].state === 'active' && (this.deadlines.get(id) ?? 0) > performance.now())
    )
  }
  async ensure(id: string) {
    const api = this.api
    if (!api) return
    const existing = this.records[id]
    if (
      existing &&
      (existing.state !== 'active' || (existing.origin && existing.origin !== api.baseUrl))
    )
      throw new Error('OWNER_CHANGED')
    if (!this.records[id]) {
      this.records[id] = { globalId: id, state: 'active', origin: api.baseUrl }
      this.save()
    }
    if ((this.deadlines.get(id) ?? 0) > performance.now() + 3000) return
    if (this.pending.has(id)) return this.pending.get(id)
    const started = performance.now(),
      record = this.records[id]
    const promise = (async () => {
      const lease = await api.courier('lease', {
        sessionId: record.globalId,
        ...(record.epoch ? { epoch: record.epoch } : {}),
      })
      if (this.api !== api || this.records[id] !== record || record.state !== 'active')
        throw new Error('OWNER_CHANGED')
      record.origin = api.baseUrl
      record.epoch = lease.epoch
      this.save()
      this.deadlines.set(id, started + 10_000)
      if (!this.allowed(id)) throw new Error('LEASE_EXPIRED')
    })().finally(() => {
      if (this.pending.get(id) === promise) this.pending.delete(id)
    })
    this.pending.set(id, promise)
    return promise
  }
  async renew() {
    for (const [id, record] of Object.entries(this.records))
      if (record.state === 'active') await this.ensure(id).catch(() => {})
  }
  async freeze(id: string) {
    await this.ensure(id)
    this.records[id].state = 'frozen'
    this.deadlines.delete(id)
    this.save()
    return this.records[id]
  }
  transferred(id: string) {
    if (this.records[id]) {
      this.records[id].state = 'transferred'
      this.deadlines.delete(id)
      this.save()
    }
  }
  stage(localId: string, globalId: string, epoch: number) {
    this.records[localId] = { globalId, epoch, state: 'frozen', origin: this.api?.baseUrl }
    this.save()
  }
  activate(localId: string) {
    this.records[localId].state = 'active'
    this.save()
    return this.ensure(localId)
  }
  snapshot(id: string) {
    return this.records[id]
  }
  writeBlocked(id: string) {
    return Boolean(this.records[id] && this.records[id].state !== 'active')
  }
}
