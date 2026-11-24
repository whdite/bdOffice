/** Stable user-facing categories; raw provider responses stay out of task metadata. */
export function taskFailureReason(error: unknown): string {
  const message = error instanceof Error ? error.message : typeof error === 'string' ? error : ''
  if (message === 'result_not_verified') return message
  if (/^task_[a-z_]{1,50}$/.test(message)) return message
  if (
    /fetch failed|failed to fetch|network(?: request)? (?:failed|error)|ECONNRESET|ECONNREFUSED|ENOTFOUND|ETIMEDOUT|UND_ERR_(?:CONNECT|HEADERS)_TIMEOUT/i.test(
      message,
    )
  )
    return 'task_provider_transport_failed'
  return 'execution_failed'
}
