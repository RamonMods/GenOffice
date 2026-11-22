import { flagBool, flagString, type ParsedArgs } from './args'
import type { OpFailure } from './op-errors'
import { CliError, EXIT, type CommandResult } from './result'

/**
 * atomic: any rejected op leaves the file untouched (exit 1).
 * best_effort: every op that can apply does; the rest are reported.
 * stop_on_error: ops apply in order up to the first rejection.
 * Both non-atomic modes write the file and answer `status: "partial"` when an op was lost.
 */
export type BatchMode = 'atomic' | 'best_effort' | 'stop_on_error'

export interface BatchCounts {
  total: number
  applied: number
  failed: number
  skipped: number
}

export const BATCH_OPTIONS = [
  {
    name: 'best-effort',
    description:
      'apply: aplique todas as operações que podem ser aplicadas e grave; as operações rejeitadas voltam em detail.failures com status "partial"',
  },
  {
    name: 'stop-on-error',
    description:
      'apply: aplique as operações em ordem até a primeira rejeição e grave o que foi aplicado (status "partial")',
  },
]

export function batchMode(args: ParsedArgs): BatchMode {
  const best = flagBool(args, 'best-effort')
  const stop = flagBool(args, 'stop-on-error')
  // `--isolation per_op` predates --best-effort on `slides apply` and means the same thing
  const isolation = flagString(args, 'isolation')
  if (isolation !== undefined && isolation !== 'atomic' && isolation !== 'per_op') {
    throw new CliError(EXIT.usage, '--isolation deve ser atomic ou per_op', undefined, {
      reason: 'invalid_argument',
    })
  }
  if (stop && (best || isolation === 'per_op')) {
    throw new CliError(
      EXIT.usage,
      '--stop-on-error não pode ser combinado com --best-effort',
      undefined,
      { reason: 'invalid_argument', suggestion: 'escolha um modo de lote' },
    )
  }
  if (best || isolation === 'per_op') return 'best_effort'
  if (stop) return 'stop_on_error'
  return 'atomic'
}

export function batchCounts(total: number, applied: number, failed: number): BatchCounts {
  return { total, applied, failed, skipped: Math.max(0, total - applied - failed) }
}

/** The op a rejection is pinned to, when the executor reported one. */
export function pinnedFailure(err: unknown): OpFailure | undefined {
  if (!(err instanceof CliError)) return undefined
  const failures = err.detail?.failures
  return Array.isArray(failures) && failures.length ? (failures[0] as OpFailure) : undefined
}

/** A result that ran as a batch: counts always, `partial` once any op was lost. */
export function batchResult(
  r: CommandResult,
  counts: BatchCounts,
  failures: OpFailure[],
): CommandResult {
  const detail = { ...r.detail, batch: counts, ...(failures.length ? { failures } : {}) }
  if (!failures.length) return { ...r, detail }
  const skipped = counts.skipped ? `, ${counts.skipped} ignorada(s)` : ''
  return {
    ...r,
    status: 'partial',
    summary: `${r.summary}; ${counts.failed} falha(s)${skipped}`,
    detail,
  }
}

/** The rejection that stopped a batch before anything was written, with its counts. */
export function failedBatch(err: CliError, total: number): CliError {
  const failed = (err.detail?.failures as unknown[] | undefined)?.length ?? 1
  return new CliError(
    err.code,
    err.message,
    { ...err.detail, batch: batchCounts(total, 0, failed) },
    { reason: err.reason, suggestion: err.suggestion },
  )
}
