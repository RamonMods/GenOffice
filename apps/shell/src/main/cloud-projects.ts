import { unlinkSync } from 'node:fs'
import type { CloudProjectEntry, CloudProjectKind, CloudProjectsSnapshot } from '../shared/home-api'

const KINDS: readonly CloudProjectKind[] = ['docs', 'sheets', 'slides']

/** 'slides_agent_git' / 'docs_agent' / 'sheets_agent_new' → module kind */
export function kindFromType(type: string): CloudProjectKind | 'other' {
  return KINDS.find((k) => type.startsWith(k)) ?? 'other'
}

export function cloudStoreOwner(): string {
  return ''
}

export function clearCloudProjectsStore(storePath: string): void {
  try {
    unlinkSync(storePath)
  } catch {
    // nothing cached; fine
  }
}

/** Cloud projects are disabled in this build; return unavailable gracefully. */
export function readCloudProjectsStore(_storePath: string): CloudProjectsSnapshot | null {
  return { available: false, projects: [], syncedAt: 0 }
}

/** Cloud projects are disabled in this build; no remote sync. */
export function syncCloudProjects(_storePath: string): Promise<CloudProjectsSnapshot> {
  return Promise.resolve({ available: false, projects: [], syncedAt: 0 })
}

/** Cloud projects are disabled in this build; never open external project URLs. */
export function cloudProjectExternalUrl(_projectUrl: unknown): string | null {
  return null
}

// Keep the type reference live for compilation in isolated module checks.
void (null as unknown as CloudProjectEntry | null)
