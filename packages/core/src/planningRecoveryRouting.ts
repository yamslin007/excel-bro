/** Only caller-owned task IDs enter semantic routing. The model never grants continuation. */
export function planningRecoveryCandidates(ids: readonly (string | null | undefined)[]): string[] {
  return [...new Set(ids.filter((id): id is string => Boolean(id?.trim())))].slice(-6);
}

export function isOfferedPlanningRecovery(turnId: string, candidates: readonly string[]): boolean {
  return candidates.includes(turnId);
}
