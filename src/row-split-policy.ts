type Policy = 'allow' | 'avoid';
/** Keep a direct setting separate from the inherited DOCX rendering policy.
 * Native/HWP rows with neither attribute retain their existing atomic flow. */
export function rowSplitPolicy(attrs: Record<string, unknown> | null | undefined): Policy | null {
  for (const value of [attrs?.rowSplit, attrs?.wordRowSplit])
    if (value === 'allow' || value === 'avoid') return value;
  return null;
}
/** Styles are flattened on export. Materialize inherited avoidance and native
 * atomic rows, but preserve an omitted Word default instead of inventing a
 * direct false override that could mask later external table-style edits. */
export function exportedRowCantSplit(attrs: Record<string, unknown> | null | undefined, continuingMerge: boolean): boolean | undefined {
  if (attrs?.rowSplit === 'allow' || attrs?.rowSplit === 'avoid') return attrs.rowSplit === 'avoid';
  if (attrs?.wordRowSplit === 'allow') return undefined;
  if (attrs?.wordRowSplit === 'avoid') return true;
  return !continuingMerge;
}
