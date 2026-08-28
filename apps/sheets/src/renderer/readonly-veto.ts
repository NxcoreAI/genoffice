/**
 * Read-only embed (host passes readonly=1). Univer's own `disableEdit` config
 * only seals the edit *entry* (double-click / typing / F2 never mounts the
 * cell editor, the fx input is disabled) — direct-write gestures still reach
 * the command service: Delete clears the selection, the context menu inserts
 * rows, paste/autofill dispatch their commands. This veto closes those by
 * throwing in `beforeCommandExecuted` for the user-editing command set.
 *
 * GenOffice's own installs (viewport streaming, file loads, journal overlay
 * replays, chart/sparkline syncs) run through the SAME commands, so the veto
 * keys on `journalSuppression`: every programmatic install raises it (see
 * univer-sync.ts / App.tsx raise sites), user gestures never do. Pure view
 * commands (selection, sheet activation, zoom, scroll) are absent from the
 * set on purpose — a preview must keep those.
 */
import { CustomCommandExecutionError, ICommandService } from '@univerjs/core'

import { journalSuppression, type UniverRuntime } from './univer-state'

/// Mirrors the editing half of Univer's permission interceptor switch
/// (@univerjs/sheets SheetPermissionCheckController) plus the
/// filter/sort/number-format/clipboard companions that live in sibling
/// bundles. `sheet.command.set-range-values` and `set-style` are included:
/// GenOffice only dispatches them with journalSuppression raised.
const EDIT_COMMAND_IDS: ReadonlySet<string> = new Set([
  // cell content / styles
  'sheet.command.set-range-values',
  'sheet.command.set-style',
  'sheet.command.set-border',
  'sheet.command.set-border-basic',
  'sheet.command.set-border-color',
  'sheet.command.set-border-position',
  'sheet.command.set-border-style',
  'sheet.command.set-bold',
  'sheet.command.set-italic',
  'sheet.command.set-underline',
  'sheet.command.set-overline',
  'sheet.command.set-stroke',
  'sheet.command.set-font-family',
  'sheet.command.set-font-size',
  'sheet.command.set-text-color',
  'sheet.command.set-background-color',
  'sheet.command.reset-text-color',
  'sheet.command.reset-background-color',
  'sheet.command.set-horizontal-text-align',
  'sheet.command.set-vertical-text-align',
  'sheet.command.set-text-rotation',
  'sheet.command.set-text-wrap',
  'sheet.command.set-range-custom-metadata',
  'sheet.command.clear-selection-all',
  'sheet.command.clear-selection-content',
  'sheet.command.clear-selection-format',
  'sheet.command.auto-clear-content',
  // number formats
  'sheet.command.numfmt.set.numfmt',
  'sheet.command.numfmt.set.currency',
  'sheet.command.numfmt.set.percent',
  'sheet.command.numfmt.add.decimal.command',
  'sheet.command.numfmt.subtract.decimal.command',
  // rows / columns / ranges
  'sheet.command.insert-row',
  'sheet.command.insert-row-before',
  'sheet.command.insert-row-after',
  'sheet.command.insert-multi-rows-above',
  'sheet.command.insert-multi-rows-after',
  'sheet.command.insert-row-by-range',
  'sheet.command.insert-col',
  'sheet.command.insert-col-before',
  'sheet.command.insert-col-after',
  'sheet.command.insert-multi-cols-before',
  'sheet.command.insert-multi-cols-right',
  'sheet.command.insert-col-by-range',
  'sheet.command.remove-row',
  'sheet.command.remove-row-by-range',
  'sheet.command.remove-col',
  'sheet.command.remove-col-by-range',
  'sheet.command.move-rows',
  'sheet.command.move-cols',
  'sheet.command.move-range',
  'sheet.command.insert-range-move-down',
  'sheet.command.insert-range-move-right',
  'sheet.command.delete-range-move-left',
  'sheet.command.delete-range-move-up',
  'sheet.command.set-row-height',
  'sheet.command.delta-row-height',
  'sheet.command.set-col-data',
  'sheet.command.set-row-data',
  'sheet.command.set-worksheet-col-width',
  'sheet.command.set-col-is-auto-width',
  'sheet.command.set-row-is-auto-height',
  'sheet.command.set-col-hidden',
  'sheet.command.set-rows-hidden',
  'sheet.command.set-selected-cols-visible',
  'sheet.command.set-selected-rows-visible',
  'sheet.command.set-specific-cols-visible',
  'sheet.command.set-specific-rows-visible',
  'sheet.command.append-row',
  // merges
  'sheet.command.add-worksheet-merge',
  'sheet.command.add-worksheet-merge-all',
  'sheet.command.add-worksheet-merge-horizontal',
  'sheet.command.add-worksheet-merge-vertical',
  'sheet.command.remove-worksheet-merge',
  // fill / copy tools
  'sheet.command.auto-fill',
  'sheet.command.refill',
  'sheet.command.copy-down',
  'sheet.command.copy-right',
  'sheet.command.reorder-range',
  'sheet.command.split-text-to-columns',
  'sheet.command.text-to-number',
  'sheet.command.toggle-cell-checkbox',
  // worksheets / workbook structure
  'sheet.command.insert-sheet',
  'sheet.command.remove-sheet',
  'sheet.command.copy-sheet',
  'sheet.command.set-worksheet-name',
  'sheet.command.set-workbook-name',
  'sheet.command.set-worksheet-order',
  'sheet.command.set-worksheet-hidden',
  'sheet.command.set-worksheet-show',
  'sheet.command.set-worksheet-row-count',
  'sheet.command.set-worksheet-column-count',
  'sheet.command.set-worksheet-default-style',
  'sheet.command.set-worksheet-right-to-left',
  'sheet.command.set-tab-color',
  // defined names
  'sheet.command.insert-defined-name',
  'sheet.command.remove-defined-name',
  'sheet.command.set-defined-name',
  // protection (preview must not rewrite it)
  'sheet.command.set-protection',
  'sheet.command.set-worksheet-protection',
  'sheet.command.add-worksheet-protection',
  'sheet.command.delete-worksheet-protection',
  'sheet.command.set-worksheet-permission-points',
  'sheet.command.add-range-protection',
  'sheet.command.delete-range-protection',
  // range theme styles
  'sheet.command.register-worksheet-range-theme-style',
  'sheet.command.unregister-worksheet-range-theme-style',
  'sheet.command.set-worksheet-range-theme-style',
  // filter / sort (grid-header UI is user-reachable without the ribbon)
  'sheet.command.smart-toggle-filter',
  'sheet.command.set-filter-range',
  'sheet.command.set-filter-criteria',
  'sheet.command.clear-filter-criteria',
  'sheet.command.remove-sheet-filter',
  'sheet.command.re-calc-filter',
  'sheet.command.sort-range',
  // clipboard writes (copy stays available in a preview)
  'univer.command.cut',
  'univer.command.paste',
])

/// Installs the veto. No-op unless the host opened the view read-only.
export function installReadonlyVeto(runtime: UniverRuntime, enabled: boolean): void {
  if (!enabled) return
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  commandService.beforeCommandExecuted((command) => {
    if (!EDIT_COMMAND_IDS.has(command.id) || journalSuppression.active) return
    // Univer cleans its command-execution stack only on the success path, so
    // a veto would strand the entry pushed for this dispatch — and every
    // later mutation findLast-scans that stack. The hook gets the very object
    // that was pushed; remove it by identity here, synchronously, before the
    // throw (same reasoning as calc-options.ts).
    const stack = (commandService as unknown as { _commandExecutionStack?: unknown[] })
      ._commandExecutionStack
    const index = stack?.indexOf(command) ?? -1
    if (index >= 0) stack?.splice(index, 1)
    throw new CustomCommandExecutionError('read-only preview')
  })
}
