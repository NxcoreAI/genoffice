import type { Editor } from '@tiptap/core'
import type {
  Block,
  DocProtection,
  HfImage,
  SectionInfo,
  SectionSettings,
  SourceInfo,
  ThemeColors,
  ThemeFonts,
  WriteProtection,
} from '@genoffice/docx-engine'
import type { AiHeaderFooterAccess } from '../tools'
import type { ReviewContext } from '../../review-actions'
import type { NumberingContext } from '../../numbering-actions'
import type { InkAnnotation } from '../../editor/ink'

/**
 * Structured operation layer for the host agent (`apply_operations`).
 *
 * `apply_commands` covers rich text / paragraph formatting, but tables need
 * structural edits (rows, columns, merges, cell properties) that a field-mask
 * command cannot express; page/section/note/field work touches App state
 * outside ProseMirror. These ops address a table by its block index and a cell
 * by 0-based row/col (table ops), or a document block / section index
 * (page/section/notes/fields), and reuse the real ribbon helpers.
 *
 * Every op is an isolated unit: one op's failure (or exception) is reported
 * with its index and never leaves a partial edit behind, and it never stops
 * the remaining ops in the batch (per-op isolation, unlike atomic batches).
 */

/**
 * App-state half of the operation context (B2 page/section/notes/fields ops).
 * App.tsx passes its live `agentDocRefs.current` bundle — the same setters the
 * ribbon uses — so an op runs the real pipeline instead of a reimplementation.
 * Pure ProseMirror ops (tables, page.break/page.blank/page.cover) work without
 * it; ops that need it fail with a clear message when it is absent.
 */
export interface AgentDocumentContext {
  /** per-section geometry state (App's sections[]) */
  sections: SectionInfo[]
  /** final-section settings (SaveOptions.section; app state) */
  section: SectionSettings | null
  /** section the caret currently sits in */
  activeSection: number
  setSection: (settings: SectionSettings) => void
  setSections: (updater: (prev: SectionInfo[]) => SectionInfo[]) => void
  setSectionDirty: (dirty: boolean) => void
  setSectionsDirty: (updater: (prev: number[]) => number[]) => void
  pageColor: string | null
  setPageColor: (color: string | null) => void
  setPageColorDirty: (dirty: boolean) => void
  themeFonts: ThemeFonts | null
  setThemeFonts: (fonts: ThemeFonts) => void
  setThemeFontsDirty: (dirty: boolean) => void
  themeColors: ThemeColors | null
  setThemeColors: (colors: ThemeColors) => void
  setThemeColorsDirty: (dirty: boolean) => void
  /**
   * App's real applyPageSetup pipeline (validates + writes section geometry /
   * page color / themes, including the final-section SaveOptions path).
   * `dryRun` validates and reports without changing any state.
   */
  applyPageSetup: (
    input: Record<string, unknown>,
    sectionIndex?: number,
    dryRun?: boolean,
  ) => { ok: true; output: string; summary: string } | { ok: false; error: string }
  /** App's real insertSectionBreak (inserts the sectPr paragraph after `afterBlockIndex`, or at the caret) */
  insertSectionBreak: (type: SectionInfo['startType'], afterBlockIndex?: number) => void
  /** page-number format state: final-section pgNumType + non-final dirty section indexes */
  pgNumEdit: { fmt?: string; start?: number } | null
  setPgNumEdit: (value: { fmt?: string; start?: number } | null) => void
  pgNumDirtySections: number[]
  setPgNumDirtySections: (updater: (prev: number[]) => number[]) => void
  /** App's real inline-field insert (DATE/TIME/PAGE/NUMPAGES/FILENAME) at the caret */
  insertField: (instr: string) => void
  /** App's real F9 field-cache recompute */
  updateFields: () => void
  /** live review-actions context (footnote/endnote lists + setters/actions) — read through the accessor so a batch sees each op's writes */
  review: () => ReviewContext
  /** live numbering-actions context (pending numbering definitions, numId allocation) for list restart/continue/multilevel ops */
  numbering: () => NumberingContext
  // ── B4 references / review / protection ──
  /** live bibliography sources (add via addSource; each add is flushed so later ops in the batch see it) */
  getSources: () => SourceInfo[]
  addSource: (source: SourceInfo) => void
  /** live document blocks (original parse; caption SEQ numbering reads their XML) */
  getBlocks: () => Block[]
  /** track-changes toggle (App state; the tracker storage flips after a flushed render) */
  setTrackChanges: (on: boolean) => void
  /** App's Protect dialog apply path (write/open protection, dirty flags, unlock state) */
  applyProtection: (result: AgentProtectResult) => Promise<void>
  // ── A3: header/footer images + ink ──
  /** parsed header/footer images + pending additions + pending removals (indexes) */
  getHeaderFooterImages: () => {
    header: HfImage[]
    footer: HfImage[]
    headerRemove: number[]
    footerRemove: number[]
  }
  /** append a NEW image to the header/footer state (marks it dirty); returns its documentState index */
  addHeaderFooterImage: (kind: 'header' | 'footer', image: HfImage) => { index: number }
  /** remove an image by documentState index (parsed → recorded removal, pending → dropped); false when out of range */
  removeHeaderFooterImage: (kind: 'header' | 'footer', index: number) => boolean
  /** append a freehand ink annotation (review-actions addInk: state + dirty flag) */
  addInk: (annotation: InkAnnotation) => void
  /** DOM metrics for a top-level block (100% zoom, overlay-relative): needed by object.ink.mark */
  inkMeasureBlock: (blockIndex: number) => {
    anchorIndex: number
    /** anchor paragraph top-left relative to the ink overlay, 100% zoom */
    anchor: { x: number; y: number }
    /** block paragraph rect relative to the ink overlay, 100% zoom */
    rect: { x: number; y: number; w: number; h: number }
    /** rendered text-column width in px (100% zoom) */
    columnWidthPx: number
  } | null
  // ── A3: character-style creation (styles.xml write-back via pending styleUpserts) ──
  /** parsed character/paragraph style ids + pending upsert ids (collision check) */
  getCharacterStyleIds: () => string[]
  /** queue a style definition for the next save (surgical styles.xml upsert) */
  addStyleUpsert: (upsert: import('@genoffice/docx-engine').StyleUpsert) => void
  /** App's docs:set-password path (open password intent; applied on the next save) */
  setOpenPassword: (password: string | null) => Promise<{ ok: boolean; error?: string }>
  /** TOC heading page lookup (kept for page/layout ops that need measured pages) */
  headingPages?: () => number[] | null
  /** header/footer access (kept for future header/footer ops) */
  headerFooter?: () => AiHeaderFooterAccess | null
}

/** Input passed to every operation handler. */
export interface OperationContext {
  editor: Editor
  /** Width in CSS px of the current text column (page width − margins, split by columns). */
  contentWidthPx: number
  /** Validate/describe only: handlers must not mutate the document. */
  dryRun: boolean
  /** App-state access for page/section/notes/fields ops; null when the host did not pass it. */
  app: AgentDocumentContext | null
}

/** Outcome of one operation. */
export interface OperationOutcome {
  ok: boolean
  /** One-line human/model readable result. */
  output: string
  /** Short label for host records (e.g. "insert table row"). */
  summary?: string
  /** Present when ok === false. */
  error?: string
  /** Structured details for the host (ids, counts, resolved addressing). */
  record?: Record<string, unknown>
  /** dryRun: the op validated and would apply, but nothing was touched. */
  planned?: boolean
}

/**
 * B4 protection input (same shape as the Protect dialog result): password
 * hashing happens inside the op before this is called, and the password itself
 * is never logged, echoed, or put into an op result.
 */
export interface AgentProtectResult {
  openPassword?: string | null
  writeProtection?: WriteProtection | null
  protection?: DocProtection | null
  removePersonalInfo?: boolean
}

export type OperationHandler = (
  input: Record<string, unknown>,
  ctx: OperationContext,
) => OperationOutcome | Promise<OperationOutcome>

/** One registry entry: vocabulary entry + input schema + executor. */
export interface OperationDef {
  /** Dotted op name, e.g. "table.insertRow". */
  name: string
  /** Vocabulary group ("table" for B1). */
  group: string
  /** One-line usage summary; surfaced through operationsGuide and the tool schema. */
  description: string
  /** JSON schema for the op payload (without the `op` key). */
  inputSchema: Record<string, unknown>
  handler: OperationHandler
}

/** One row in the batch result the host/model receives. */
export interface AppliedOperation {
  /** Index of the op in the caller's `ops` array. */
  index: number
  /** Op name as submitted ('' when malformed). */
  op: string
  ok: boolean
  output: string
  summary?: string
  error?: string
  record?: Record<string, unknown>
  planned?: boolean
}

/** Result of one `apply_operations` batch. */
export interface OperationBatchResult {
  /** true when every submitted op applied (or validated, for dryRun). */
  ok: boolean
  dryRun: boolean
  /** Ops that mutated the document (0 for dryRun). */
  applied: number
  failed: number
  results: AppliedOperation[]
  /** Compact text block: one line per op, safe to return to the model as-is. */
  output: string
}