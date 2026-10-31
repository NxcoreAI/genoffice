import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type {
  Block,
  NoteInfo,
  SectionInfo,
  SectionSettings,
  SourceInfo,
  ThemeColors,
  ThemeFonts,
} from '@genoffice/docx-engine'
import type { NumberingContext } from '../src/renderer/numbering-actions'
import type { DocState, PendingNumbering } from '../src/renderer/doc-state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import type { ReviewContext } from '../src/renderer/review-actions'
import { collectRevisions } from '../src/renderer/editor/revisions'
import {
  MAX_OPERATIONS,
  OPERATIONS_GUIDE,
  contentWidthPxFromSection,
  describeOperations,
  runOperations,
  type AgentDocumentContext,
  type AgentProtectResult,
  type OperationBatchResult,
} from '../src/renderer/ai/operations'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const editors: Editor[] = []

afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

const para = (text = ''): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: text ? [{ type: 'text', text }] : [],
})

const cell = (text: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docTableCell',
  attrs: { colspan: 1, rowspan: 1, ...attrs },
  content: [para(text)],
})

const headerCell = (text: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docTableHeader',
  attrs: { colspan: 1, rowspan: 1, ...attrs },
  content: [para(text)],
})

const row = (cells: JsonNode[], attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docTableRow',
  attrs: { heightTwips: null, heightRule: null, repeatHeader: false, ...attrs },
  content: cells,
})

const table = (rows: JsonNode[][], attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docTable',
  attrs: { docxIndex: null, colWidthsPct: [100 / 3, 100 / 3, 100 / 3], ...attrs },
  content: rows.map((cells) => row(cells)),
})

/** 0 paragraph | 1 table (3×3, header row) | 2 paragraph */
function fixture3x3(): JsonNode[] {
  return [
    para('intro'),
    table([
      [headerCell('A'), headerCell('B'), headerCell('C')],
      [cell('D'), cell('E'), cell('F')],
      [cell('G'), cell('H'), cell('I')],
    ]),
    para('outro'),
  ]
}

const TABLE_BLOCK = 1

async function run(
  editor: Editor,
  ops: unknown,
  options: { dryRun?: boolean; app?: AgentDocumentContext | null } = {},
): Promise<OperationBatchResult> {
  return runOperations(editor, ops, { contentWidthPx: 600, ...options })
}

// ── fake App-state context (B2 page/section/notes/fields) ────────────────────

interface FakeAppState {
  sections: SectionInfo[]
  section: SectionSettings | null
  activeSection: number
  sectionDirty: boolean
  sectionsDirty: number[]
  pageColor: string | null
  pageColorDirty: boolean
  themeFonts: ThemeFonts | null
  themeFontsDirty: boolean
  themeColors: ThemeColors | null
  themeColorsDirty: boolean
  pgNumEdit: { fmt?: string; start?: number } | null
  pgNumDirtySections: number[]
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  notesDirty: boolean
  pageSetupCalls: Array<{ input: Record<string, unknown>; sectionIndex?: number; dryRun: boolean }>
  sectionBreaks: Array<{ type: string; afterBlockIndex?: number }>
  fieldInsertions: string[]
  fieldUpdates: number
  pendingNumbering: PendingNumbering
  sources: SourceInfo[]
  trackChanges: boolean
  protectionRequests: AgentProtectResult[]
  passwordRequests: Array<string | null>
}

function fakeSection(index: number): SectionInfo {
  return {
    settings: {
      pageWidth: 11906,
      pageHeight: 16838,
      orientation: 'portrait',
      marginTop: 1440,
      marginRight: 1440,
      marginBottom: 1440,
      marginLeft: 1440,
      pageBorder: false,
      columns: 1,
      colSpace: 720,
    } as SectionSettings,
    startType: 'nextPage',
    firstBlockIndex: index,
    lastBlockIndex: index + 1,
    sectPrXml: '<w:sectPr/>',
    titlePg: false,
    headerRefs: {},
    footerRefs: {},
  } as unknown as SectionInfo
}

/** App-state double for B2 ops: real editor + mutable fake app state + the real
 *  review-actions pipeline (so note list/marker semantics are genuinely tested). */
function makeFakeApp(
  editor: Editor,
  opts: {
    sectionCount?: number
    activeSection?: number
    applyPageSetup?: AgentDocumentContext['applyPageSetup']
  } = {},
): { ctx: AgentDocumentContext; state: FakeAppState } {
  const count = opts.sectionCount ?? 2
  const state: FakeAppState = {
    sections: Array.from({ length: count }, (_v, index) => fakeSection(index)),
    section: fakeSection(count - 1).settings,
    activeSection: opts.activeSection ?? 0,
    sectionDirty: false,
    sectionsDirty: [],
    pageColor: null,
    pageColorDirty: false,
    themeFonts: null,
    themeFontsDirty: false,
    themeColors: null,
    themeColorsDirty: false,
    pgNumEdit: null,
    pgNumDirtySections: [],
    footnotes: [],
    endnotes: [],
    notesDirty: false,
    pageSetupCalls: [],
    sectionBreaks: [],
    fieldInsertions: [],
    fieldUpdates: 0,
    pendingNumbering: { newDefs: [], restartNums: [] },
    sources: [],
    trackChanges: false,
    protectionRequests: [],
    passwordRequests: [],
  }
  // minimal DocState for the numbering actions (parsed numbering definitions)
  const fakeDoc = {
    parsed: {
      numbering: new Map([
        [
          '5',
          {
            numId: '5',
            abstractNumId: '1',
            levels: {
              0: { numFmt: 'decimal', lvlText: '%1.', start: 1, indentLeft: 720, hanging: 360 },
            },
            startOverrides: {},
          },
        ],
      ]),
    },
  } as unknown as DocState
  const numbering = (): NumberingContext => ({
    editor,
    doc: fakeDoc,
    pendingNumbering: state.pendingNumbering,
    setPendingNumbering: (update) => {
      state.pendingNumbering = update(state.pendingNumbering)
    },
    numIdFloorRef: { current: 0 },
    setStatus: () => {},
  })
  const review = (): ReviewContext => ({
    editor,
    doc: null,
    dirtyRef: { current: false },
    setStatus: () => {},
    notePrompt: null,
    setNotePrompt: () => {},
    get footnotes() {
      return state.footnotes
    },
    get endnotes() {
      return state.endnotes
    },
    setFootnotes: (updater) => {
      state.footnotes =
        typeof updater === 'function'
          ? (updater as (prev: NoteInfo[]) => NoteInfo[])(state.footnotes)
          : updater
    },
    setEndnotes: (updater) => {
      state.endnotes =
        typeof updater === 'function'
          ? (updater as (prev: NoteInfo[]) => NoteInfo[])(state.endnotes)
          : updater
    },
    setNotesDirty: (dirty) => {
      state.notesDirty = dirty
    },
    comments: [],
    setComments: () => {},
    setCommentsDirty: () => {},
    setCommentComposing: () => {},
    setShowComments: () => {},
    setInkAnnotations: () => {},
    setInksDirty: () => {},
    setCompareResult: () => {},
  })
  const defaultApplyPageSetup: AgentDocumentContext['applyPageSetup'] = (
    input,
    sectionIndex,
    dryRun,
  ) => {
    state.pageSetupCalls.push({ input, sectionIndex, dryRun: dryRun === true })
    if (dryRun) return { ok: true, output: 'Would apply: test setup', summary: 'page setup' }
    const index = sectionIndex ?? state.activeSection
    if (input.pageColor !== undefined) {
      state.pageColor = input.pageColor as string | null
      state.pageColorDirty = true
    }
    const sec = state.sections[index]
    if (sec && input.margins && typeof input.margins === 'object') {
      const margins = input.margins as Record<string, number>
      const settings = { ...sec.settings }
      if (margins.top !== undefined) settings.marginTop = margins.top
      if (margins.left !== undefined) settings.marginLeft = margins.left
      state.sections = state.sections.map((item, i) =>
        i === index ? { ...item, settings } : item,
      )
    }
    if (index >= state.sections.length - 1) {
      state.section = state.sections[index].settings
      state.sectionDirty = true
    } else {
      state.sectionsDirty = [...new Set([...state.sectionsDirty, index])]
    }
    return { ok: true, output: 'Applied: test setup', summary: 'page setup' }
  }
  const ctx: AgentDocumentContext = {
    get sections() {
      return state.sections
    },
    get section() {
      return state.section
    },
    get activeSection() {
      return state.activeSection
    },
    setSection: (settings) => {
      state.section = settings
      state.sectionDirty = true
    },
    setSections: (updater) => {
      state.sections = updater(state.sections)
    },
    setSectionDirty: (dirty) => {
      state.sectionDirty = dirty
    },
    setSectionsDirty: (updater) => {
      state.sectionsDirty = updater(state.sectionsDirty)
    },
    get pageColor() {
      return state.pageColor
    },
    setPageColor: (color) => {
      state.pageColor = color
    },
    setPageColorDirty: (dirty) => {
      state.pageColorDirty = dirty
    },
    get themeFonts() {
      return state.themeFonts
    },
    setThemeFonts: (fonts) => {
      state.themeFonts = fonts
    },
    setThemeFontsDirty: (dirty) => {
      state.themeFontsDirty = dirty
    },
    get themeColors() {
      return state.themeColors
    },
    setThemeColors: (colors) => {
      state.themeColors = colors
    },
    setThemeColorsDirty: (dirty) => {
      state.themeColorsDirty = dirty
    },
    applyPageSetup: opts.applyPageSetup ?? defaultApplyPageSetup,
    insertSectionBreak: (type, afterBlockIndex) => {
      state.sectionBreaks.push({ type, afterBlockIndex })
      const doc = editor.state.doc
      const pos =
        afterBlockIndex === undefined
          ? editor.state.selection.from
          : (() => {
              let p = 0
              for (let i = 0; i <= afterBlockIndex; i++) p += doc.child(i).nodeSize
              return p
            })()
      editor
        .chain()
        .insertContentAt(pos, {
          type: 'docProtected',
          attrs: {
            docxIndex: null,
            blockType: 'passthrough',
            label: 'Section break paragraph',
            previewText: '',
            genXml: '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>',
          },
        })
        .run()
    },
    get pgNumEdit() {
      return state.pgNumEdit
    },
    setPgNumEdit: (value) => {
      state.pgNumEdit = value
    },
    get pgNumDirtySections() {
      return state.pgNumDirtySections
    },
    setPgNumDirtySections: (updater) => {
      state.pgNumDirtySections = updater(state.pgNumDirtySections)
    },
    insertField: (instr) => {
      state.fieldInsertions.push(instr)
      const value = instr === 'FILENAME' ? 'test.docx' : '1'
      editor
        .chain()
        .focus()
        .insertContent({
          type: 'text',
          text: value,
          marks: [{ type: 'instrField', attrs: { instr } }],
        })
        .unsetMark('instrField')
        .run()
    },
    updateFields: () => {
      state.fieldUpdates += 1
    },
    review,
    headerFooter: () => null,
    numbering,
    getSources: () => state.sources,
    addSource: (source) => {
      state.sources = [...state.sources, source]
    },
    getBlocks: () => [] as Block[],
    setTrackChanges: (on) => {
      state.trackChanges = on
      const storage = editor.storage as unknown as { trackChanges?: { enabled?: boolean } }
      if (storage.trackChanges) storage.trackChanges.enabled = on
    },
    applyProtection: async (result) => {
      state.protectionRequests.push(result)
    },
    setOpenPassword: async (password) => {
      state.passwordRequests.push(password)
      return { ok: true }
    },
  }
  return { ctx, state }
}

/** Simple 3-paragraph document for page/section/notes/fields tests. */
function paragraphsDoc(): JsonNode[] {
  return [para('first'), para('second'), para('third')]
}


function jsonTable(editor: Editor, blockIndex = TABLE_BLOCK): JsonNode {
  return editor.getJSON().content![blockIndex] as unknown as JsonNode
}

function rowTexts(tableNode: JsonNode, rowIndex: number): string[] {
  const cells = tableNode.content![rowIndex].content ?? []
  return cells.map((c) => (c.content ?? []).map((p) => (p.content ?? []).map((t) => t.text ?? '').join('')).join(''))
}

function allCellTexts(tableNode: JsonNode): string[][] {
  return tableNode.content!.map((_row, index) => rowTexts(tableNode, index))
}

function cellAt(tableNode: JsonNode, rowIndex: number, cellIndex: number): JsonNode {
  return tableNode.content![rowIndex].content![cellIndex]
}

function snapshot(editor: Editor): string {
  return JSON.stringify(editor.getJSON())
}

describe('operations registry', () => {
  it('registers every B1 table op with an object schema and guide entry', async () => {
    const defs = describeOperations()
    const names = defs.map((def) => def.name)
    expect(names).toEqual(
      expect.arrayContaining([
        'table.insert',
        'table.delete',
        'table.insertRow',
        'table.deleteRow',
        'table.insertColumn',
        'table.deleteColumn',
        'table.mergeCells',
        'table.splitCell',
        'table.setCell',
        'table.setProperties',
        'table.autoFit',
      ]),
    )
    expect(defs.length).toBeGreaterThanOrEqual(11)
    for (const def of defs) {
      expect(['table', 'page', 'section', 'notes', 'field', 'format', 'object', 'reference', 'review']).toContain(def.group)
      expect(def.inputSchema.type).toBe('object')
      expect(OPERATIONS_GUIDE).toContain(def.name)
    }
    expect(OPERATIONS_GUIDE).toContain('apply_operations')
    expect(OPERATIONS_GUIDE).toContain('0-based row/col')
  })

  it('rejects malformed batches and empty ops', async () => {
    const editor = createEditor(fixture3x3())
    expect((await run(editor, 'nope')).ok).toBe(false)
    expect((await run(editor, [])).ok).toBe(false)
  })

  it('reports an unknown op with its index without touching the document', async () => {
    const editor = createEditor(fixture3x3())
    const before = snapshot(editor)
    const outcome = await run(editor, [
      { op: 'table.setCell', blockIndex: TABLE_BLOCK, row: 1, col: 1, text: 'X' },
      { op: 'table.explode' },
      { op: 'table.setCell', blockIndex: TABLE_BLOCK, row: 2, col: 2, text: 'Y' },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.applied).toBe(2)
    expect(outcome.failed).toBe(1)
    expect(outcome.results[1]).toMatchObject({ index: 1, op: 'table.explode', ok: false })
    expect(outcome.results[1].error).toContain('unknown operation')
    expect(outcome.output).toContain('[1] table.explode: FAILED')
    expect(before).not.toBe(snapshot(editor))
    expect(rowTexts(jsonTable(editor), 1)).toEqual(['D', 'X', 'F'])
    expect(rowTexts(jsonTable(editor), 2)).toEqual(['G', 'H', 'Y'])
  })
})

describe('table.insert / table.delete', () => {
  it('inserts a table at the end of the document with data and a header row', async () => {
    const editor = createEditor([para('intro')])
    const outcome = await run(editor, [
      {
        op: 'table.insert',
        rows: 2,
        cols: 2,
        data: [
          ['Name', 'Qty'],
          ['Widget', '3'],
        ],
        headerRow: true,
      },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.applied).toBe(1)
    const doc = editor.getJSON()
    expect(doc.content).toHaveLength(2)
    const inserted = doc.content![1] as unknown as JsonNode
    expect(inserted.type).toBe('docTable')
    expect(inserted.content).toHaveLength(2)
    expect(cellAt(inserted, 0, 0).type).toBe('docTableHeader')
    expect(allCellTexts(inserted)).toEqual([
      ['Name', 'Qty'],
      ['Widget', '3'],
    ])
    // default Word grid borders so a fresh table is visible before reload
    expect(inserted.attrs?.borders).toMatchObject({ top: { style: 'single' } })
  })

  it('inserts after a block index and deletes the table block', async () => {
    const editor = createEditor(fixture3x3())
    const inserted = await run(editor, [
      { op: 'table.insert', afterBlockIndex: 0, rows: 1, cols: 1, headerRow: false },
    ])
    expect(inserted.ok).toBe(true)
    expect(inserted.results[0].record).toMatchObject({ blockIndex: 1, rows: 1, cols: 1 })
    expect(editor.getJSON().content).toHaveLength(4)
    expect(editor.getJSON().content![1].type).toBe('docTable')

    // the 3×3 fixture table shifted to block 2; delete it
    const deleted = await run(editor, [{ op: 'table.delete', blockIndex: 2 }])
    expect(deleted.ok).toBe(true)
    expect(editor.getJSON().content!.map((node) => node.type)).toEqual([
      'docParagraph',
      'docTable',
      'docParagraph',
    ])
  })
})

describe('row ops', () => {
  it('insertRow appends a row by default and can insert before a reference row', async () => {
    const editor = createEditor(fixture3x3())
    const appended = await run(editor, [{ op: 'table.insertRow', blockIndex: TABLE_BLOCK }])
    expect(appended.ok).toBe(true)
    let tableNode = jsonTable(editor)
    expect(tableNode.content).toHaveLength(4)
    expect(rowTexts(tableNode, 3)).toEqual(['', '', ''])

    const before = await run(editor, [
      { op: 'table.insertRow', blockIndex: TABLE_BLOCK, row: 1, position: 'before' },
    ])
    expect(before.ok).toBe(true)
    tableNode = jsonTable(editor)
    expect(tableNode.content).toHaveLength(5)
    expect(rowTexts(tableNode, 1)).toEqual(['', '', ''])
    expect(rowTexts(tableNode, 2)).toEqual(['D', 'E', 'F'])
  })

  it('deleteRow removes the target grid row', async () => {
    const editor = createEditor(fixture3x3())
    const outcome = await run(editor, [{ op: 'table.deleteRow', blockIndex: TABLE_BLOCK, row: 1 }])
    expect(outcome.ok).toBe(true)
    expect(jsonTable(editor).content).toHaveLength(2)
    expect(allCellTexts(jsonTable(editor))).toEqual([
      ['A', 'B', 'C'],
      ['G', 'H', 'I'],
    ])
  })

  it('refuses to delete the only row', async () => {
    const editor = createEditor([
      para('intro'),
      table([[headerCell('A'), headerCell('B')], [cell('C'), cell('D')]]),
    ])
    await run(editor, [{ op: 'table.deleteRow', blockIndex: TABLE_BLOCK, row: 1 }])
    const failed = await run(editor, [{ op: 'table.deleteRow', blockIndex: TABLE_BLOCK }])
    expect(failed.ok).toBe(false)
    expect(failed.failed).toBe(1)
    expect(failed.results[0].error).toContain('row must be an integer')
  })
})

describe('column ops', () => {
  it('insertColumn appends by default and can insert before a reference column', async () => {
    const editor = createEditor(fixture3x3())
    const appended = await run(editor, [{ op: 'table.insertColumn', blockIndex: TABLE_BLOCK }])
    expect(appended.ok).toBe(true)
    let tableNode = jsonTable(editor)
    expect(tableNode.content![0].content).toHaveLength(4)
    expect(rowTexts(tableNode, 0)).toEqual(['A', 'B', 'C', ''])

    const before = await run(editor, [
      { op: 'table.insertColumn', blockIndex: TABLE_BLOCK, col: 0, position: 'before' },
    ])
    expect(before.ok).toBe(true)
    tableNode = jsonTable(editor)
    expect(rowTexts(tableNode, 0)).toEqual(['', 'A', 'B', 'C', ''])
  })

  it('deleteColumn removes the target grid column', async () => {
    const editor = createEditor(fixture3x3())
    const outcome = await run(editor, [{ op: 'table.deleteColumn', blockIndex: TABLE_BLOCK, col: 1 }])
    expect(outcome.ok).toBe(true)
    expect(allCellTexts(jsonTable(editor))).toEqual([
      ['A', 'C'],
      ['D', 'F'],
      ['G', 'I'],
    ])
  })
})

describe('merge / split', () => {
  it('mergeCells merges a rectangular range into the top-left cell', async () => {
    const editor = createEditor(fixture3x3())
    const outcome = await run(editor, [
      { op: 'table.mergeCells', blockIndex: TABLE_BLOCK, row: 0, col: 0, rowspan: 2, colspan: 2 },
    ])
    expect(outcome.ok).toBe(true)
    let tableNode = jsonTable(editor)
    const merged = cellAt(tableNode, 0, 0)
    expect(merged.attrs).toMatchObject({ colspan: 2, rowspan: 2 })
    // mergeCells concatenates the merged cells' content into the top-left cell
    expect(rowTexts(tableNode, 0)).toEqual(['ABDE', 'C'])
    expect(rowTexts(tableNode, 1)).toEqual(['F'])
    expect(rowTexts(tableNode, 2)).toEqual(['G', 'H', 'I'])

    const split = await run(editor, [
      { op: 'table.splitCell', blockIndex: TABLE_BLOCK, row: 0, col: 0 },
    ])
    expect(split.ok).toBe(true)
    tableNode = jsonTable(editor)
    expect(cellAt(tableNode, 0, 0).attrs).toMatchObject({ colspan: 1, rowspan: 1 })
    // the split restores the 3×3 grid; the merged text stays in the top-left cell
    expect(tableNode.content).toHaveLength(3)
    for (const tableRow of tableNode.content!) {
      expect(tableRow.content).toHaveLength(3)
    }
    expect(rowTexts(tableNode, 0)[0]).toBe('ABDE')
    expect(rowTexts(tableNode, 2)).toEqual(['G', 'H', 'I'])
  })

  it('splitCell reports a non-merged cell as a per-op failure', async () => {
    const editor = createEditor(fixture3x3())
    const before = snapshot(editor)
    const outcome = await run(editor, [
      { op: 'table.splitCell', blockIndex: TABLE_BLOCK, row: 1, col: 1 },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.results[0].error).toContain('not merged')
    expect(snapshot(editor)).toBe(before)
  })
})

describe('table.setCell', () => {
  it('sets text, shading, borders, vAlign and width in one op', async () => {
    const editor = createEditor(fixture3x3())
    const outcome = await run(editor, [
      {
        op: 'table.setCell',
        blockIndex: TABLE_BLOCK,
        row: 1,
        col: 1,
        text: 'NEW',
        shading: 'fff2cc',
        vAlign: 'center',
        width: 120,
        borders: { bottom: { style: 'double', szEighths: 8, color: 'FF0000' } },
      },
    ])
    expect(outcome.ok).toBe(true)
    const target = cellAt(jsonTable(editor), 1, 1)
    expect(rowTexts(jsonTable(editor), 1)).toEqual(['D', 'NEW', 'F'])
    expect(target.attrs).toMatchObject({
      fill: 'FFF2CC',
      vAlign: 'center',
      colwidth: [120],
      borders: { bottom: { style: 'double', szEighths: 8, color: 'FF0000' } },
    })
  })

  it('clears text and shading with null and validates enum values', async () => {
    const editor = createEditor(fixture3x3())
    await run(editor, [
      { op: 'table.setCell', blockIndex: TABLE_BLOCK, row: 1, col: 0, text: 'X', shading: 'AABBCC' },
    ])
    const cleared = await run(editor, [
      { op: 'table.setCell', blockIndex: TABLE_BLOCK, row: 1, col: 0, text: null, shading: null },
    ])
    expect(cleared.ok).toBe(true)
    const target = cellAt(jsonTable(editor), 1, 0)
    expect(target.attrs?.fill).toBeNull()
    expect(rowTexts(jsonTable(editor), 1)).toEqual(['', 'E', 'F'])

    const bad = await run(editor, [
      { op: 'table.setCell', blockIndex: TABLE_BLOCK, row: 1, col: 1, vAlign: 'middle' },
    ])
    expect(bad.ok).toBe(false)
    expect(bad.results[0].error).toContain('vAlign')
  })
})

describe('table.setProperties', () => {
  it('sets alignment, indent, row height, colWidths, borders, repeat headers and a style preset', async () => {
    const editor = createEditor(fixture3x3())
    const outcome = await run(editor, [
      {
        op: 'table.setProperties',
        blockIndex: TABLE_BLOCK,
        alignment: 'center',
        indent: 567,
        rowHeight: { twips: 720, rule: 'exact', row: 1 },
        colWidths: [30, 30, 40],
        borders: { top: { style: 'single', szEighths: 8, color: '000000' } },
        repeatHeaderRows: true,
        stylePreset: 'grid',
      },
    ])
    expect(outcome.ok).toBe(true)
    const tableNode = jsonTable(editor)
    expect(tableNode.attrs).toMatchObject({
      tblAlign: 'center',
      indentTwips: 567,
      tblAutoFit: 'fixed',
      tblStyleId: null,
    })
    expect(tableNode.attrs?.colWidthsPct).toEqual([30, 30, 40])
    expect(tableNode.attrs?.borders).toMatchObject({
      top: { style: 'single', szEighths: 8, color: '000000' },
    })
    expect(tableNode.content![1].attrs).toMatchObject({ heightTwips: 720, heightRule: 'exact' })
    expect(tableNode.content![0].attrs?.repeatHeader).toBe(true)
    expect(tableNode.content![1].attrs?.repeatHeader).toBe(false)
    // grid preset: header fill applied through the table command
    expect(cellAt(tableNode, 0, 0).attrs?.fill).toBe('D9E2F3')
  })

  it('validates colWidths length and unknown border sides', async () => {
    const editor = createEditor(fixture3x3())
    const badWidths = await run(editor, [
      { op: 'table.setProperties', blockIndex: TABLE_BLOCK, colWidths: [50, 50] },
    ])
    expect(badWidths.ok).toBe(false)
    expect(badWidths.results[0].error).toContain('colWidths must be an array of 3')
    const badBorder = await run(editor, [
      { op: 'table.setProperties', blockIndex: TABLE_BLOCK, borders: { middle: {} } },
    ])
    expect(badBorder.ok).toBe(false)
    expect(badBorder.results[0].error).toContain('unknown border side')
  })

  it('repeatHeaderRows: false clears previously repeating header rows', async () => {
    const editor = createEditor(fixture3x3())
    await run(editor, [{ op: 'table.setProperties', blockIndex: TABLE_BLOCK, repeatHeaderRows: true }])
    expect(jsonTable(editor).content![0].attrs?.repeatHeader).toBe(true)
    await run(editor, [{ op: 'table.setProperties', blockIndex: TABLE_BLOCK, repeatHeaderRows: false }])
    expect(jsonTable(editor).content![0].attrs?.repeatHeader).toBe(false)
  })
})

describe('table.autoFit', () => {
  it('switches between window, contents and fixed modes', async () => {
    const editor = createEditor(fixture3x3())
    expect(
      (await run(editor, [{ op: 'table.autoFit', blockIndex: TABLE_BLOCK, mode: 'window' }])).ok,
    ).toBe(true)
    expect(jsonTable(editor).attrs).toMatchObject({ tblAutoFit: 'window', widthPct: 100 })

    expect(
      (await run(editor, [{ op: 'table.autoFit', blockIndex: TABLE_BLOCK, mode: 'contents' }])).ok,
    ).toBe(true)
    expect(jsonTable(editor).attrs).toMatchObject({ tblAutoFit: 'contents', widthPx: null })

    expect(
      (await run(editor, [{ op: 'table.autoFit', blockIndex: TABLE_BLOCK, mode: 'fixed' }])).ok,
    ).toBe(true)
    expect(jsonTable(editor).attrs?.tblAutoFit).toBe('fixed')
    expect(Number(jsonTable(editor).attrs?.widthPx)).toBeGreaterThan(0)
  })

  it('validates the mode', async () => {
    const editor = createEditor(fixture3x3())
    const outcome = await run(editor, [{ op: 'table.autoFit', blockIndex: TABLE_BLOCK, mode: 'wide' }])
    expect(outcome.ok).toBe(false)
    expect(outcome.results[0].error).toContain('mode must be')
  })
})

describe('per-op isolation and dryRun', () => {
  it('dryRun validates the whole batch without changing the document', async () => {
    const editor = createEditor(fixture3x3())
    const before = snapshot(editor)
    const outcome = await run(
      editor,
      [
        { op: 'table.insertRow', blockIndex: TABLE_BLOCK },
        { op: 'table.setCell', blockIndex: TABLE_BLOCK, row: 0, col: 0, text: 'changed' },
        { op: 'table.delete', blockIndex: TABLE_BLOCK },
      ],
      { dryRun: true },
    )
    expect(outcome.ok).toBe(true)
    expect(outcome.dryRun).toBe(true)
    expect(outcome.applied).toBe(0)
    expect(outcome.results.every((result) => result.ok && result.planned)).toBe(true)
    expect(snapshot(editor)).toBe(before)
  })

  it('a failing op does not abort later ops (per-op isolation)', async () => {
    const editor = createEditor(fixture3x3())
    const outcome = await run(editor, [
      { op: 'table.deleteRow', blockIndex: 99, row: 0 },
      { op: 'table.setCell', blockIndex: TABLE_BLOCK, row: 2, col: 2, text: 'Z' },
      { op: 'table.mergeCells', blockIndex: TABLE_BLOCK, row: 0, col: 0, rowspan: 9 },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.failed).toBe(2)
    expect(outcome.applied).toBe(1)
    expect(outcome.results.map((result) => result.ok)).toEqual([false, true, false])
    expect(rowTexts(jsonTable(editor), 2)).toEqual(['G', 'H', 'Z'])
  })

  it('caps a batch at MAX_OPERATIONS and says so', async () => {
    const editor = createEditor(fixture3x3())
    const ops = Array.from({ length: MAX_OPERATIONS + 1 }, () => ({
      op: 'table.autoFit',
      blockIndex: TABLE_BLOCK,
      mode: 'window',
    }))
    const outcome = await run(editor, ops, { dryRun: true })
    expect(outcome.results).toHaveLength(MAX_OPERATIONS)
    expect(outcome.output).toContain(`Only the first ${MAX_OPERATIONS} of ${MAX_OPERATIONS + 1}`)
  })
})

describe('contentWidthPxFromSection', () => {
  it('derives the text column width from page size, margins and columns', async () => {
    const a4 = contentWidthPxFromSection({
      pageWidth: 11906,
      pageHeight: 16838,
      marginLeft: 1440,
      marginRight: 1440,
      columns: 1,
    } as never)
    expect(a4).toBe(Math.round((11906 - 2880) / 15))
    expect(contentWidthPxFromSection(null)).toBe(600)
  })
})
describe('page ops', () => {
  it('page.break inserts a page-break paragraph after a block (or at the end)', async () => {
    const editor = createEditor(paragraphsDoc())
    const after = await run(editor, [{ op: 'page.break', afterBlockIndex: 0 }])
    expect(after.ok).toBe(true)
    let blocks = editor.getJSON().content!
    expect(blocks).toHaveLength(4)
    expect(blocks[1].type).toBe('docParagraph')
    expect(blocks[1].attrs?.pageBreakBefore).toBe(true)

    const atEnd = await run(editor, [{ op: 'page.break' }])
    expect(atEnd.ok).toBe(true)
    blocks = editor.getJSON().content!
    expect(blocks).toHaveLength(5)
    expect(blocks[4].attrs?.pageBreakBefore).toBe(true)
  })

  it('page.blank inserts a blank page and pushes the following block to the next page', async () => {
    const editor = createEditor(paragraphsDoc())
    const outcome = await run(editor, [{ op: 'page.blank', afterBlockIndex: 0 }])
    expect(outcome.ok).toBe(true)
    const blocks = editor.getJSON().content!
    expect(blocks).toHaveLength(4)
    expect(blocks[1].attrs?.pageBreakBefore).toBe(true)
    expect((blocks[1].content ?? []).length).toBe(0)
    // Word Blank Page semantics: the pushed block also starts a new page
    expect(blocks[2].attrs?.pageBreakBefore).toBe(true)
  })

  it('page.cover inserts a built-in preset at the start (and validates the preset)', async () => {
    const editor = createEditor(paragraphsDoc())
    const outcome = await run(editor, [{ op: 'page.cover', preset: 'classic' }])
    expect(outcome.ok).toBe(true)
    const blocks = editor.getJSON().content as unknown as JsonNode[]
    const texts = blocks.map((block) =>
      (block.content ?? []).map((child) => child.text ?? '').join(''),
    )
    expect(texts.some((text) => text.includes('文档标题'))).toBe(true)
    expect(texts.indexOf('first')).toBeGreaterThan(0)
    // the cover builder ends with a page-break paragraph
    expect(blocks.some((block) => block.attrs?.pageBreakBefore === true)).toBe(true)

    const bad = await run(editor, [{ op: 'page.cover', preset: 'no-such-preset' }])
    expect(bad.ok).toBe(false)
    expect(bad.results[0].error).toContain('preset must be one of')
  })

  it('page.cover afterBlockIndex inserts behind that block', async () => {
    const editor = createEditor(paragraphsDoc())
    const outcome = await run(editor, [{ op: 'page.cover', preset: 'classic', afterBlockIndex: 0 }])
    expect(outcome.ok).toBe(true)
    const blocks = editor.getJSON().content as unknown as JsonNode[]
    expect((blocks[0].content ?? []).map((child) => child.text).join('')).toBe('first')
    // the classic preset opens with an empty spacer paragraph, then the title
    expect(blocks.slice(1).map((block) => (block.content ?? []).map((child) => child.text ?? '').join(''))).toEqual(
      expect.arrayContaining([expect.stringContaining('文档标题')]),
    )
    expect((blocks[blocks.length - 1].content ?? []).map((child) => child.text).join('')).toBe('third')
  })

  it('page.numberFormat writes the page-number state for a non-final and the final section', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor, { sectionCount: 3, activeSection: 0 })

    const nonFinal = await run(
      editor,
      [{ op: 'page.numberFormat', sectionIndex: 1, fmt: 'romanLower', startAt: 3 }],
      { app: ctx },
    )
    expect(nonFinal.ok).toBe(true)
    expect(state.sections[1].pageNumberFmt).toBe('lowerRoman')
    expect(state.sections[1].pageNumberStart).toBe(3)
    expect(state.pgNumDirtySections).toEqual([1])
    expect(state.pgNumEdit).toBeNull()

    const final = await run(
      editor,
      [{ op: 'page.numberFormat', sectionIndex: 2, fmt: 'lettersUpper' }],
      { app: ctx },
    )
    expect(final.ok).toBe(true)
    expect(state.sections[2].pageNumberFmt).toBe('upperLetter')
    expect(state.pgNumEdit).toEqual({ fmt: 'upperLetter' })

    const decimal = await run(
      editor,
      [{ op: 'page.numberFormat', sectionIndex: 2, fmt: 'decimal' }],
      { app: ctx },
    )
    expect(decimal.ok).toBe(true)
    expect(state.sections[2].pageNumberFmt).toBeUndefined()
    expect(state.pgNumEdit).toEqual({})

    const badFmt = await run(editor, [{ op: 'page.numberFormat', fmt: 'greek' }], { app: ctx })
    expect(badFmt.ok).toBe(false)
    expect(badFmt.results[0].error).toContain('fmt must be one of')
  })
})

describe('section ops', () => {
  it('section.break inserts the break after a block through the real helper', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    const outcome = await run(
      editor,
      [{ op: 'section.break', type: 'continuous', afterBlockIndex: 1 }],
      { app: ctx },
    )
    expect(outcome.ok).toBe(true)
    expect(state.sectionBreaks).toEqual([{ type: 'continuous', afterBlockIndex: 1 }])
    expect(editor.getJSON().content!.some((block) => block.type === 'docProtected')).toBe(true)

    const badType = await run(editor, [{ op: 'section.break', type: 'page' }], { app: ctx })
    expect(badType.ok).toBe(false)
    expect(badType.results[0].error).toContain('type must be one of')
    const badBlock = await run(
      editor,
      [{ op: 'section.break', type: 'nextPage', afterBlockIndex: 99 }],
      { app: ctx },
    )
    expect(badBlock.ok).toBe(false)
    expect(badBlock.results[0].error).toContain('afterBlockIndex must be an integer')
  })

  it('section.set runs applyPageSetup for the target section (validate then apply)', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor, { sectionCount: 2, activeSection: 0 })
    const outcome = await run(
      editor,
      [
        {
          op: 'section.set',
          sectionIndex: 1,
          margins: { top: 720 },
          pageColor: 'EFEFEF',
        },
      ],
      { app: ctx },
    )
    expect(outcome.ok).toBe(true)
    expect(state.pageSetupCalls).toHaveLength(2)
    expect(state.pageSetupCalls[0]).toMatchObject({ sectionIndex: 1, dryRun: true })
    expect(state.pageSetupCalls[1]).toMatchObject({ sectionIndex: 1, dryRun: false })
    // final section: SaveOptions.section + dirty flag drive the save
    expect(state.section?.marginTop).toBe(720)
    expect(state.sectionDirty).toBe(true)
    expect(state.pageColor).toBe('EFEFEF')
    expect(state.pageColorDirty).toBe(true)
  })

  it('section.set applies a non-final section through sectionsDirty', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor, { sectionCount: 2, activeSection: 0 })
    const outcome = await run(
      editor,
      [{ op: 'section.set', sectionIndex: 0, margins: { top: 567 } }],
      { app: ctx },
    )
    expect(outcome.ok).toBe(true)
    expect(state.sections[0].settings.marginTop).toBe(567)
    expect(state.sectionsDirty).toEqual([0])
    expect(state.sectionDirty).toBe(false)
  })

  it('section.set validates unknown fields, empty payloads and pipeline errors', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx } = makeFakeApp(editor)
    const unknown = await run(editor, [{ op: 'section.set', sectionIndex: 0, header: 'x' }], { app: ctx })
    expect(unknown.ok).toBe(false)
    expect(unknown.results[0].error).toContain('unknown section.set field')
    const empty = await run(editor, [{ op: 'section.set' }], { app: ctx })
    expect(empty.ok).toBe(false)
    expect(empty.results[0].error).toContain('nothing to change')
    const badRange = await run(editor, [{ op: 'section.set', sectionIndex: 9, columns: 2 }], { app: ctx })
    expect(badRange.ok).toBe(false)
    expect(badRange.results[0].error).toContain('sectionIndex')

    const failing = makeFakeApp(editor, {
      applyPageSetup: () => ({ ok: false as const, error: 'margins.top must be twips' }),
    })
    const rejected = await run(
      editor,
      [{ op: 'section.set', margins: { top: -1 } }],
      { app: failing.ctx },
    )
    expect(rejected.ok).toBe(false)
    expect(rejected.results[0].error).toContain('margins.top must be twips')
  })
})

describe('note ops', () => {
  it('notes.insert anchors the marker at the end of a block and updates the note list', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    const outcome = await run(
      editor,
      [{ op: 'notes.insert', kind: 'footnote', text: 'See appendix', blockIndex: 1 }],
      { app: ctx },
    )
    expect(outcome.ok).toBe(true)
    expect(state.footnotes).toEqual([{ id: '1', text: 'See appendix' }])
    expect(state.notesDirty).toBe(true)
    expect(JSON.stringify(editor.getJSON().content![1])).toContain('docNoteRef')
    expect(outcome.results[0].record).toMatchObject({ kind: 'footnote', id: '1' })
  })

  it('notes.insert assigns distinct ids for two inserts in one batch', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    const outcome = await run(
      editor,
      [
        { op: 'notes.insert', kind: 'endnote', text: 'First endnote', blockIndex: 0 },
        { op: 'notes.insert', kind: 'endnote', text: 'Second endnote', blockIndex: 2 },
      ],
      { app: ctx },
    )
    expect(outcome.ok).toBe(true)
    expect(state.endnotes.map((note) => note.text)).toEqual(['First endnote', 'Second endnote'])
    expect(new Set(state.endnotes.map((note) => note.id)).size).toBe(2)
  })

  it('notes.update replaces the note text and notes.delete removes list entry + marker', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    await run(editor, [{ op: 'notes.insert', kind: 'footnote', text: 'old', blockIndex: 0 }], {
      app: ctx,
    })
    const updated = await run(
      editor,
      [{ op: 'notes.update', kind: 'footnote', id: '1', text: 'new text' }],
      { app: ctx },
    )
    expect(updated.ok).toBe(true)
    expect(state.footnotes).toEqual([{ id: '1', text: 'new text' }])

    const deleted = await run(editor, [{ op: 'notes.delete', kind: 'footnote', id: '1' }], { app: ctx })
    expect(deleted.ok).toBe(true)
    expect(state.footnotes).toEqual([])
    expect(JSON.stringify(editor.getJSON())).not.toContain('docNoteRef')

    const missing = await run(editor, [{ op: 'notes.delete', kind: 'footnote', id: '42' }], { app: ctx })
    expect(missing.ok).toBe(false)
    expect(missing.results[0].error).toContain('does not exist')
  })

  it('notes ops validate kind/text and stay inert under dryRun', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    const before = snapshot(editor)
    const dry = await run(
      editor,
      [
        { op: 'notes.insert', kind: 'footnote', text: 'planned', blockIndex: 0 },
        { op: 'notes.delete', kind: 'footnote', id: '1' },
      ],
      { app: ctx, dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(dry.results[1].ok).toBe(false) // no footnote 1 exists yet
    expect(snapshot(editor)).toBe(before)
    expect(state.footnotes).toEqual([])

    const badKind = await run(
      editor,
      [{ op: 'notes.insert', kind: 'margin-note', text: 'x' }],
      { app: ctx },
    )
    expect(badKind.ok).toBe(false)
    expect(badKind.results[0].error).toContain('kind must be')
  })
})

describe('field ops', () => {
  it('field.insert appends the inline field at the end of a block', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    const outcome = await run(editor, [{ op: 'field.insert', kind: 'PAGE', blockIndex: 1 }], {
      app: ctx,
    })
    expect(outcome.ok).toBe(true)
    expect(state.fieldInsertions).toEqual(['PAGE'])
    const block = editor.getJSON().content![1] as unknown as JsonNode
    const marks = block.content!.flatMap((child) => child.marks ?? [])
    expect(marks.some((mark) => mark.type === 'instrField')).toBe(true)
    expect(JSON.stringify(block)).toContain('"instr":"PAGE"')
  })

  it('field.insert afterBlockIndex puts the field in its own paragraph', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx } = makeFakeApp(editor)
    const outcome = await run(editor, [{ op: 'field.insert', kind: 'FILENAME', afterBlockIndex: 0 }], {
      app: ctx,
    })
    expect(outcome.ok).toBe(true)
    const blocks = editor.getJSON().content!
    expect(blocks).toHaveLength(4)
    expect(blocks[1].type).toBe('docParagraph')
    expect(JSON.stringify(blocks[1])).toContain('"instr":"FILENAME"')
  })

  it('field.insert validates kind/addressing and never mutates under dryRun', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    const before = snapshot(editor)
    const dry = await run(editor, [{ op: 'field.insert', kind: 'DATE', blockIndex: 0 }], {
      app: ctx,
      dryRun: true,
    })
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(snapshot(editor)).toBe(before)
    expect(state.fieldInsertions).toEqual([])

    const badKind = await run(editor, [{ op: 'field.insert', kind: 'AUTHOR' }], { app: ctx })
    expect(badKind.ok).toBe(false)
    expect(badKind.results[0].error).toContain('kind must be one of')
    const both = await run(
      editor,
      [{ op: 'field.insert', kind: 'PAGE', blockIndex: 0, afterBlockIndex: 1 }],
      { app: ctx },
    )
    expect(both.ok).toBe(false)
    expect(both.results[0].error).toContain('either blockIndex or afterBlockIndex')
  })

  it('field.updateAll runs the real F9 helper and reports the field count', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    await run(editor, [{ op: 'field.insert', kind: 'PAGE', blockIndex: 0 }], { app: ctx })
    const outcome = await run(editor, [{ op: 'field.updateAll' }], { app: ctx })
    expect(outcome.ok).toBe(true)
    expect(state.fieldUpdates).toBe(1)
    expect(outcome.results[0].record).toMatchObject({ fields: 1 })

    const dry = await run(editor, [{ op: 'field.updateAll' }], { app: ctx, dryRun: true })
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(state.fieldUpdates).toBe(1)
  })
})

describe('B2 registry + isolation', () => {
  it('registers all B2 ops in the guide with their groups', async () => {
    const defs = describeOperations()
    const names = defs.map((def) => def.name)
    for (const name of [
      'page.break',
      'page.blank',
      'page.cover',
      'page.numberFormat',
      'section.break',
      'section.set',
      'notes.insert',
      'notes.update',
      'notes.delete',
      'field.insert',
      'field.updateAll',
    ]) {
      expect(names).toContain(name)
      expect(OPERATIONS_GUIDE).toContain(name)
    }
    expect(new Set(defs.map((def) => def.group))).toEqual(
      new Set(['table', 'page', 'section', 'notes', 'field', 'format', 'object', 'reference', 'review']),
    )
  })

  it('keeps per-op isolation across B2 ops (a failure does not abort later ops)', async () => {
    const editor = createEditor(paragraphsDoc())
    const { ctx, state } = makeFakeApp(editor)
    const outcome = await run(
      editor,
      [
        { op: 'field.insert', kind: 'NOPE' },
        { op: 'page.break', afterBlockIndex: 0 },
        { op: 'notes.insert', kind: 'footnote', text: 'ok', blockIndex: 2 },
      ],
      { app: ctx },
    )
    expect(outcome.ok).toBe(false)
    expect(outcome.failed).toBe(1)
    expect(outcome.applied).toBe(2)
    expect(outcome.results.map((result) => result.ok)).toEqual([false, true, true])
    expect(state.footnotes).toHaveLength(1)
  })
})

// ── B3 inline/object fixtures ────────────────────────────────────────────────

const listItemJson = (text: string, numId: string, ilvl = 0): JsonNode => ({
  type: 'docListItem',
  attrs: { docxIndex: null, kind: 'ordered', numId, ilvl },
  content: text ? [{ type: 'text', text }] : [],
})

const headingJson = (text: string, level = 1): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [{ type: 'text', text }],
})

const formattedParagraph = (): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, align: 'center', indentLeft: 1440, spaceBefore: 240 },
  content: [{ type: 'text', text: 'Hello World', marks: [{ type: 'bold' }] }],
})

const imageBlock = (attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docProtected',
  attrs: {
    docxIndex: null,
    blockType: 'image',
    label: 'Image',
    previewText: 'chart',
    imageDataUrl: 'data:image/png;base64,AAAA',
    imageWidthPx: 100,
    imageHeightPx: 50,
    ...attrs,
  },
})

const shapeBlock = (attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docProtected',
  attrs: {
    docxIndex: null,
    blockType: 'passthrough',
    label: 'Text Box',
    previewText: '',
    genXml: '<w:p><w:r><w:drawing/></w:r></w:p>',
    textboxes: [
      {
        fill: 'FFFFFF',
        borderColor: '000000',
        widthPx: 189,
        heightPx: 113,
        paras: [{ runs: [{ text: '' }] }],
      },
    ],
    ...attrs,
  },
})

function blockText(block: JsonNode | undefined): string {
  return (block?.content ?? []).map((child) => child.text ?? '').join('')
}

function nodeTexts(editor: Editor): string[] {
  return (editor.getJSON().content as unknown as JsonNode[]).map(blockText)
}

describe('format ops', () => {
  it('format.changeCase changes only the matched blocks', async () => {
    const editor = createEditor([para('hello world'), para('second')])
    const outcome = await run(editor, [
      { op: 'format.changeCase', target: { blockIndexes: [0] }, mode: 'upper' },
    ])
    expect(outcome.ok).toBe(true)
    expect(nodeTexts(editor)).toEqual(['HELLO WORLD', 'second'])

    await run(editor, [{ op: 'format.changeCase', target: { blockIndexes: [0] }, mode: 'title' }])
    expect(nodeTexts(editor)[0]).toBe('Hello World')

    const bad = await run(editor, [
      { op: 'format.changeCase', target: { blockIndexes: [0] }, mode: 'camel' },
    ])
    expect(bad.ok).toBe(false)
    expect(bad.results[0].error).toContain('mode must be one of')
  })

  it('format.changeCase dryRun does not mutate and validates the target', async () => {
    const editor = createEditor([para('hello')])
    const before = snapshot(editor)
    const dry = await run(
      editor,
      [{ op: 'format.changeCase', target: { blockIndexes: [0] }, mode: 'upper' }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(snapshot(editor)).toBe(before)
    const noMatch = await run(editor, [
      { op: 'format.changeCase', target: { containsText: 'absent' }, mode: 'upper' },
    ])
    expect(noMatch.ok).toBe(false)
    expect(noMatch.results[0].error).toContain('matched no blocks')
  })

  it('format.clear drops direct character and paragraph formatting', async () => {
    const editor = createEditor([formattedParagraph()])
    const outcome = await run(editor, [{ op: 'format.clear', target: { blockIndexes: [0] } }])
    expect(outcome.ok).toBe(true)
    const block = editor.getJSON().content![0] as unknown as JsonNode
    const marks = (block.content ?? []).flatMap((child) => child.marks ?? [])
    expect(marks).toHaveLength(0)
    expect(block.attrs).toMatchObject({
      align: null,
      indentLeft: null,
      spaceBefore: null,
    })
  })

  it('format.setBidi flips paragraph direction both ways', async () => {
    const editor = createEditor([para('مرحبا'), para('hello')])
    await run(editor, [{ op: 'format.setBidi', target: { blockIndexes: [0] }, direction: 'rtl' }])
    expect(editor.getJSON().content![0].attrs).toMatchObject({ bidi: true, bidiInferred: false })
    await run(editor, [{ op: 'format.setBidi', target: { blockIndexes: [0] }, direction: 'ltr' }])
    expect(editor.getJSON().content![0].attrs).toMatchObject({ bidi: false })
    const bad = await run(editor, [
      { op: 'format.setBidi', target: { blockIndexes: [0] }, direction: 'up' },
    ])
    expect(bad.ok).toBe(false)
  })

  it('format.hangingIndent steps the hanging indent and steps back', async () => {
    const editor = createEditor([para('first'), para('second')])
    await run(editor, [{ op: 'format.hangingIndent', target: { blockIndexes: [0] }, steps: 1 }])
    expect(editor.getJSON().content![0].attrs).toMatchObject({
      indentLeft: 720,
      indentFirstLine: -720,
    })
    await run(editor, [{ op: 'format.hangingIndent', target: { blockIndexes: [0] }, steps: -1 }])
    expect(editor.getJSON().content![0].attrs).toMatchObject({
      indentLeft: null,
      indentFirstLine: null,
    })
    const bad = await run(editor, [
      { op: 'format.hangingIndent', target: { blockIndexes: [0] }, steps: 0 },
    ])
    expect(bad.ok).toBe(false)
    expect(bad.results[0].error).toContain('steps must be a non-zero integer')
  })

  it('format.applyStyle applies gallery paragraph styles', async () => {
    const editor = createEditor([formattedParagraph(), para('second')])
    await run(editor, [{ op: 'format.applyStyle', target: { blockIndexes: [0] }, style: 'h2' }])
    const heading = editor.getJSON().content![0] as unknown as JsonNode
    expect(heading.type).toBe('docHeading')
    expect(heading.attrs?.level).toBe(2)
    await run(editor, [{ op: 'format.applyStyle', target: { blockIndexes: [0] }, style: 'p' }])
    expect(editor.getJSON().content![0].type).toBe('docParagraph')
    const bad = await run(editor, [
      { op: 'format.applyStyle', target: { blockIndexes: [0] }, style: 'h9' },
    ])
    expect(bad.ok).toBe(false)
  })

  it('format.listRestart creates a restart definition and rewrites the following items', async () => {
    const editor = createEditor([listItemJson('one', '5'), listItemJson('two', '5')])
    const { ctx, state } = makeFakeApp(editor)
    const outcome = await run(editor, [{ op: 'format.listRestart', target: { blockIndexes: [0] } }], {
      app: ctx,
    })
    expect(outcome.ok).toBe(true)
    expect(editor.getJSON().content!.map((block) => block.attrs?.numId)).toEqual(['6', '6'])
    expect(state.pendingNumbering.restartNums).toHaveLength(1)
    expect(state.pendingNumbering.restartNums[0].numId).toBe('6')

    const notList = await run(editor, [{ op: 'format.listRestart', target: { blockIndexes: [0] } }], {
      app: ctx,
    })
    expect(notList.ok).toBe(true) // the rewritten items are still list items

    const dry = await run(
      editor,
      [{ op: 'format.listRestart', target: { blockIndexes: [1] } }],
      { app: ctx, dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
  })

  it('format.listContinue merges the target list into the previous numbering', async () => {
    const editor = createEditor([listItemJson('one', '5'), listItemJson('two', '7')])
    const { ctx } = makeFakeApp(editor)
    const outcome = await run(editor, [{ op: 'format.listContinue', target: { blockIndexes: [1] } }], {
      app: ctx,
    })
    expect(outcome.ok).toBe(true)
    expect(editor.getJSON().content!.map((block) => block.attrs?.numId)).toEqual(['5', '5'])

    const single = createEditor([listItemJson('only', '5')])
    const singleApp = makeFakeApp(single)
    const noPrev = await run(
      single,
      [{ op: 'format.listContinue', target: { blockIndexes: [0] } }],
      { app: singleApp.ctx },
    )
    expect(noPrev.ok).toBe(false)
    expect(noPrev.results[0].error).toContain('continue produced no change')
  })

  it('format.multilevelList creates the gallery definition and converts paragraphs', async () => {
    const editor = createEditor([para('one'), para('two'), headingJson('title')])
    const { ctx, state } = makeFakeApp(editor)
    const outcome = await run(
      editor,
      [
        {
          op: 'format.multilevelList',
          target: { blockIndexes: [0, 1, 2] },
          preset: 'decimal',
        },
      ],
      { app: ctx },
    )
    expect(outcome.ok).toBe(true)
    const blocks = editor.getJSON().content! as unknown as JsonNode[]
    expect(blocks[0].type).toBe('docListItem')
    expect(blocks[1].type).toBe('docListItem')
    expect(blocks[0].attrs).toMatchObject({ kind: 'ordered', numId: '6', ilvl: 0 })
    // headings are not converted (they keep their level)
    expect(blocks[2].type).toBe('docHeading')
    expect(outcome.results[0].record).toMatchObject({ blocks: 2, skipped: 1 })
    expect(state.pendingNumbering.newDefs).toHaveLength(1)

    const bulletApp = makeFakeApp(createEditor([para('x')]))
    const bulletEditor = bulletApp.ctx.review().editor!
    const bullet = await run(
      bulletEditor,
      [{ op: 'format.multilevelList', target: { blockIndexes: [0] }, preset: 'bullet' }],
      { app: bulletApp.ctx },
    )
    expect(bullet.ok).toBe(true)
    expect(bulletEditor.getJSON().content![0].attrs).toMatchObject({ kind: 'bullet' })

    const badPreset = await run(
      editor,
      [{ op: 'format.multilevelList', target: { blockIndexes: [0] }, preset: 'nope' }],
      { app: ctx },
    )
    expect(badPreset.ok).toBe(false)
    expect(badPreset.results[0].error).toContain('preset must be one of')
  })
})

describe('object ops', () => {
  it('object.textbox.insert adds a text box after a block and applies its text', async () => {
    const editor = createEditor([para('first'), para('second')])
    const outcome = await run(
      editor,
      [{ op: 'object.textbox.insert', afterBlockIndex: 0, text: 'Box note' }],
    )
    expect(outcome.ok).toBe(true)
    const blocks = editor.getJSON().content! as unknown as JsonNode[]
    expect(blocks).toHaveLength(3)
    const box = blocks[1]
    expect(box.type).toBe('docProtected')
    const textboxes = box.attrs?.textboxes as Array<{
      paras: Array<{ runs: Array<{ text: string }> }>
    }>
    expect(textboxes[0].paras[0].runs[0].text).toBe('Box note')
    expect(typeof box.attrs?.genXml).toBe('string')

    const dry = await run(
      editor,
      [{ op: 'object.textbox.insert', afterBlockIndex: 0, text: 'planned' }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(editor.getJSON().content).toHaveLength(3)
  })

  it('object.wordArt.insert uses the gallery preset and optional text', async () => {
    const editor = createEditor([para('first')])
    const outcome = await run(editor, [
      { op: 'object.wordArt.insert', afterBlockIndex: 0, preset: 'blue', text: 'Hello' },
    ])
    expect(outcome.ok).toBe(true)
    const block = editor.getJSON().content![1] as unknown as JsonNode
    expect(block.type).toBe('docProtected')
    const textboxes = block.attrs?.textboxes as Array<{
      wordArtId?: string
      paras: Array<{ runs: Array<{ text: string }> }>
    }>
    expect(textboxes[0].wordArtId).toBe('blue')
    expect(textboxes[0].paras[0].runs[0].text).toBe('Hello')

    const bad = await run(editor, [{ op: 'object.wordArt.insert', preset: 'neon' }])
    expect(bad.ok).toBe(false)
    expect(bad.results[0].error).toContain('preset must be one of')
  })

  it('object.dropCap sets and clears the paragraph drop cap', async () => {
    const editor = createEditor([para('Drop cap paragraph'), para('second')])
    await run(editor, [{ op: 'object.dropCap', target: { blockIndexes: [0] }, mode: 'dropped' }])
    expect(JSON.parse(String(editor.getJSON().content![0].attrs?.dropCap))).toEqual({
      type: 'drop',
      lines: 3,
    })
    await run(editor, [{ op: 'object.dropCap', target: { blockIndexes: [0] }, mode: 'margin' }])
    expect(JSON.parse(String(editor.getJSON().content![0].attrs?.dropCap))).toEqual({
      type: 'margin',
      lines: 3,
    })
    await run(editor, [{ op: 'object.dropCap', target: { blockIndexes: [0] }, mode: 'none' }])
    expect(editor.getJSON().content![0].attrs?.dropCap).toBeNull()
    const bad = await run(editor, [
      { op: 'object.dropCap', target: { blockIndexes: [0] }, mode: 'huge' },
    ])
    expect(bad.ok).toBe(false)
  })

  it('object.equation.insert creates display and inline equations and rejects bad LaTeX', async () => {
    const editor = createEditor([para('first')])
    const display = await run(editor, [{ op: 'object.equation.insert', latex: 'a^2 + b^2 = c^2' }])
    expect(display.ok).toBe(true)
    const blocks = editor.getJSON().content! as unknown as JsonNode[]
    // the insert helper splits the caret paragraph: the equation is its own block
    const equations = blocks.filter((block) => block.attrs?.formulaDisplay)
    expect(equations).toHaveLength(1)
    expect(equations[0].type).toBe('docProtected')
    expect(equations[0].attrs?.blockType).toBe('passthrough')

    const inline = await run(editor, [
      { op: 'object.equation.insert', latex: 'x = 1', inline: true, afterBlockIndex: 0 },
    ])
    expect(inline.ok).toBe(true)
    expect(JSON.stringify(editor.getJSON().content![0])).toContain('docInlineMath')

    const bad = await run(editor, [{ op: 'object.equation.insert', latex: '\\frac{1' }])
    expect(bad.ok).toBe(false)
    expect(bad.results[0].error).toContain('invalid LaTeX')
  })

  it('object.symbol.insert inserts inline at a block end or in its own paragraph', async () => {
    const editor = createEditor([para('first'), para('second')])
    const inline = await run(editor, [
      { op: 'object.symbol.insert', char: '→', afterBlockIndex: 0 },
    ])
    expect(inline.ok).toBe(true)
    expect(nodeTexts(editor)[0]).toBe('first→')

    const own = await run(editor, [
      { op: 'object.symbol.insert', char: '©', afterBlockIndex: 0, ownParagraph: true },
    ])
    expect(own.ok).toBe(true)
    const blocks = editor.getJSON().content! as unknown as JsonNode[]
    expect(blockText(blocks[1])).toBe('©')

    const bad = await run(editor, [{ op: 'object.symbol.insert', char: 'a\nb' }])
    expect(bad.ok).toBe(false)
    expect(bad.results[0].error).toContain('control characters')
  })

  it('object.image.format updates image size, alignment, wrap, rotation and flip attrs', async () => {
    const editor = createEditor([para('first'), imageBlock()])
    const size = await run(editor, [
      { op: 'object.image.format', blockIndex: 1, widthPx: 200, align: 'center' },
    ])
    expect(size.ok).toBe(true)
    let image = editor.getJSON().content![1] as unknown as JsonNode
    expect(image.attrs).toMatchObject({
      imageWidthPx: 200,
      imageHeightPx: 100,
      imageAlign: 'center',
    })

    const wrap = await run(editor, [
      {
        op: 'object.image.format',
        blockIndex: 1,
        wrap: 'square-left',
        rotationDeg: 45,
        flipH: true,
        flipV: false,
      },
    ])
    expect(wrap.ok).toBe(true)
    image = editor.getJSON().content![1] as unknown as JsonNode
    expect(image.attrs).toMatchObject({
      imageWrap: 'square-left',
      imageRotDeg: 45,
      imageFlipH: true,
      imageFlipV: false,
    })

    const z = await run(editor, [{ op: 'object.image.format', blockIndex: 1, zOrder: 5 }])
    expect(z.ok).toBe(true)
    image = editor.getJSON().content![1] as unknown as JsonNode
    expect(image.attrs).toMatchObject({ imageZOrder: 5, imageWrap: 'front' })

    const invalid = await run(editor, [
      { op: 'object.image.format', blockIndex: 1, wrap: 'diagonal' },
    ])
    expect(invalid.ok).toBe(false)
    const notImage = await run(editor, [{ op: 'object.image.format', blockIndex: 0, widthPx: 100 }])
    expect(notImage.ok).toBe(false)
    expect(notImage.results[0].error).toContain('not an image block')

    const before = snapshot(editor)
    const dry = await run(
      editor,
      [{ op: 'object.image.format', blockIndex: 1, widthPx: 150 }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(snapshot(editor)).toBe(before)
  })

  it('object.shape.format updates fill, outline, text, size and stacking order', async () => {
    const editor = createEditor([para('first'), shapeBlock()])
    const outcome = await run(editor, [
      {
        op: 'object.shape.format',
        blockIndex: 1,
        fillHex: 'FF0000',
        lineHex: '00FF00',
        text: 'Label',
        widthEmu: 914400,
        heightEmu: 457200,
      },
    ])
    expect(outcome.ok).toBe(true)
    const shape = editor.getJSON().content![1] as unknown as JsonNode
    const box = (shape.attrs?.textboxes as Array<Record<string, unknown>>)[0] as {
      fill: string
      borderColor: string
      widthPx: number
      heightPx: number
      paras: Array<{ runs: Array<{ text: string }> }>
    }
    expect(box.fill).toBe('FF0000')
    expect(box.borderColor).toBe('00FF00')
    expect(box.widthPx).toBe(96)
    expect(box.heightPx).toBe(48)
    expect(box.paras[0].runs[0].text).toBe('Label')

    const z = await run(editor, [{ op: 'object.shape.format', blockIndex: 1, zOrder: 2 }])
    expect(z.ok).toBe(true)
    expect((editor.getJSON().content![1] as unknown as JsonNode).attrs).toMatchObject({
      imageZOrder: 2,
      imageWrap: 'front',
    })

    const wrongBlock = await run(editor, [{ op: 'object.shape.format', blockIndex: 0, fillHex: 'FF0000' }])
    expect(wrongBlock.ok).toBe(false)
    expect(wrongBlock.results[0].error).toContain('not a textbox/shape block')

    // imported shapes keep their original anchor bytes: a reorder must not fake success
    const imported = createEditor([para('first'), shapeBlock({ docxIndex: 3 })])
    const importedZ = await run(imported, [{ op: 'object.shape.format', blockIndex: 1, zOrder: 2 }])
    expect(importedZ.ok).toBe(false)
    expect(importedZ.results[0].error).toContain('zOrder on imported')
  })

  it('object.image.unsupported documents the host-bytes limitation as a clear error', async () => {
    const editor = createEditor([para('first'), imageBlock()])
    const before = snapshot(editor)
    for (const feature of ['replace', 'crop', 'removeBackground']) {
      const outcome = await run(editor, [{ op: 'object.image.unsupported', feature }])
      expect(outcome.ok).toBe(false)
      expect(outcome.results[0].error).toContain('unsupported')
      expect(outcome.results[0].error).toContain('host')
    }
    expect(snapshot(editor)).toBe(before)
  })
})

describe('B3 registry + isolation', () => {
  it('registers every B3 op in the guide with its group', async () => {
    const defs = describeOperations()
    const names = defs.map((def) => def.name)
    for (const name of [
      'format.changeCase',
      'format.clear',
      'format.setBidi',
      'format.hangingIndent',
      'format.applyStyle',
      'format.listRestart',
      'format.listContinue',
      'format.multilevelList',
      'object.textbox.insert',
      'object.wordArt.insert',
      'object.dropCap',
      'object.equation.insert',
      'object.symbol.insert',
      'object.image.format',
      'object.shape.format',
      'object.image.unsupported',
    ]) {
      expect(names).toContain(name)
      expect(OPERATIONS_GUIDE).toContain(name)
    }
    const groups = new Set(defs.map((def) => def.group))
    expect(groups).toEqual(
      new Set(['table', 'page', 'section', 'notes', 'field', 'format', 'object', 'reference', 'review']),
    )
  })

  it('keeps per-op isolation across B3 ops', async () => {
    const editor = createEditor([para('one'), para('two')])
    const outcome = await run(editor, [
      { op: 'format.changeCase', target: { blockIndexes: [0] }, mode: 'nope' },
      { op: 'format.changeCase', target: { blockIndexes: [1] }, mode: 'upper' },
      { op: 'object.symbol.insert', char: '!', afterBlockIndex: 1 },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.failed).toBe(1)
    expect(outcome.applied).toBe(2)
    expect(nodeTexts(editor)).toEqual(['one', 'TWO!'])
  })
})

// ── B4 references / review fixtures ──────────────────────────────────────────

const revisionParagraph = (text: string, kind: 'ins' | 'del' = 'ins'): JsonNode => ({
  type: 'docParagraph',
  attrs: {
    docxIndex: null,
    blockRevision: { kind, author: 'Agent', date: '2026-01-01T00:00:00Z' },
  },
  content: [{ type: 'text', text }],
})

function pendingRevisions(editor: Editor): number {
  return collectRevisions(editor.state.doc).length
}

describe('reference ops', () => {
  it('reference.bookmark.insert / remove manage paragraph bookmarks', async () => {
    const editor = createEditor([para('intro'), para('second')])
    const inserted = await run(editor, [
      { op: 'reference.bookmark.insert', blockIndex: 0, name: 'Intro' },
    ])
    expect(inserted.ok).toBe(true)
    expect(editor.getJSON().content![0].attrs?.bookmarks).toEqual(['Intro'])

    const duplicate = await run(editor, [
      { op: 'reference.bookmark.insert', blockIndex: 1, name: 'Intro' },
    ])
    expect(duplicate.ok).toBe(false)
    expect(duplicate.results[0].error).toContain('already exists')

    const badName = await run(editor, [
      { op: 'reference.bookmark.insert', blockIndex: 1, name: 'has space' },
    ])
    expect(badName.ok).toBe(false)
    expect(badName.results[0].error).toContain('bookmark name')

    const removed = await run(editor, [{ op: 'reference.bookmark.remove', name: 'Intro' }])
    expect(removed.ok).toBe(true)
    expect(editor.getJSON().content![0].attrs?.bookmarks).toBeNull()

    const missing = await run(editor, [{ op: 'reference.bookmark.remove', name: 'Nope' }])
    expect(missing.ok).toBe(false)
    expect(missing.results[0].error).toContain('does not exist')

    const before = snapshot(editor)
    const dry = await run(
      editor,
      [{ op: 'reference.bookmark.insert', blockIndex: 0, name: 'Dry' }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(snapshot(editor)).toBe(before)
  })

  it('reference.crossRef.insert inserts a REF field to an existing bookmark', async () => {
    const editor = createEditor([para('Chapter One'), para('body')])
    await run(editor, [{ op: 'reference.bookmark.insert', blockIndex: 0, name: 'Chapter1' }])
    const outcome = await run(editor, [
      { op: 'reference.crossRef.insert', bookmark: 'Chapter1', afterBlockIndex: 0 },
    ])
    expect(outcome.ok).toBe(true)
    expect(JSON.stringify(editor.getJSON().content![0])).toContain('refField')
    expect(JSON.stringify(editor.getJSON().content![0])).toContain('Chapter One')

    const missing = await run(editor, [{ op: 'reference.crossRef.insert', bookmark: 'Nope' }])
    expect(missing.ok).toBe(false)
    expect(missing.results[0].error).toContain('does not exist')
  })

  it('reference.caption.insert numbers captions sequentially per label', async () => {
    const editor = createEditor([para('figure area')])
    const { ctx } = makeFakeApp(editor)
    const first = await run(
      editor,
      [{ op: 'reference.caption.insert', blockIndex: 0, label: 'Figure', text: 'Diagram' }],
      { app: ctx },
    )
    expect(first.ok).toBe(true)
    expect(first.results[0].record).toMatchObject({ label: 'Figure', number: 1 })
    let blocks = editor.getJSON().content! as unknown as JsonNode[]
    expect(blocks[1].type).toBe('docProtected')
    expect(String(blocks[1].attrs?.genXml)).toContain('SEQ Figure')
    expect(blocks[1].attrs?.fieldDisplay).toMatchObject({ left: 'Figure 1 Diagram' })

    const second = await run(
      editor,
      [{ op: 'reference.caption.insert', blockIndex: 0, label: 'Figure' }],
      { app: ctx },
    )
    expect(second.results[0].record).toMatchObject({ number: 2 })
    blocks = editor.getJSON().content! as unknown as JsonNode[]
    const secondCaption = blocks.find(
      (block) => (block.attrs?.fieldDisplay as { left?: string } | undefined)?.left === 'Figure 2',
    )
    expect(secondCaption).toBeTruthy()

    const noApp = await run(editor, [{ op: 'reference.caption.insert', label: 'Figure' }])
    expect(noApp.ok).toBe(false)
    expect(noApp.results[0].error).toContain('app context')
  })

  it('reference.source.add assigns unique tags and citation/bibliography use the sources', async () => {
    const editor = createEditor([para('body')])
    const { ctx, state } = makeFakeApp(editor)
    const added = await run(
      editor,
      [
        {
          op: 'reference.source.add',
          title: 'Deep Work',
          author: 'Wang',
          year: '2024',
          publisher: 'Press',
        },
      ],
      { app: ctx },
    )
    expect(added.ok).toBe(true)
    expect(state.sources).toHaveLength(1)
    expect(state.sources[0].tag).toBe('Wang2024')

    const duplicate = await run(
      editor,
      [{ op: 'reference.source.add', title: 'Deep Work', author: 'Wang', year: '2024' }],
      { app: ctx },
    )
    expect(duplicate.results[0].record).toMatchObject({ tag: 'Wang2024_2' })

    const noTitle = await run(editor, [{ op: 'reference.source.add', title: '  ' }], { app: ctx })
    expect(noTitle.ok).toBe(false)
    expect(noTitle.results[0].error).toContain('title is required')

    const citation = await run(
      editor,
      [{ op: 'reference.citation.insert', sourceId: 'Wang2024', afterBlockIndex: 0 }],
      { app: ctx },
    )
    expect(citation.ok).toBe(true)
    expect(JSON.stringify(editor.getJSON().content![0])).toContain('(Wang, 2024)')

    const bibliography = await run(
      editor,
      [{ op: 'reference.bibliography.insert', afterBlockIndex: 0 }],
      { app: ctx },
    )
    expect(bibliography.ok).toBe(true)
    expect(bibliography.results[0].record).toMatchObject({ lines: 2 })
    const texts = nodeTexts(editor)
    expect(texts.some((text) => text.includes('Deep Work'))).toBe(true)

    const emptyApp = makeFakeApp(createEditor([para('x')]))
    const noSources = await run(
      emptyApp.ctx.review().editor!,
      [{ op: 'reference.bibliography.insert' }],
      { app: emptyApp.ctx },
    )
    expect(noSources.ok).toBe(false)
    expect(noSources.results[0].error).toContain('no bibliography sources')

    const dry = await run(
      editor,
      [{ op: 'reference.source.add', title: 'Another' }],
      { app: ctx, dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(state.sources).toHaveLength(2)
  })

  it('reference.index.mark / insert build an INDEX field from marked terms', async () => {
    const editor = createEditor([para('alpha beta'), para('gamma')])
    const marked = await run(editor, [
      { op: 'reference.index.mark', blockIndex: 0, term: 'alpha' },
      { op: 'reference.index.mark', blockIndex: 1, term: 'gamma' },
    ])
    expect(marked.ok).toBe(true)
    expect(JSON.stringify(editor.getJSON())).toContain('docXeMark')

    const inserted = await run(editor, [{ op: 'reference.index.insert', afterBlockIndex: 1 }])
    expect(inserted.ok).toBe(true)
    expect(inserted.results[0].record).toMatchObject({ terms: 2 })
    expect(JSON.stringify(editor.getJSON())).toContain('INDEX')

    const noEntries = await run(createEditor([para('plain')]), [{ op: 'reference.index.insert' }])
    expect(noEntries.ok).toBe(false)
    expect(noEntries.results[0].error).toContain('no index entries')
  })
})

describe('review ops', () => {
  it('review.trackChanges toggles the editor tracker and supports dryRun', async () => {
    const editor = createEditor([para('body')])
    const { ctx, state } = makeFakeApp(editor)
    const dry = await run(editor, [{ op: 'review.trackChanges', enabled: true }], {
      app: ctx,
      dryRun: true,
    })
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(state.trackChanges).toBe(false)

    const on = await run(editor, [{ op: 'review.trackChanges', enabled: true }], { app: ctx })
    expect(on.ok).toBe(true)
    expect(state.trackChanges).toBe(true)
    const storage = editor.storage as unknown as { trackChanges: { enabled: boolean } }
    expect(storage.trackChanges.enabled).toBe(true)

    const off = await run(editor, [{ op: 'review.trackChanges', enabled: false }], { app: ctx })
    expect(off.ok).toBe(true)
    expect(storage.trackChanges.enabled).toBe(false)
  })

  it('review.accept / review.reject require confirm and take [rN] ids from read_revisions', async () => {
    const editor = createEditor([revisionParagraph('one'), revisionParagraph('two')])
    expect(pendingRevisions(editor)).toBe(2)

    const gated = await run(editor, [{ op: 'review.accept', all: true }])
    expect(gated.ok).toBe(false)
    expect(gated.results[0].error).toContain('CONFIRMATION_REQUIRED')
    expect(pendingRevisions(editor)).toBe(2)

    const badId = await run(editor, [
      { op: 'review.accept', revisionIds: ['r9'], confirm: true },
    ])
    expect(badId.ok).toBe(false)
    expect(badId.results[0].error).toContain('do not exist')

    const both = await run(editor, [
      { op: 'review.accept', revisionIds: ['r0'], all: true, confirm: true },
    ])
    expect(both.ok).toBe(false)
    expect(both.results[0].error).toContain('either revisionIds or all:true')

    const dryEditor = createEditor([revisionParagraph('x')])
    const dry = await run(
      dryEditor,
      [{ op: 'review.accept', all: true, confirm: true }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(pendingRevisions(dryEditor)).toBe(1)

    const accepted = await run(editor, [
      { op: 'review.accept', revisionIds: ['r0'], confirm: true },
    ])
    expect(accepted.ok).toBe(true)
    expect(accepted.results[0].record).toMatchObject({ mode: 'accept', revisionIds: ['r0'] })
    expect(pendingRevisions(editor)).toBe(1)

    const rejected = await run(editor, [{ op: 'review.reject', all: true, confirm: true }])
    expect(rejected.ok).toBe(true)
    expect(pendingRevisions(editor)).toBe(0)
  })

  it('review.protect is confirm-gated, hashes passwords and never echoes them', async () => {
    const editor = createEditor([para('body')])
    const { ctx, state } = makeFakeApp(editor)

    const gated = await run(editor, [{ op: 'review.protect', mode: 'readOnly' }], { app: ctx })
    expect(gated.ok).toBe(false)
    expect(gated.results[0].error).toContain('CONFIRMATION_REQUIRED')
    expect(state.protectionRequests).toHaveLength(0)

    const readOnly = await run(
      editor,
      [{ op: 'review.protect', mode: 'readOnly', confirm: true }],
      { app: ctx },
    )
    expect(readOnly.ok).toBe(true)
    expect(state.protectionRequests[0]).toEqual({
      protection: { edit: 'readOnly', enforced: true },
    })
    expect(JSON.stringify(readOnly)).not.toContain('password":')

    const secret = 'correct horse battery staple'
    const protectedWithPassword = await run(
      editor,
      [{ op: 'review.protect', mode: 'comments', password: secret, confirm: true }],
      { app: ctx },
    )
    expect(protectedWithPassword.ok).toBe(true)
    const payload = state.protectionRequests[1]
    expect(payload.protection).toMatchObject({ edit: 'comments', enforced: true })
    expect(typeof payload.protection?.hash).toBe('string')
    expect(payload.protection?.salt).toBeTruthy()
    expect(payload.protection?.spinCount).toBe(100000)
    // password must never leak into the op result
    expect(JSON.stringify(protectedWithPassword)).not.toContain(secret)
    expect(protectedWithPassword.results[0].record).toMatchObject({
      mode: 'comments',
      passwordSet: true,
    })

    const recommended = await run(
      editor,
      [{ op: 'review.protect', mode: 'readOnlyRecommended', confirm: true }],
      { app: ctx },
    )
    expect(recommended.ok).toBe(true)
    expect(state.protectionRequests[2]).toEqual({
      writeProtection: { recommended: true },
    })

    const badMode = await run(
      editor,
      [{ op: 'review.protect', mode: 'nope', confirm: true }],
      { app: ctx },
    )
    expect(badMode.ok).toBe(false)
    expect(badMode.results[0].error).toContain('mode must be one of')

    const pwWithRecommended = await run(
      editor,
      [{ op: 'review.protect', mode: 'readOnlyRecommended', password: 'x', confirm: true }],
      { app: ctx },
    )
    expect(pwWithRecommended.ok).toBe(false)

    const dry = await run(
      editor,
      [{ op: 'review.protect', mode: 'forms', password: 'dry-secret', confirm: true }],
      { app: ctx, dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(state.protectionRequests).toHaveLength(3)
    expect(JSON.stringify(dry)).not.toContain('dry-secret')
  })

  it('review.encrypt is confirm-gated, updates the open password and redacts it', async () => {
    const editor = createEditor([para('body')])
    const { ctx, state } = makeFakeApp(editor)

    const gated = await run(editor, [{ op: 'review.encrypt', password: 'secret-1' }], { app: ctx })
    expect(gated.ok).toBe(false)
    expect(gated.results[0].error).toContain('CONFIRMATION_REQUIRED')
    expect(state.passwordRequests).toHaveLength(0)

    const set = await run(
      editor,
      [{ op: 'review.encrypt', password: 'secret-1', confirm: true }],
      { app: ctx },
    )
    expect(set.ok).toBe(true)
    expect(state.passwordRequests).toEqual(['secret-1'])
    expect(JSON.stringify(set)).not.toContain('secret-1')
    expect(set.results[0].record).toMatchObject({ passwordSet: true })

    const cleared = await run(
      editor,
      [{ op: 'review.encrypt', password: null, confirm: true }],
      { app: ctx },
    )
    expect(cleared.ok).toBe(true)
    expect(state.passwordRequests).toEqual(['secret-1', null])

    // main-process rejection must not echo any error payload (could contain the password)
    const failing = makeFakeApp(createEditor([para('x')]))
    failing.ctx.setOpenPassword = async () => ({ ok: false, error: 'ipc said secret-2' })
    const failed = await run(
      failing.ctx.review().editor!,
      [{ op: 'review.encrypt', password: 'secret-2', confirm: true }],
      { app: failing.ctx },
    )
    expect(failed.ok).toBe(false)
    expect(JSON.stringify(failed)).not.toContain('secret-2')
    expect(failed.results[0].error).toContain('rejected')
  })

  it('review.findReplace replaces through the apply_commands engine and validates unsupported flags', async () => {
    const editor = createEditor([para('Hello world'), para('hello again')])
    const outcome = await run(editor, [
      { op: 'review.findReplace', find: 'hello', replace: 'Hi', matchCase: false },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].record).toMatchObject({ matches: 2 })
    expect(nodeTexts(editor)).toEqual(['Hi world', 'Hi again'])

    const before = snapshot(editor)
    const dry = await run(
      editor,
      [{ op: 'review.findReplace', find: 'Hi', replace: 'Bye' }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(snapshot(editor)).toBe(before)

    const noMatch = await run(editor, [{ op: 'review.findReplace', find: 'absent', replace: 'x' }])
    expect(noMatch.ok).toBe(false)
    expect(noMatch.results[0].error).toContain('no occurrence')

    // whole-word and first-only replace are supported (see the B5 find/replace tests)
  })
})

describe('B4 registry + isolation', () => {
  it('registers every B4 op in the guide with its group', async () => {
    const defs = describeOperations()
    const names = defs.map((def) => def.name)
    for (const name of [
      'reference.bookmark.insert',
      'reference.bookmark.remove',
      'reference.crossRef.insert',
      'reference.caption.insert',
      'reference.source.add',
      'reference.citation.insert',
      'reference.bibliography.insert',
      'reference.index.mark',
      'reference.index.insert',
      'review.trackChanges',
      'review.accept',
      'review.reject',
      'review.protect',
      'review.encrypt',
      'review.findReplace',
    ]) {
      expect(names).toContain(name)
      expect(OPERATIONS_GUIDE).toContain(name)
    }
    expect(new Set(defs.map((def) => def.group))).toEqual(
      new Set([
        'table',
        'page',
        'section',
        'notes',
        'field',
        'format',
        'object',
        'reference',
        'review',
      ]),
    )
  })

  it('keeps per-op isolation across B4 ops (a confirm-gated failure does not abort later ops)', async () => {
    const editor = createEditor([para('one'), para('two')])
    const outcome = await run(editor, [
      { op: 'review.encrypt', password: 'x' }, // no confirm -> CONFIRMATION_REQUIRED
      { op: 'reference.bookmark.insert', blockIndex: 0, name: 'Keep' },
      { op: 'review.findReplace', find: 'two', replace: 'TWO' },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.failed).toBe(1)
    expect(outcome.applied).toBe(2)
    expect(editor.getJSON().content![0].attrs?.bookmarks).toEqual(['Keep'])
    expect(nodeTexts(editor)[1]).toBe('TWO')
  })
})

// ── B5 operation-area parity tests ─────────────────────────────────────────

describe('format.applyCharacterStyle', () => {
  it('applies and toggles the emphasis/strong presets with the theme accent', async () => {
    const editor = createEditor([para('Styled text')])
    const applied = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [0] }, preset: 'emphasis', mode: 'apply' },
    ])
    expect(applied.ok).toBe(true)
    let block = editor.getJSON().content![0] as unknown as JsonNode
    let marks = block.content!.flatMap((child) => child.marks ?? [])
    expect(marks.some((mark) => mark.type === 'italic')).toBe(true)
    expect(marks.find((mark) => mark.type === 'docTextStyle')?.attrs?.color).toBe('4472C4')

    const toggledOff = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [0] }, preset: 'emphasis' },
    ])
    expect(toggledOff.ok).toBe(true)
    block = editor.getJSON().content![0] as unknown as JsonNode
    marks = block.content!.flatMap((child) => child.marks ?? [])
    expect(marks.some((mark) => mark.type === 'italic')).toBe(false)
    expect(marks.find((mark) => mark.type === 'docTextStyle')?.attrs?.color ?? null).toBeNull()

    const strong = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [0] }, preset: 'strong', mode: 'apply' },
    ])
    expect(strong.ok).toBe(true)
    block = editor.getJSON().content![0] as unknown as JsonNode
    marks = block.content!.flatMap((child) => child.marks ?? [])
    expect(marks.some((mark) => mark.type === 'bold')).toBe(true)
  })

  it('applies and toggles a document character styleId', async () => {
    const editor = createEditor([para('text')])
    const applied = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [0] }, styleId: 'Emphasis', mode: 'apply' },
    ])
    expect(applied.ok).toBe(true)
    let marks = (editor.getJSON().content![0] as unknown as JsonNode).content!.flatMap(
      (child) => child.marks ?? [],
    )
    expect(marks.find((mark) => mark.type === 'docTextStyle')?.attrs?.styleId).toBe('Emphasis')

    const toggled = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [0] }, styleId: 'Emphasis' },
    ])
    expect(toggled.ok).toBe(true)
    marks = (editor.getJSON().content![0] as unknown as JsonNode).content!.flatMap(
      (child) => child.marks ?? [],
    )
    expect(marks.some((mark) => mark.type === 'docTextStyle')).toBe(false)
  })

  it('validates preset/styleId/target and stays inert under dryRun', async () => {
    const editor = createEditor([para('text'), para('')])
    const before = snapshot(editor)
    const dry = await run(
      editor,
      [{ op: 'format.applyCharacterStyle', target: { blockIndexes: [0] }, preset: 'strong' }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(snapshot(editor)).toBe(before)

    const both = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [0] }, preset: 'strong', styleId: 'X' },
    ])
    expect(both.ok).toBe(false)
    expect(both.results[0].error).toContain('exactly one of preset or styleId')
    const neither = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [0] } },
    ])
    expect(neither.ok).toBe(false)
    const badPreset = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [0] }, preset: 'fancy' },
    ])
    expect(badPreset.ok).toBe(false)
    const emptyText = await run(editor, [
      { op: 'format.applyCharacterStyle', target: { blockIndexes: [1] }, preset: 'strong' },
    ])
    expect(emptyText.ok).toBe(false)
    expect(emptyText.results[0].error).toContain('no text')
  })
})

describe('table.setProperties look options', () => {
  it('writes the Table Styles look flags through the real table-look command', async () => {
    const editor = createEditor(fixture3x3())
    const outcome = await run(editor, [
      {
        op: 'table.setProperties',
        blockIndex: TABLE_BLOCK,
        look: { firstRow: false, bandedRows: false, firstColumn: true },
      },
    ])
    expect(outcome.ok).toBe(true)
    expect(jsonTable(editor).attrs).toMatchObject({
      tblLook: { firstRow: false, firstColumn: true, bandedRows: false },
      tblLookEdited: true,
    })

    const dry = await run(
      editor,
      [{ op: 'table.setProperties', blockIndex: TABLE_BLOCK, look: { lastRow: true } }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(jsonTable(editor).attrs?.tblLook).toMatchObject({ lastRow: false })

    const unknown = await run(editor, [
      { op: 'table.setProperties', blockIndex: TABLE_BLOCK, look: { middleRow: true } },
    ])
    expect(unknown.ok).toBe(false)
    expect(unknown.results[0].error).toContain('unknown look option')
    const nonBool = await run(editor, [
      { op: 'table.setProperties', blockIndex: TABLE_BLOCK, look: { firstRow: 'yes' } },
    ])
    expect(nonBool.ok).toBe(false)
    expect(nonBool.results[0].error).toContain('must be a boolean')
  })
})

describe('format.setTabStops', () => {
  function tabStopsOf(editor: Editor): Array<Record<string, unknown>> | null {
    const raw = (editor.getJSON().content![0] as unknown as JsonNode).attrs?.tabStops as string | null
    return raw ? (JSON.parse(raw) as Array<Record<string, unknown>>) : null
  }

  it('replaces normal stops while preserving rel/ptab stops, and [] clears them', async () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: {
          docxIndex: null,
          tabStops: JSON.stringify([
            { pos: 720, val: 'left' },
            { pos: 2000, val: 'right', rel: 'margin' },
          ]),
        },
        content: [{ type: 'text', text: 'tabbed' }],
      },
    ])
    const outcome = await run(editor, [
      {
        op: 'format.setTabStops',
        target: { blockIndexes: [0] },
        stops: [{ posTwips: 1440, val: 'center', leader: 'dot' }],
      },
    ])
    expect(outcome.ok).toBe(true)
    let stops = tabStopsOf(editor)!
    expect(stops).toHaveLength(2)
    expect(stops[0]).toMatchObject({ pos: 1440, val: 'center', leader: 'dot' })
    expect(stops[1]).toMatchObject({ pos: 2000, val: 'right', rel: 'margin' })

    const cleared = await run(editor, [
      { op: 'format.setTabStops', target: { blockIndexes: [0] }, stops: [] },
    ])
    expect(cleared.ok).toBe(true)
    stops = tabStopsOf(editor)!
    expect(stops).toEqual([{ pos: 2000, val: 'right', rel: 'margin' }])

    const duplicate = await run(editor, [
      {
        op: 'format.setTabStops',
        target: { blockIndexes: [0] },
        stops: [
          { posTwips: 720, val: 'left' },
          { posTwips: 720, val: 'right' },
        ],
      },
    ])
    expect(duplicate.ok).toBe(false)
    expect(duplicate.results[0].error).toContain('duplicate tab stop')

    const badVal = await run(editor, [
      { op: 'format.setTabStops', target: { blockIndexes: [0] }, stops: [{ posTwips: 720, val: 'justify' }] },
    ])
    expect(badVal.ok).toBe(false)
    expect(badVal.results[0].error).toContain('stop.val must be one of')
  })
})

describe('format.insertBreak', () => {
  it('inserts line and column breaks at the start/end of a block', async () => {
    const editor = createEditor([para('hello')])
    const line = await run(editor, [
      { op: 'format.insertBreak', kind: 'line', blockIndex: 0, at: 'end' },
    ])
    expect(line.ok).toBe(true)
    let content = (editor.getJSON().content![0] as unknown as JsonNode).content ?? []
    expect(content[content.length - 1]).toMatchObject({ type: 'hardBreak' })
    expect(content[content.length - 1].attrs?.colBreak ?? false).toBe(false)

    const column = await run(editor, [
      { op: 'format.insertBreak', kind: 'column', blockIndex: 0, at: 'start' },
    ])
    expect(column.ok).toBe(true)
    content = (editor.getJSON().content![0] as unknown as JsonNode).content ?? []
    expect(content[0]).toMatchObject({ type: 'hardBreak', attrs: { colBreak: true } })

    const badKind = await run(editor, [
      { op: 'format.insertBreak', kind: 'page', blockIndex: 0 },
    ])
    expect(badKind.ok).toBe(false)
    expect(badKind.results[0].error).toContain('kind must be')

    const before = snapshot(editor)
    const dry = await run(
      editor,
      [{ op: 'format.insertBreak', kind: 'line', blockIndex: 0 }],
      { dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(snapshot(editor)).toBe(before)
  })
})

describe('format.applyListStyle', () => {
  it('creates a definition and converts paragraphs, skipping headings', async () => {
    const editor = createEditor([para('one'), para('two'), headingJson('title')])
    const { ctx, state } = makeFakeApp(editor)
    const outcome = await run(
      editor,
      [{ op: 'format.applyListStyle', target: { blockIndexes: [0, 1, 2] }, kind: 'ordered' }],
      { app: ctx },
    )
    expect(outcome.ok).toBe(true)
    const blocks = editor.getJSON().content! as unknown as JsonNode[]
    expect(blocks[0].type).toBe('docListItem')
    expect(blocks[1].type).toBe('docListItem')
    expect(blocks[0].attrs).toMatchObject({ kind: 'ordered', numId: '6', ilvl: 0 })
    expect(blocks[2].type).toBe('docHeading')
    expect(outcome.results[0].record).toMatchObject({ blocks: 2, skipped: 1, kind: 'ordered' })
    expect(state.pendingNumbering.newDefs).toHaveLength(1)
  })

  it('accepts custom levels and supports the bullet fallback', async () => {
    const customEditor = createEditor([para('item')])
    const customApp = makeFakeApp(customEditor)
    const custom = await run(
      customEditor,
      [
        {
          op: 'format.applyListStyle',
          target: { blockIndexes: [0] },
          kind: 'bullet',
          levels: [{ level: 0, numFmt: 'bullet', lvlText: '◆' }],
        },
      ],
      { app: customApp.ctx },
    )
    expect(custom.ok).toBe(true)
    expect(customEditor.getJSON().content![0].attrs).toMatchObject({ kind: 'bullet' })
    expect(customApp.state.pendingNumbering.newDefs[0].levels?.[0]?.lvlText).toBe('◆')

    const bulletEditor = createEditor([para('item')])
    const bulletApp = makeFakeApp(bulletEditor)
    const fallback = await run(
      bulletEditor,
      [{ op: 'format.applyListStyle', target: { blockIndexes: [0] }, kind: 'bullet' }],
      { app: bulletApp.ctx },
    )
    expect(fallback.ok).toBe(true)
    expect(bulletEditor.getJSON().content![0].attrs).toMatchObject({ kind: 'bullet' })

    const badLevels = await run(
      customEditor,
      [
        {
          op: 'format.applyListStyle',
          target: { blockIndexes: [0] },
          kind: 'ordered',
          levels: [{ level: 1, numFmt: 'decimal' }],
        },
      ],
      { app: customApp.ctx },
    )
    expect(badLevels.ok).toBe(false)
    expect(badLevels.results[0].error).toContain('contiguous from level 0')

    const plainEditor = createEditor([para('x')])
    const noApp = await run(plainEditor, [
      { op: 'format.applyListStyle', target: { blockIndexes: [0] }, kind: 'ordered' },
    ])
    expect(noApp.ok).toBe(false)
    expect(noApp.results[0].error).toContain('needs the app context')

    const dryEditor = createEditor([para('dry')])
    const dryApp = makeFakeApp(dryEditor)
    const before = snapshot(dryEditor)
    const dry = await run(
      dryEditor,
      [{ op: 'format.applyListStyle', target: { blockIndexes: [0] }, kind: 'ordered' }],
      { app: dryApp.ctx, dryRun: true },
    )
    expect(dry.results[0]).toMatchObject({ ok: true, planned: true })
    expect(snapshot(dryEditor)).toBe(before)
    expect(dryApp.state.pendingNumbering.newDefs).toHaveLength(0)
  })
})

describe('object.equation.update', () => {
  it('re-edits an existing block equation in place without inserting a new one', async () => {
    const editor = createEditor([para('first')])
    await run(editor, [{ op: 'object.equation.insert', latex: 'a^2 + b^2 = c^2' }])
    const before = editor.getJSON().content! as unknown as JsonNode[]
    const equationIndex = before.findIndex((block) => block.attrs?.formulaDisplay)
    expect(equationIndex).toBeGreaterThanOrEqual(0)
    const countBefore = before.length

    const updated = await run(editor, [
      { op: 'object.equation.update', blockIndex: equationIndex, latex: 'x = 2' },
    ])
    expect(updated.ok).toBe(true)
    const after = editor.getJSON().content! as unknown as JsonNode[]
    expect(after).toHaveLength(countBefore)
    expect(after[equationIndex].attrs?.formulaDisplay).toMatchObject({ latex: 'x = 2' })
  })

  it('updates an inline equation by inlineIndex', async () => {
    const editor = createEditor([para('inline here')])
    await run(editor, [
      { op: 'object.equation.insert', latex: 'x = 1', inline: true, afterBlockIndex: 0 },
    ])
    const blocks = editor.getJSON().content! as unknown as JsonNode[]
    const inlineBlockIndex = blocks.findIndex((block) => JSON.stringify(block).includes('docInlineMath'))
    expect(inlineBlockIndex).toBeGreaterThanOrEqual(0)
    const updated = await run(editor, [
      { op: 'object.equation.update', blockIndex: inlineBlockIndex, latex: 'y = 3', inlineIndex: 0 },
    ])
    expect(updated.ok).toBe(true)
    expect(JSON.stringify(editor.getJSON().content![inlineBlockIndex])).toContain('"latex":"y = 3"')
  })

  it('rejects non-equation blocks, out-of-range inline indexes and bad LaTeX', async () => {
    const editor = createEditor([para('plain')])
    const notEquation = await run(editor, [
      { op: 'object.equation.update', blockIndex: 0, latex: 'x = 1' },
    ])
    expect(notEquation.ok).toBe(false)
    expect(notEquation.results[0].error).toContain('not an equation block')

    const badLatex = await run(editor, [
      { op: 'object.equation.update', blockIndex: 0, latex: '\\frac{1' },
    ])
    expect(badLatex.ok).toBe(false)
    expect(badLatex.results[0].error).toContain('invalid LaTeX')

    const noInline = await run(editor, [
      { op: 'object.equation.update', blockIndex: 0, latex: 'x = 1', inlineIndex: 0 },
    ])
    expect(noInline.ok).toBe(false)
    expect(noInline.results[0].error).toContain('no inline equation')
  })
})

describe('review.findReplace parity (B5)', () => {
  it('supports whole-word matches (ASCII boundaries)', async () => {
    const editor = createEditor([para('cat concatenate cat'), para('dog cat')])
    const outcome = await run(editor, [
      { op: 'review.findReplace', find: 'cat', replace: 'X', wholeWord: true },
    ])
    expect(outcome.ok).toBe(true)
    expect(nodeTexts(editor)).toEqual(['X concatenate X', 'dog X'])
    expect(outcome.results[0].record).toMatchObject({ matches: 3 })
  })

  it('supports replacing only the first occurrence', async () => {
    const editor = createEditor([para('cat cat cat')])
    const outcome = await run(editor, [
      { op: 'review.findReplace', find: 'cat', replace: 'X', all: false },
    ])
    expect(outcome.ok).toBe(true)
    expect(nodeTexts(editor)).toEqual(['X cat cat'])
    expect(outcome.results[0].record).toMatchObject({ matches: 1 })

    const dry = await run(
      editor,
      [{ op: 'review.findReplace', find: 'cat', replace: 'Y', all: false }],
      { dryRun: true },
    )
    expect(dry.results[0].record).toMatchObject({ matches: 1 })
    expect(nodeTexts(editor)).toEqual(['X cat cat'])
  })
})

describe('B5 registry', () => {
  it('registers the B5 ops in the guide', async () => {
    const names = describeOperations().map((def) => def.name)
    for (const name of [
      'format.applyCharacterStyle',
      'format.setTabStops',
      'format.insertBreak',
      'format.applyListStyle',
      'object.equation.update',
    ]) {
      expect(names).toContain(name)
      expect(OPERATIONS_GUIDE).toContain(name)
    }
    // table look + find/replace parity are extensions of existing ops
    expect(OPERATIONS_GUIDE).toContain('look')
    expect(OPERATIONS_GUIDE).toContain('wholeWord')
  })
})
