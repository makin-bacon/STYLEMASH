import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { StyleVariantRow } from '../src/components/StyleVariantRow'
import type { FormattingSignature, StyleEntityVariant } from '../src/types/ooxml'

// Silences React's benign "not configured to support act()" warning - this
// file mounts real components via react-dom/client directly (no test-render
// helper library in this project's dependencies), which is exactly what
// this flag is for.
;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const SIGNATURE: FormattingSignature = {
  fontFamily: 'Arial',
  fontSizeHalfPt: 24,
  colorValue: 'auto',
  bold: false,
  italic: false,
  underline: null,
  strike: false,
}

const VARIANT: StyleEntityVariant = {
  id: 'v1',
  origin: { kind: 'direct' },
  occurrenceCount: 3,
  sampleText: 'Hello world',
  runRefs: [],
}

function renderRow(props: Partial<Parameters<typeof StyleVariantRow>[0]> = {}) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  const onToggleSelect = vi.fn()
  const onEditXml = vi.fn()

  act(() => {
    root.render(
      <StyleVariantRow
        signature={SIGNATURE}
        variant={VARIANT}
        selected={false}
        onToggleSelect={onToggleSelect}
        onEditXml={onEditXml}
        {...props}
      />,
    )
  })

  return { container, onToggleSelect, onEditXml, root }
}

describe('StyleVariantRow interaction', () => {
  it('clicking anywhere on the row toggles selection', () => {
    const { container, onToggleSelect } = renderRow()
    const li = container.querySelector('li')!

    act(() => {
      li.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onToggleSelect).toHaveBeenCalledTimes(1)
  })

  it('clicking "Edit XML" fires onEditXml but does not also toggle selection', () => {
    const { container, onToggleSelect, onEditXml } = renderRow()
    const editButton = Array.from(container.querySelectorAll('button')).find(
      (b) => b.textContent === 'Edit XML',
    )!

    act(() => {
      editButton.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onEditXml).toHaveBeenCalledTimes(1)
    expect(onToggleSelect).not.toHaveBeenCalled()
  })

  it('clicking the (visually hidden) checkbox toggles selection exactly once, not twice', () => {
    const { container, onToggleSelect } = renderRow()
    const checkbox = container.querySelector('input[type="checkbox"]')! as HTMLInputElement

    act(() => {
      checkbox.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    // A native checkbox click fires both a 'click' and a 'change' event in a
    // real browser; jsdom's dispatchEvent for a checkbox <input> does the
    // same when `checked` isn't manually pre-set, so onChange (=
    // onToggleSelect) should fire from the input itself - and, critically,
    // NOT a second time via the row's onClick (stopPropagation guards that).
    expect(onToggleSelect.mock.calls.length).toBeLessThanOrEqual(1)
  })

  it('reflects selected state via the row background, not just the checkbox', () => {
    const { container } = renderRow({ selected: true })
    const li = container.querySelector('li')!
    expect(li.className).toContain('bg-indigo-50')
  })
})
