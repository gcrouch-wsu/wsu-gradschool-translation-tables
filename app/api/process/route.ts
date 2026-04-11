import { NextRequest, NextResponse } from 'next/server'

interface TranslationRow {
  Input: string
  Output: string
}

interface EvaluationQuestion {
  id: string
  label: string
  key: string
  helpText: string
  options: string[]
  conditional: boolean
}

interface EvaluationSection {
  id: string
  title: string
  introduction: string
  conditional: boolean
  questions: EvaluationQuestion[]
}

interface EvaluationForm {
  id: string
  formTitle: string
  displayName: string
  formKey: string
  sections: EvaluationSection[]
}

const NAV_NOISE = new Set([
  'Settings',
  'Search settings...',
  'Portals',
  'Contacts',
  'Application Setup',
  'Application Review',
  'Application Summary View',
  'Application Timeline Summary',
  'Evaluation Forms',
  'Decisions',
  'Decision Letters',
  'Phases',
  'Application Segments',
  'Review Routing Tables',
  'Automation',
  'Workflows',
  'Marketing',
  'Calendar',
  'Import/Export',
  'System',
  'Greg Crouch',
  'Save Changes',
  'menumore_horiz',
  'KeyTry it play_circle_outline',
  'Add Question',
  'add Add Section',
  'add Add Calculated Field',
  'Introduction',
])

function isQuestionStart(line: string): boolean {
  return line.startsWith('editcontent_copymenu ')
}

function isNoiseLine(line: string): boolean {
  const normalized = String(line || '').trim()
  if (!normalized) return true
  if (NAV_NOISE.has(normalized)) return true
  if (normalized.includes('expand_more') || normalized.includes('expand_less')) return true
  if (normalized.startsWith('arrow_back')) return true
  if (normalized.startsWith('Press Enter to activate drag mode')) return true
  if (normalized === 'Organization 3') return true
  return false
}

function isMetadataMarker(line: string): boolean {
  if (line === 'Section Title') return true
  if (line === 'visibility Conditional visibility') return true
  if (isQuestionStart(line)) return true
  return false
}

function slugify(value: string): string {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80)
}

function parseQuestion(lines: string[], startIndex: number, fallbackIndex: number): { question: EvaluationQuestion; nextIndex: number } {
  const first = lines[startIndex] || ''
  const label = first.replace(/^editcontent_copymenu\s+/, '').trim() || `Question ${fallbackIndex}`
  const bodyLines: string[] = []
  let key = ''
  let conditional = false
  let i = startIndex + 1

  while (i < lines.length) {
    const line = String(lines[i] || '').trim()
    if (!line) {
      i += 1
      continue
    }
    if (line === 'Section Title' || line === 'Save Changes' || isQuestionStart(line)) {
      break
    }
    if (line.startsWith('Press Enter to activate drag mode')) {
      i += 1
      break
    }
    if (line === 'visibility Conditional visibility') {
      conditional = true
      i += 1
      continue
    }
    const keyMatch = line.match(/^Key:\s*(.+)$/i)
    if (keyMatch) {
      key = keyMatch[1].trim()
      i += 1
      continue
    }
    if (!isNoiseLine(line)) {
      bodyLines.push(line)
    }
    i += 1
  }

  const helpLines: string[] = []
  const optionLines: string[] = []
  bodyLines.forEach((line) => {
    if (/^Response may contain up to /i.test(line) || /^Select /i.test(line) || /:$/i.test(line) || /\.$/.test(line)) {
      helpLines.push(line)
      return
    }
    if (line.includes('\t')) {
      optionLines.push(
        ...line
          .split('\t')
          .map((value) => value.trim())
          .filter(Boolean)
      )
      return
    }
    if (line.length <= 110) {
      optionLines.push(line)
      return
    }
    helpLines.push(line)
  })

  const dedupedOptions = Array.from(new Set(optionLines.map((line) => line.trim()).filter(Boolean)))
  const questionKey = key || slugify(label) || `question-${fallbackIndex}`
  return {
    question: {
      id: `${questionKey}-${fallbackIndex}`,
      label,
      key: questionKey,
      helpText: helpLines.join('\n').trim(),
      options: dedupedOptions,
      conditional,
    },
    nextIndex: i,
  }
}

function parseEvaluationForms(content: string): EvaluationForm[] {
  const lines = content
    .split('\n')
    .map((line) => line.replace(/\r/g, '').trim())
    .filter((line) => line.length > 0)

  const saveChangesIndexes: number[] = []
  lines.forEach((line, index) => {
    if (line === 'Save Changes') {
      saveChangesIndexes.push(index)
    }
  })

  const ranges: Array<{ start: number; end: number }> = []
  if (saveChangesIndexes.length === 0) {
    ranges.push({ start: 0, end: lines.length })
  } else {
    saveChangesIndexes.forEach((startIndex, idx) => {
      const start = startIndex + 1
      const end = idx + 1 < saveChangesIndexes.length ? saveChangesIndexes[idx + 1] : lines.length
      if (start < end) {
        ranges.push({ start, end })
      }
    })
  }

  const parsedForms: EvaluationForm[] = []

  ranges.forEach((range, rangeIndex) => {
    const segment = lines.slice(range.start, range.end)
    const firstSectionIndex = segment.findIndex((line) => line === 'Section Title')
    const metadataBoundary = firstSectionIndex >= 0 ? firstSectionIndex : Math.min(segment.length, 25)
    const metadataRegion = segment.slice(0, metadataBoundary)
    const metadataCandidates = metadataRegion.filter((line) => !isNoiseLine(line) && !isMetadataMarker(line) && line !== 'Name')

    const formTitle = metadataCandidates[0] || `Imported Evaluation Form ${rangeIndex + 1}`
    const displayName = metadataCandidates[1] || formTitle
    let formKey = slugify(displayName || formTitle)

    const nameIndex = segment.indexOf('Name')
    if (nameIndex >= 0) {
      for (let i = nameIndex + 1; i < segment.length; i += 1) {
        const candidate = segment[i]
        if (isNoiseLine(candidate) || isMetadataMarker(candidate) || candidate === 'Name') continue
        formKey = candidate
        break
      }
    }

    const sections: EvaluationSection[] = []
    let pendingSectionConditional = false
    let currentSection: EvaluationSection | null = null
    let i = 0

    while (i < segment.length) {
      const line = segment[i]
      if (line === 'Save Changes') break

      if (line === 'visibility Conditional visibility') {
        pendingSectionConditional = true
        i += 1
        continue
      }

      if (line === 'Section Title') {
        let sectionTitle = `Section ${sections.length + 1}`
        for (let j = i - 1; j >= 0; j -= 1) {
          const candidate = segment[j]
          if (isNoiseLine(candidate)) continue
          if (candidate === 'Section Title' || candidate === 'visibility Conditional visibility' || isQuestionStart(candidate)) continue
          sectionTitle = candidate
          break
        }

        const section: EvaluationSection = {
          id: `${slugify(sectionTitle) || 'section'}-${sections.length + 1}`,
          title: sectionTitle,
          introduction: '',
          conditional: pendingSectionConditional,
          questions: [],
        }
        pendingSectionConditional = false
        sections.push(section)
        currentSection = section
        i += 1

        if (i < segment.length && segment[i] === 'Introduction') {
          i += 1
        }

        const introLines: string[] = []
        while (i < segment.length) {
          const introLine = segment[i]
          if (introLine === 'Section Title' || introLine === 'Save Changes' || isQuestionStart(introLine)) break
          if (introLine === 'visibility Conditional visibility' || introLine === 'Add Question' || introLine === 'add Add Section' || introLine === 'add Add Calculated Field' || introLine === 'menumore_horiz') break
          if (!isNoiseLine(introLine)) {
            introLines.push(introLine)
          }
          i += 1
        }
        section.introduction = introLines.join('\n').trim()
        continue
      }

      if (isQuestionStart(line)) {
        if (!currentSection) {
          const defaultSection: EvaluationSection = {
            id: `section-${sections.length + 1}`,
            title: `Section ${sections.length + 1}`,
            introduction: '',
            conditional: false,
            questions: [],
          }
          sections.push(defaultSection)
          currentSection = defaultSection
        }

        const parsedQuestion = parseQuestion(segment, i, currentSection.questions.length + 1)
        currentSection.questions.push(parsedQuestion.question)
        i = parsedQuestion.nextIndex
        continue
      }

      i += 1
    }

    const hasQuestion = sections.some((section) => section.questions.length > 0)
    if (hasQuestion) {
      parsedForms.push({
        id: `${slugify(formTitle) || 'evaluation-form'}-${rangeIndex + 1}`,
        formTitle,
        displayName,
        formKey,
        sections,
      })
    }
  })

  return parsedForms
}

function parseTxtContent(content: string): TranslationRow[] {
  /**
   * Parses the raw text content using a strict anchor-based approach.
   * - Input anchor: Line must be exactly 'Input'.
   * - Output anchor: Line must start with 'Output' but not be 'Output Type'.
   */
  const lines = content
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0)

  const inputs: string[] = []
  const outputs: string[] = []

  // We start from 1 because we need a preceding line (i-1)
  for (let i = 1; i < lines.length; i++) {
    const lineClean = lines[i]

    // Strict Input Anchor: exactly "Input"
    if (lineClean === 'Input') {
      inputs.push(lines[i - 1])
    }
    // Strict Output Anchor: Starts with "Output", excludes headers like "Output Type"
    else if (lineClean.startsWith('Output') && lineClean !== 'Output Type') {
      outputs.push(lines[i - 1])
    }
  }

  // Pair them up
  const minLen = Math.min(inputs.length, outputs.length)
  const pairedInputs = inputs.slice(0, minLen)
  const pairedOutputs = outputs.slice(0, minLen)

  const data: TranslationRow[] = []
  for (let i = 0; i < minLen; i++) {
    data.push({ Input: pairedInputs[i], Output: pairedOutputs[i] })
  }

  return data
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json()
    const mode = String(body.mode || 'translation-table')
    const requestedFilename =
      typeof body.filename === 'string' && body.filename.trim()
        ? body.filename.trim()
        : ''

    if (!body.text || typeof body.text !== 'string') {
      return NextResponse.json(
        { error: 'No text provided' },
        { status: 400 }
      )
    }

    if (mode === 'evaluation-form') {
      const forms = parseEvaluationForms(body.text)
      return NextResponse.json({
        mode: 'evaluation-form',
        forms,
        filename: requestedFilename || 'evaluation_form.txt',
      })
    }

    const data = parseTxtContent(body.text)
    const filename = requestedFilename || 'pasted_data.txt'

    return NextResponse.json({
      mode: 'translation-table',
      data,
      filename,
    })
  } catch (error) {
    console.error('Error processing text:', error)
    return NextResponse.json(
      { error: 'Invalid request format. JSON expected.' },
      { status: 400 }
    )
  }
}


