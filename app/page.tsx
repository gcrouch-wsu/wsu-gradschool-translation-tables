'use client'

import { useMemo, useState, type ChangeEvent } from 'react'
import JSZip from 'jszip'

type WorkflowMode = 'translation-table' | 'evaluation-form'

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
  sourceFile?: string
}

interface TranslationProcessResponse {
  mode: 'translation-table'
  data: TranslationRow[]
  filename: string
}

interface EvaluationProcessResponse {
  mode: 'evaluation-form'
  forms: EvaluationForm[]
  filename: string
}

type ProcessResponse = TranslationProcessResponse | EvaluationProcessResponse

function slugify(value: string): string {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
}

function downloadClientFile(content: BlobPart | BlobPart[], contentType: string, filename: string): void {
  const parts = Array.isArray(content) ? content : [content]
  const blob = new Blob(parts, { type: contentType })
  const url = window.URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.style.display = 'none'
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  window.URL.revokeObjectURL(url)
}

function buildEditableFormHtml(form: EvaluationForm): string {
  const schemaJson = JSON.stringify(form, null, 2)
    .replace(/</g, '\\u003c')
    .replace(/<\/script/gi, '<\\/script')

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${form.formTitle || 'Evaluation Form'} - Editable Export</title>
  <style>
    body { margin: 0; font-family: Arial, Helvetica, sans-serif; background: #f4f4f4; color: #333; }
    .wrap { max-width: 1100px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border: 1px solid #e0e0e0; border-radius: 10px; padding: 16px; margin-bottom: 14px; }
    textarea { width: 100%; min-height: 260px; border: 1px solid #d9d9d9; border-radius: 6px; padding: 10px; font-family: Consolas, monospace; box-sizing: border-box; }
    button { border: 0; background: #A60F2D; color: #fff; padding: 9px 12px; border-radius: 6px; cursor: pointer; font-weight: 700; margin-right: 8px; }
    .btn-gray { background: #5E6A71; }
    h1, h2, h3 { margin: 0 0 10px 0; }
    .section { border: 1px solid #e8e8e8; border-radius: 8px; padding: 10px; margin-bottom: 10px; }
    .question { border: 1px solid #efefef; border-radius: 8px; padding: 8px; margin-top: 8px; background: #fafafa; }
    .muted { color: #5E6A71; font-size: 12px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>Editable Evaluation Form Export</h1>
      <p class="muted">Edit JSON, render preview, and download JSON from this standalone file.</p>
      <button id="btnRender">Render Preview</button>
      <button class="btn-gray" id="btnDownload">Download JSON</button>
      <h3 style="margin-top:12px;">Schema JSON</h3>
      <textarea id="schema"></textarea>
    </div>
    <div class="card">
      <h2>Preview</h2>
      <div id="preview"></div>
    </div>
  </div>
  <script>
    const schemaEl = document.getElementById('schema');
    const previewEl = document.getElementById('preview');
    const initialSchema = ${schemaJson};
    schemaEl.value = JSON.stringify(initialSchema, null, 2);

    function parseSchema() {
      try {
        const parsed = JSON.parse(schemaEl.value || '{}');
        if (!parsed.sections || !Array.isArray(parsed.sections)) parsed.sections = [];
        return parsed;
      } catch (error) {
        alert('Invalid JSON: ' + (error && error.message ? error.message : 'Unknown error'));
        return null;
      }
    }

    function renderPreview() {
      const schema = parseSchema();
      if (!schema) return;
      previewEl.innerHTML = '';
      const meta = document.createElement('div');
      meta.innerHTML = '<h3>' + (schema.formTitle || 'Untitled Form') + '</h3>' +
        '<p class="muted">Display: ' + (schema.displayName || '') + ' | Key: ' + (schema.formKey || '') + '</p>';
      previewEl.appendChild(meta);

      schema.sections.forEach((section, sectionIndex) => {
        const sectionDiv = document.createElement('div');
        sectionDiv.className = 'section';
        sectionDiv.innerHTML = '<strong>Section ' + (sectionIndex + 1) + ': ' + (section.title || 'Untitled') + '</strong>' +
          (section.introduction ? '<p>' + section.introduction + '</p>' : '');

        (section.questions || []).forEach((question, questionIndex) => {
          const q = document.createElement('div');
          q.className = 'question';
          q.innerHTML = '<strong>Q' + (questionIndex + 1) + ':</strong> ' + (question.label || 'Untitled') +
            '<div class="muted">Key: ' + (question.key || '') + '</div>' +
            (question.helpText ? '<div class="muted">' + question.helpText + '</div>' : '');

          const options = Array.isArray(question.options) ? question.options.filter(Boolean) : [];
          if (options.length) {
            const ul = document.createElement('ul');
            options.forEach((opt) => {
              const li = document.createElement('li');
              li.textContent = opt;
              ul.appendChild(li);
            });
            q.appendChild(ul);
          }
          sectionDiv.appendChild(q);
        });
        previewEl.appendChild(sectionDiv);
      });
    }

    function downloadJson() {
      const schema = parseSchema();
      if (!schema) return;
      const blob = new Blob([JSON.stringify(schema, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = ((schema.formKey || 'evaluation-form') + '.json');
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    document.getElementById('btnRender').addEventListener('click', renderPreview);
    document.getElementById('btnDownload').addEventListener('click', downloadJson);
    renderPreview();
  </script>
</body>
</html>`
}

function normalizeParsedForms(
  forms: EvaluationForm[],
  sourceFile: string
): EvaluationForm[] {
  const sourceSlug = slugify(sourceFile) || 'form-source'
  return forms.map((form, index) => ({
    ...form,
    id: `${sourceSlug}-${index + 1}-${form.id || 'form'}`,
    sourceFile,
  }))
}

function buildZipHtmlFilename(form: EvaluationForm, index: number, usedNames: Set<string>): string {
  const sourceBase = slugify(form.sourceFile || 'uploaded') || 'uploaded'
  const formBase = slugify(form.formKey || form.formTitle || `form-${index + 1}`) || `form-${index + 1}`
  let candidate = `${sourceBase}_${formBase}.html`
  let suffix = 2
  while (usedNames.has(candidate)) {
    candidate = `${sourceBase}_${formBase}-${suffix}.html`
    suffix += 1
  }
  usedNames.add(candidate)
  return candidate
}

export default function TranslationTablesPage() {
  const [workflowMode, setWorkflowMode] = useState<WorkflowMode>('translation-table')
  const [pasteInput, setPasteInput] = useState('')
  const [errorMessage, setErrorMessage] = useState('')

  const [showEditSection, setShowEditSection] = useState(false)
  const [showFinalReview, setShowFinalReview] = useState(false)
  const [tableData, setTableData] = useState<TranslationRow[]>([])
  const [finalData, setFinalData] = useState<TranslationRow[]>([])
  const [selectAll, setSelectAll] = useState(true)
  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set())

  const [showFormEditor, setShowFormEditor] = useState(false)
  const [parsedForms, setParsedForms] = useState<EvaluationForm[]>([])
  const [activeFormIndex, setActiveFormIndex] = useState(0)
  const [currentFilename, setCurrentFilename] = useState('pasted_data.txt')
  const [uploadedTxtFiles, setUploadedTxtFiles] = useState<File[]>([])
  const [processingUploads, setProcessingUploads] = useState(false)

  const activeForm = useMemo(() => parsedForms[activeFormIndex] || null, [parsedForms, activeFormIndex])

  const resetViews = () => {
    setShowEditSection(false)
    setShowFinalReview(false)
    setShowFormEditor(false)
  }

  const handleWorkflowChange = (nextMode: WorkflowMode) => {
    setWorkflowMode(nextMode)
    setErrorMessage('')
    resetViews()
    setTableData([])
    setFinalData([])
    setSelectAll(true)
    setSelectedRows(new Set())
    setParsedForms([])
    setActiveFormIndex(0)
    setUploadedTxtFiles([])
    setProcessingUploads(false)
  }

  const requestProcess = async (payload: {
    text: string
    mode: WorkflowMode
    filename?: string
  }): Promise<ProcessResponse> => {
    const response = await fetch('/api/process', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    })
    if (!response.ok) {
      const error = await response.json()
      throw new Error(error.error || 'Processing failed')
    }
    return (await response.json()) as ProcessResponse
  }

  const handleProcess = async () => {
    if (!pasteInput.trim()) {
      setErrorMessage('Please paste some text to process.')
      return
    }

    setErrorMessage('')
    resetViews()

    try {
      const sourceFilename =
        workflowMode === 'evaluation-form'
          ? 'pasted-evaluation-form.txt'
          : 'pasted_data.txt'
      const result = await requestProcess({
        text: pasteInput,
        mode: workflowMode,
        filename: sourceFilename,
      })
      setCurrentFilename(result.filename || sourceFilename)

      if (workflowMode === 'evaluation-form') {
        const forms = Array.isArray((result as EvaluationProcessResponse).forms)
          ? (result as EvaluationProcessResponse).forms
          : []
        if (forms.length === 0) {
          throw new Error(
            'No evaluation forms were detected. Paste a full Evaluation Forms page capture and try again.'
          )
        }
        setParsedForms(normalizeParsedForms(forms, sourceFilename))
        setActiveFormIndex(0)
        setShowFormEditor(true)
        return
      }

      const rows = Array.isArray((result as TranslationProcessResponse).data)
        ? (result as TranslationProcessResponse).data
        : []
      setTableData(rows)
      setSelectedRows(new Set(rows.map((_, index) => index)))
      setSelectAll(true)
      setShowEditSection(true)
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : 'An error occurred'
      setErrorMessage(message)
    }
  }

  const handleUploadFilesChange = (event: ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files || []).filter((file) =>
      file.name.toLowerCase().endsWith('.txt')
    )
    setUploadedTxtFiles(files)
    if (!files.length) {
      setErrorMessage('No .txt files selected.')
      return
    }
    setErrorMessage('')
  }

  const handleProcessUploadedFiles = async () => {
    if (workflowMode !== 'evaluation-form') {
      setErrorMessage('File upload parsing is available only in Evaluation form parser mode.')
      return
    }
    if (!uploadedTxtFiles.length) {
      setErrorMessage('Select one or more .txt files first.')
      return
    }

    setErrorMessage('')
    resetViews()
    setProcessingUploads(true)

    try {
      const parsed: EvaluationForm[] = []
      const emptyFiles: string[] = []
      for (const file of uploadedTxtFiles) {
        const text = await file.text()
        const result = (await requestProcess({
          text,
          mode: 'evaluation-form',
          filename: file.name,
        })) as EvaluationProcessResponse

        const forms = Array.isArray(result.forms) ? result.forms : []
        if (!forms.length) {
          emptyFiles.push(file.name)
          continue
        }
        parsed.push(...normalizeParsedForms(forms, file.name))
      }

      if (!parsed.length) {
        throw new Error('No evaluation forms were detected in the uploaded files.')
      }

      if (emptyFiles.length) {
        setErrorMessage(
          `Parsed with skips. No forms detected in: ${emptyFiles.join(', ')}`
        )
      }
      setParsedForms(parsed)
      setActiveFormIndex(0)
      setCurrentFilename(uploadedTxtFiles.length === 1 ? uploadedTxtFiles[0].name : 'evaluation-forms.txt')
      setShowFormEditor(true)
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : 'Failed to process uploaded files.'
      setErrorMessage(message)
    } finally {
      setProcessingUploads(false)
    }
  }

  const handleSelectAll = (checked: boolean) => {
    setSelectAll(checked)
    if (checked) {
      setSelectedRows(new Set(tableData.map((_, i) => i)))
    } else {
      setSelectedRows(new Set())
    }
  }

  const handleRowToggle = (index: number, checked: boolean) => {
    const next = new Set(selectedRows)
    if (checked) next.add(index)
    else next.delete(index)
    setSelectedRows(next)
    setSelectAll(next.size === tableData.length)
  }

  const handleCellEdit = (index: number, field: 'Input' | 'Output', value: string) => {
    const next = [...tableData]
    next[index] = { ...next[index], [field]: value }
    setTableData(next)
  }

  const handlePreviewSelection = () => {
    const selected = tableData.filter((_, i) => selectedRows.has(i))
    if (selected.length === 0) {
      alert('No rows selected. Please select at least one row.')
      return
    }
    setFinalData(selected)
    setShowEditSection(false)
    setShowFinalReview(true)
  }

  const handleDownloadTranslation = async (format: 'xlsx' | 'txt') => {
    if (finalData.length === 0) {
      alert('No data to download.')
      return
    }
    const response = await fetch('/api/download', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: finalData, format, filename: currentFilename }),
    })
    if (!response.ok) {
      alert('Download failed')
      return
    }
    const blob = await response.blob()
    const url = window.URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.style.display = 'none'
    a.href = url
    const ext = format === 'xlsx' ? 'xlsx' : 'txt'
    const baseName = currentFilename.replace(/\.[^/.]+$/, '')
    a.download = `${baseName}_processed.${ext}`
    document.body.appendChild(a)
    a.click()
    window.URL.revokeObjectURL(url)
    document.body.removeChild(a)
  }

  const updateActiveForm = (updater: (form: EvaluationForm) => EvaluationForm) => {
    setParsedForms((prev) => prev.map((form, index) => (index === activeFormIndex ? updater(form) : form)))
  }

  const handleFormMetaChange = (field: 'formTitle' | 'displayName' | 'formKey', value: string) => {
    updateActiveForm((form) => ({ ...form, [field]: value }))
  }

  const handleSectionChange = (
    sectionIndex: number,
    field: 'title' | 'introduction' | 'conditional',
    value: string | boolean
  ) => {
    updateActiveForm((form) => ({
      ...form,
      sections: form.sections.map((section, idx) => (idx === sectionIndex ? { ...section, [field]: value } : section)),
    }))
  }

  const handleQuestionChange = (
    sectionIndex: number,
    questionIndex: number,
    field: 'label' | 'key' | 'helpText' | 'options' | 'conditional',
    value: string | boolean
  ) => {
    updateActiveForm((form) => ({
      ...form,
      sections: form.sections.map((section, sIdx) => {
        if (sIdx !== sectionIndex) return section
        return {
          ...section,
          questions: section.questions.map((question, qIdx) => {
            if (qIdx !== questionIndex) return question
            if (field === 'options' && typeof value === 'string') {
              return { ...question, options: value.split('\n').map((v) => v.trim()).filter(Boolean) }
            }
            return { ...question, [field]: value }
          }),
        }
      }),
    }))
  }

  const handleDownloadFormJson = () => {
    if (!activeForm) return
    const base = slugify(activeForm.formKey || activeForm.displayName || activeForm.formTitle) || 'evaluation-form'
    downloadClientFile(JSON.stringify(activeForm, null, 2), 'application/json;charset=utf-8', `${base}.json`)
  }

  const handleDownloadEditableHtml = () => {
    if (!activeForm) return
    const base = slugify(activeForm.formKey || activeForm.displayName || activeForm.formTitle) || 'evaluation-form'
    downloadClientFile(buildEditableFormHtml(activeForm), 'text/html;charset=utf-8', `${base}_editable.html`)
  }

  const handleDownloadAllHtmlZip = async () => {
    if (!parsedForms.length) {
      setErrorMessage('No parsed forms are available for ZIP export.')
      return
    }
    try {
      const zip = new JSZip()
      const usedNames = new Set<string>()
      parsedForms.forEach((form, index) => {
        const entryName = buildZipHtmlFilename(form, index, usedNames)
        zip.file(entryName, buildEditableFormHtml(form))
      })
      const zipBlob = await zip.generateAsync({ type: 'blob' })
      const base = slugify(currentFilename.replace(/\.[^/.]+$/, '')) || 'evaluation-forms'
      downloadClientFile(zipBlob, 'application/zip', `${base}_editable_html.zip`)
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : 'Failed to build ZIP.'
      setErrorMessage(message)
    }
  }

  const processButtonLabel =
    workflowMode === 'translation-table'
      ? 'Process Translation Data'
      : 'Parse Pasted Evaluation Form'
  const pastePlaceholder =
    workflowMode === 'translation-table'
      ? 'Paste Outcomes translation table text here...'
      : 'Paste Outcomes Evaluation Form page text here...'

  return (
    <div className="min-h-screen bg-wsu-bg-light flex flex-col">
      <header className="bg-wsu-crimson text-white py-6 px-8 shadow-md">
        <h1 className="text-3xl font-bold uppercase tracking-wide">Washington State University</h1>
        <h2 className="text-xl mt-1 opacity-90">Translation + Form Extraction Tools</h2>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 w-full">
        <section className="bg-white rounded-lg shadow-md p-8 mb-8">
          <h3 className="text-2xl font-semibold text-wsu-crimson mb-4">Import Data</h3>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
            <label className="border border-wsu-border-medium rounded-lg p-3 cursor-pointer hover:border-wsu-crimson">
              <input
                type="radio"
                name="workflow-mode"
                value="translation-table"
                checked={workflowMode === 'translation-table'}
                onChange={() => handleWorkflowChange('translation-table')}
                className="mr-2"
              />
              <span className="font-semibold text-wsu-text-dark">Translation table cleanup</span>
              <p className="text-sm text-wsu-text-muted mt-1">Edit/select page-pasted Input/Output rows, then download XLSX/TXT.</p>
            </label>
            <label className="border border-wsu-border-medium rounded-lg p-3 cursor-pointer hover:border-wsu-crimson">
              <input
                type="radio"
                name="workflow-mode"
                value="evaluation-form"
                checked={workflowMode === 'evaluation-form'}
                onChange={() => handleWorkflowChange('evaluation-form')}
                className="mr-2"
              />
              <span className="font-semibold text-wsu-text-dark">Evaluation form parser</span>
              <p className="text-sm text-wsu-text-muted mt-1">Parse pasted or uploaded Evaluation Form text into editable sections/questions, then download JSON/HTML or multi-file HTML ZIP.</p>
            </label>
          </div>

          {workflowMode === 'evaluation-form' && (
            <div className="mb-4 p-3 border border-gray-200 rounded-lg bg-gray-50">
              <label className="block text-sm font-semibold text-wsu-text-muted mb-2">
                Upload .txt files (single or multiple)
              </label>
              <input
                type="file"
                accept=".txt,text/plain"
                multiple
                onChange={handleUploadFilesChange}
                className="block w-full text-sm text-wsu-text-body"
              />
              <div className="mt-2 flex flex-wrap items-center gap-3">
                <button
                  type="button"
                  onClick={handleProcessUploadedFiles}
                  disabled={processingUploads}
                  className="bg-wsu-gray-light text-white px-4 py-2 rounded font-semibold hover:bg-wsu-gray transition-colors disabled:opacity-60 disabled:cursor-not-allowed"
                >
                  {processingUploads ? 'Processing uploads...' : 'Process Uploaded Files'}
                </button>
                <span className="text-xs text-wsu-text-muted">
                  {uploadedTxtFiles.length
                    ? `${uploadedTxtFiles.length} file(s) selected`
                    : 'No files selected'}
                </span>
              </div>
              {uploadedTxtFiles.length > 0 && (
                <ul className="mt-2 text-xs text-wsu-text-muted list-disc list-inside">
                  {uploadedTxtFiles.slice(0, 8).map((file) => (
                    <li key={`${file.name}-${file.lastModified}`}>{file.name}</li>
                  ))}
                  {uploadedTxtFiles.length > 8 && (
                    <li>...and {uploadedTxtFiles.length - 8} more</li>
                  )}
                </ul>
              )}
            </div>
          )}

          <textarea
            id="pasteInput"
            value={pasteInput}
            onChange={(e) => setPasteInput(e.target.value)}
            placeholder={pastePlaceholder}
            className="w-full h-44 p-3 border border-gray-300 rounded-lg font-mono resize-y focus:border-wsu-crimson focus:outline-none"
          />
          <button
            onClick={handleProcess}
            className="w-full mt-3 bg-wsu-crimson text-white px-6 py-3 rounded font-semibold hover:bg-wsu-crimson-dark transition-colors"
          >
            {processButtonLabel}
          </button>
          {errorMessage && <div className="text-red-600 font-semibold mt-4">{errorMessage}</div>}
        </section>

        {workflowMode === 'translation-table' && showEditSection && (
          <section className="bg-white rounded-lg shadow-md p-8 mb-8">
            <h3 className="text-2xl font-semibold text-wsu-crimson mb-1">Edit and Select Data</h3>
            <p className="text-wsu-text-muted mb-4">Uncheck rows to exclude them. Click cells to edit content.</p>

            <div className="max-h-[500px] overflow-y-auto border border-gray-300 rounded mb-6">
              <table className="w-full border-collapse text-sm">
                <thead className="bg-wsu-gray-light text-white sticky top-0">
                  <tr>
                    <th className="w-10 text-center p-3">
                      <input
                        type="checkbox"
                        checked={selectAll}
                        onChange={(e) => handleSelectAll(e.target.checked)}
                        className="cursor-pointer"
                      />
                    </th>
                    <th className="p-3 text-left font-semibold">Input</th>
                    <th className="p-3 text-left font-semibold">Output</th>
                  </tr>
                </thead>
                <tbody>
                  {tableData.length === 0 ? (
                    <tr>
                      <td colSpan={3} className="text-center p-4">No valid data found.</td>
                    </tr>
                  ) : (
                    tableData.map((row, index) => (
                      <tr key={index} className={`${index % 2 === 0 ? 'bg-white' : 'bg-gray-50'} hover:bg-gray-100`}>
                        <td className="text-center p-3">
                          <input
                            type="checkbox"
                            checked={selectedRows.has(index)}
                            onChange={(e) => handleRowToggle(index, e.target.checked)}
                            className="cursor-pointer"
                          />
                        </td>
                        <td
                          contentEditable
                          suppressContentEditableWarning
                          onBlur={(e) => handleCellEdit(index, 'Input', e.currentTarget.textContent || '')}
                          className="p-3 border border-transparent focus:border-wsu-crimson focus:outline-none focus:bg-red-50"
                        >
                          {row.Input}
                        </td>
                        <td
                          contentEditable
                          suppressContentEditableWarning
                          onBlur={(e) => handleCellEdit(index, 'Output', e.currentTarget.textContent || '')}
                          className="p-3 border border-transparent focus:border-wsu-crimson focus:outline-none focus:bg-red-50"
                        >
                          {row.Output}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            <div className="flex justify-end">
              <button
                onClick={handlePreviewSelection}
                className="bg-wsu-crimson text-white px-6 py-3 rounded font-semibold hover:bg-wsu-crimson-dark transition-colors"
              >
                Preview Selection -&gt;
              </button>
            </div>
          </section>
        )}

        {workflowMode === 'translation-table' && showFinalReview && (
          <section className="bg-white rounded-lg shadow-md p-8 mb-8">
            <h3 className="text-2xl font-semibold text-wsu-crimson mb-4">Final Review</h3>
            <div className="max-h-[500px] overflow-y-auto border border-gray-300 rounded mb-6">
              <table className="w-full border-collapse text-sm">
                <thead className="bg-wsu-gray-light text-white sticky top-0">
                  <tr>
                    <th className="p-3 text-left font-semibold">Input</th>
                    <th className="p-3 text-left font-semibold">Output</th>
                  </tr>
                </thead>
                <tbody>
                  {finalData.map((row, index) => (
                    <tr key={index} className={`${index % 2 === 0 ? 'bg-white' : 'bg-gray-50'}`}>
                      <td className="p-3">{row.Input}</td>
                      <td className="p-3">{row.Output}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="flex justify-end gap-4">
              <button
                onClick={() => {
                  setShowFinalReview(false)
                  setShowEditSection(true)
                }}
                className="bg-wsu-gray-light text-white px-6 py-3 rounded font-semibold hover:bg-wsu-gray transition-colors"
              >
                &lt;- Back to Edit
              </button>
              <button
                onClick={() => handleDownloadTranslation('xlsx')}
                className="bg-wsu-crimson text-white px-6 py-3 rounded font-semibold hover:bg-wsu-crimson-dark transition-colors"
              >
                Download Excel
              </button>
              <button
                onClick={() => handleDownloadTranslation('txt')}
                className="bg-wsu-crimson text-white px-6 py-3 rounded font-semibold hover:bg-wsu-crimson-dark transition-colors"
              >
                Download Text
              </button>
            </div>
          </section>
        )}

        {workflowMode === 'evaluation-form' && showFormEditor && activeForm && (
          <section className="bg-white rounded-lg shadow-md p-8 mb-8">
            <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
              <h3 className="text-2xl font-semibold text-wsu-crimson">Parsed Evaluation Form</h3>
              <div className="flex gap-2">
                <button
                  onClick={handleDownloadFormJson}
                  className="bg-wsu-gray-light text-white px-4 py-2 rounded font-semibold hover:bg-wsu-gray transition-colors"
                >
                  Download JSON
                </button>
                <button
                  onClick={handleDownloadEditableHtml}
                  className="bg-wsu-crimson text-white px-4 py-2 rounded font-semibold hover:bg-wsu-crimson-dark transition-colors"
                >
                  Download Editable HTML
                </button>
                {parsedForms.length > 1 && (
                  <button
                    onClick={handleDownloadAllHtmlZip}
                    className="bg-wsu-crimson text-white px-4 py-2 rounded font-semibold hover:bg-wsu-crimson-dark transition-colors"
                  >
                    Download All HTML as ZIP
                  </button>
                )}
              </div>
            </div>

            {parsedForms.length > 1 && (
              <div className="mb-4">
                <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Detected Forms</label>
                <select
                  value={String(activeFormIndex)}
                  onChange={(e) => setActiveFormIndex(Number(e.target.value))}
                  className="w-full max-w-xl border border-gray-300 rounded px-3 py-2"
                >
                  {parsedForms.map((form, index) => (
                    <option key={form.id || String(index)} value={String(index)}>
                      {index + 1}. {form.formTitle}
                    </option>
                  ))}
                </select>
              </div>
            )}
            <p className="text-xs text-wsu-text-muted mb-3">
              Source file: {activeForm.sourceFile || 'pasted-evaluation-form.txt'}
            </p>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-6">
              <div>
                <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Form Title</label>
                <input
                  value={activeForm.formTitle}
                  onChange={(e) => handleFormMetaChange('formTitle', e.target.value)}
                  className="w-full border border-gray-300 rounded px-3 py-2"
                />
              </div>
              <div>
                <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Display Name</label>
                <input
                  value={activeForm.displayName}
                  onChange={(e) => handleFormMetaChange('displayName', e.target.value)}
                  className="w-full border border-gray-300 rounded px-3 py-2"
                />
              </div>
              <div>
                <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Form Key</label>
                <input
                  value={activeForm.formKey}
                  onChange={(e) => handleFormMetaChange('formKey', e.target.value)}
                  className="w-full border border-gray-300 rounded px-3 py-2 font-mono"
                />
              </div>
            </div>

            <div className="space-y-4">
              {activeForm.sections.map((section, sectionIndex) => (
                <details key={section.id || String(sectionIndex)} className="border border-gray-200 rounded-lg p-4" open>
                  <summary className="cursor-pointer font-semibold text-wsu-text-dark mb-3">
                    Section {sectionIndex + 1}: {section.title || 'Untitled'}
                  </summary>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                    <div>
                      <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Section Title</label>
                      <input
                        value={section.title}
                        onChange={(e) => handleSectionChange(sectionIndex, 'title', e.target.value)}
                        className="w-full border border-gray-300 rounded px-3 py-2"
                      />
                    </div>
                    <label className="inline-flex items-center mt-6 text-sm text-wsu-text-body">
                      <input
                        type="checkbox"
                        checked={Boolean(section.conditional)}
                        onChange={(e) => handleSectionChange(sectionIndex, 'conditional', e.target.checked)}
                        className="mr-2"
                      />
                      Conditional section
                    </label>
                  </div>
                  <div className="mb-4">
                    <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Introduction</label>
                    <textarea
                      value={section.introduction}
                      onChange={(e) => handleSectionChange(sectionIndex, 'introduction', e.target.value)}
                      className="w-full border border-gray-300 rounded px-3 py-2 min-h-[80px]"
                    />
                  </div>

                  {section.questions.map((question, questionIndex) => (
                    <div key={question.id || String(questionIndex)} className="border border-gray-200 rounded p-3 bg-gray-50 mb-3">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                        <div>
                          <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Question Label</label>
                          <input
                            value={question.label}
                            onChange={(e) => handleQuestionChange(sectionIndex, questionIndex, 'label', e.target.value)}
                            className="w-full border border-gray-300 rounded px-3 py-2"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Question Key</label>
                          <input
                            value={question.key}
                            onChange={(e) => handleQuestionChange(sectionIndex, questionIndex, 'key', e.target.value)}
                            className="w-full border border-gray-300 rounded px-3 py-2 font-mono"
                          />
                        </div>
                      </div>
                      <div className="mb-3">
                        <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Help Text</label>
                        <textarea
                          value={question.helpText}
                          onChange={(e) => handleQuestionChange(sectionIndex, questionIndex, 'helpText', e.target.value)}
                          className="w-full border border-gray-300 rounded px-3 py-2 min-h-[70px]"
                        />
                      </div>
                      <div className="mb-2">
                        <label className="block text-sm font-semibold text-wsu-text-muted mb-1">Options (one per line)</label>
                        <textarea
                          value={question.options.join('\n')}
                          onChange={(e) => handleQuestionChange(sectionIndex, questionIndex, 'options', e.target.value)}
                          className="w-full border border-gray-300 rounded px-3 py-2 min-h-[90px]"
                        />
                      </div>
                      <label className="inline-flex items-center text-sm text-wsu-text-body">
                        <input
                          type="checkbox"
                          checked={Boolean(question.conditional)}
                          onChange={(e) => handleQuestionChange(sectionIndex, questionIndex, 'conditional', e.target.checked)}
                          className="mr-2"
                        />
                        Conditional question
                      </label>
                    </div>
                  ))}
                </details>
              ))}
            </div>
          </section>
        )}
      </main>

      <footer className="border-t border-wsu-border-light bg-white mt-auto">
        <div className="max-w-5xl mx-auto px-4 py-6">
          <p className="text-sm text-wsu-text-muted text-center">Graduate School | Washington State University</p>
        </div>
      </footer>
    </div>
  )
}
