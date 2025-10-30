import React, { useMemo, useRef, useState } from 'react'
import JSZip from 'jszip'
import { readWorkbookFromFile, firstSheetToRows, convertRows, toWorkbookFromJson, workbookToBlob, previewRows, countValidRows } from './lib/convert.js'

function TablePreview({ rows, title }) {
    if (!rows?.length) return <div className="muted">No preview</div>
    const headers = rows[0]?.map((_, idx) => `Col ${idx + 1}`) || []
    return (
        <div>
            {title ? <div className="muted" style={{ marginBottom: 6 }}>{title}</div> : null}
            <table aria-label={title || 'Preview table'}>
                <caption className="sr-only">{title || 'Preview table'}</caption>
                <thead>
                    <tr>
                        {headers.map((h, i) => <th key={i} scope="col">{h}</th>)}
                    </tr>
                </thead>
                <tbody>
                    {rows.map((r, i) => (
                        <tr key={i}>
                            {r.map((c, j) => <td key={j}>{String(c)}</td>)}
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    )
}

function DownloadLink({ blob, filename, children }) {
    const url = useMemo(() => blob ? URL.createObjectURL(blob) : null, [blob])
    return (
        <a href={url || '#'} download={filename} onClick={e => { if (!url) e.preventDefault() }}>
            {children}
        </a>
    )
}

export default function App() {
    const [items, setItems] = useState([])
    const [busy, setBusy] = useState(false)
    const [drag, setDrag] = useState(false)
    const inputRef = useRef(null)
    const [expanded, setExpanded] = useState({})
    const [rtl, setRtl] = useState(false)
    const [tab, setTab] = useState('upload') // 'upload' | 'files' | 'results'
    const [live, setLive] = useState('')

    const convertedCount = items.filter(i => i.status === 'done').length
    const totalCount = items.length
    const progress = totalCount ? Math.round((convertedCount / totalCount) * 100) : 0
    // Improved clarity: map internal statuses to friendly labels for the UI
    const statusLabel = (s) => s === 'ready' ? 'Ready' : s === 'converting' ? 'Converting…' : s === 'done' ? 'Converted' : s
    const formatBytes = (bytes) => {
        if (typeof bytes !== 'number') return ''
        if (bytes < 1024) return `${bytes} B`
        const units = ['KB','MB','GB']
        let i = -1; do { bytes = bytes / 1024; i++ } while (bytes >= 1024 && i < units.length - 1)
        return `${bytes.toFixed(1)} ${units[i]}`
    }

    const onPick = async (files) => {
        const fileArr = Array.from(files || [])
        const appended = await Promise.all(fileArr.map(async (f) => {
            const wb = await readWorkbookFromFile(f)
            const rows = firstSheetToRows(wb)
            const recognized = countValidRows(rows) // Improved pre-conversion feedback
            return {
                id: crypto.randomUUID(),
                file: f,
                name: f.name.replace(/\.[^.]+$/, ''),
                rows,
                preview: previewRows(rows, 5),
                status: 'ready',
                resultBlob: null,
                resultName: null,
                validCount: recognized,
                totalCount: rows.length
            }
        }))
        setItems(prev => [...prev, ...appended])
        if (inputRef.current) inputRef.current.value = ''
    }

    const handleFiles = (e) => onPick(e.target.files)

    const onDrop = async (e) => {
        e.preventDefault()
        e.stopPropagation()
        setDrag(false)
        if (e.dataTransfer?.files?.length) {
            await onPick(e.dataTransfer.files)
        }
    }

    const onDragOver = (e) => { e.preventDefault(); e.stopPropagation(); setDrag(true) }
    const onDragLeave = (e) => { e.preventDefault(); e.stopPropagation(); setDrag(false) }
    const onDropzoneKey = (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault()
            inputRef.current?.click()
        }
    }

    const convertOne = async (id) => {
        setItems(prev => prev.map(x => x.id === id ? { ...x, status: 'converting' } : x))
        setLive('Converting ' + (items.find(x => x.id === id)?.name || 'file') + '...')
        try {
            const item = items.find(x => x.id === id)
            const converted = convertRows(item.rows)
            const wb = toWorkbookFromJson(converted)
            const { blob, filename } = workbookToBlob(wb, item.name)
            setItems(prev => prev.map(x => x.id === id ? { ...x, status: 'done', resultBlob: blob, resultName: filename } : x))
            setLive('Converted ' + item.name)
        } catch (e) {
            setItems(prev => prev.map(x => x.id === id ? { ...x, status: 'error' } : x))
            setLive('Failed to convert file')
        }
    }

    const convertAll = async () => {
        setBusy(true)
        try {
            for (const it of items) {
                if (it.status === 'done') continue
                // eslint-disable-next-line no-await-in-loop
                await convertOne(it.id)
            }
        } finally {
            setBusy(false)
        }
    }

    const downloadAllZip = async () => {
        const zip = new JSZip()
        let any = false
        items.forEach(it => {
            if (it.resultBlob && it.resultName) {
                zip.file(it.resultName, it.resultBlob)
                any = true
            }
        })
        if (!any) return
        const blob = await zip.generateAsync({ type: 'blob' })
        const url = URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = url
        a.download = 'converted_files.zip'
        document.body.appendChild(a)
        a.click()
        a.remove()
        URL.revokeObjectURL(url)
    }

    const removeItem = (id) => setItems(prev => prev.filter(x => x.id !== id))
    const clearAll = () => { setItems([]); setExpanded({}) }

    return (
        <div className="container">
            <div className="actions" dir={rtl ? 'rtl' : 'ltr'} role="banner">
                <div className="row">
                    <input aria-label="Upload Excel files" ref={inputRef} type="file" accept=".xlsx,.xls" multiple onChange={handleFiles} />
                    <button className="ghost" onClick={() => inputRef.current?.click()}>📁 Choose Files</button>
                    <div className="spacer" />
                    <button onClick={convertAll} disabled={!items.length || busy} aria-busy={busy} aria-disabled={!items.length || busy}>
                        {busy ? <span className="spin" aria-hidden /> : '⚙️'} Convert All
                    </button>
                    <button onClick={downloadAllZip} disabled={!items.some(i => i.resultBlob)}>
                        ⬇️ Download All (ZIP)
                    </button>
                    <button className="danger" onClick={clearAll} disabled={!items.length} title="Remove all uploaded files">
                        🗑️ Clear All
                    </button>
                    <button className="secondary" onClick={() => setRtl(v => !v)} title="Toggle direction">
                        {rtl ? '🔁 LTR' : '🔁 RTL'}
                    </button>
                </div>
                <div className="progress" aria-hidden={totalCount === 0} title={`Progress: ${progress}%`}>
                    <span style={{ width: `${progress}%` }} />
                </div>
            </div>
            <div className="header" dir={rtl ? 'rtl' : 'ltr'}>
                <h1>Excel Converter</h1>
                <div className="sub">Upload .xlsx/.xls, preview a few rows, convert individually or in bulk, and download results.</div>
                {/* Step indicators to make the flow obvious */}
                <div className="steps">
                    <span className="step-pill">1. Upload</span>
                    <span className="step-pill">2. Convert</span>
                    <span className="step-pill">3. Download</span>
                </div>
                <div className="stats">
                    <div className="stat">Files: <strong>{totalCount}</strong></div>
                    <div className="stat">Converted: <strong>{convertedCount}</strong></div>
                </div>
                <nav className="tabs" aria-label="Sections">
                    <button className="tab" aria-current={tab === 'upload' ? 'page' : undefined} onClick={() => setTab('upload')}>Upload</button>
                    <button className="tab" aria-current={tab === 'files' ? 'page' : undefined} onClick={() => setTab('files')}>
                        Files <span className="badge">{totalCount}</span>
                    </button>
                    <button className="tab" aria-current={tab === 'results' ? 'page' : undefined} onClick={() => setTab('results')}>
                        Results <span className="badge">{items.filter(i => i.resultBlob).length}</span>
                    </button>
                </nav>
            </div>
            <main id="main" role="main" className="card" dir={rtl ? 'rtl' : 'ltr'}>
                {tab === 'upload' && (
                    <div>
                        <div className={`drop ${drag ? 'drag' : ''}`} onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop} tabIndex={0} role="button" aria-label="Drop files here or press Enter to choose files" onKeyDown={onDropzoneKey}>
                            <div>
                                <div style={{ fontWeight: 700, marginBottom: 6 }}>☁️ Drag and drop files here</div>
                                <div className="muted" style={{ marginBottom: 8 }}>or click Choose Files above to select .xlsx/.xls</div>
                                {/* Improved clarity of accepted formats and expectation setting */}
                                <div className="muted" style={{ fontSize: 12 }}>Accepted: .xlsx, .xls · We will preview the first 5 rows</div>
                                <div className="muted">{items.length ? `${items.length} file(s) added` : 'No files added yet'}</div>
                            </div>
                        </div>
                        {!items.length && (
                            <div className="empty" style={{ marginTop: 12 }}>
                                <div className="icon">📄</div>
                                <div>No files uploaded yet.</div>
                                <div className="muted">Drag and drop above or choose files to begin.</div>
                            </div>
                        )}
                    </div>
                )}
                {tab === 'files' && (
                    <ul className="list">
                        {items.map(item => (
                            <li className="item" key={item.id}>
                                <div className="row" style={{ marginBottom: 8 }}>
                                    <div style={{ minWidth: 0 }}>
                                        <strong style={{ display: 'block' }}>{item.name}</strong>
                                        {/* Improved data density: show recognized rows vs total */}
                                        <span className="muted" style={{ fontSize: 12 }}>{formatBytes(item.file?.size)} • Sheet 1 • {item.validCount ?? 0}/{item.totalCount ?? 0} rows recognized</span>
                                    </div>
                                    <div className="spacer" />
                                    <span className={`chip ${item.status === 'done' ? 'ok' : item.status === 'error' ? 'err' : ''}`} aria-live="polite">{statusLabel(item.status)}</span>
                                    <div className="btn-row">
                                        <button onClick={() => convertOne(item.id)} disabled={item.status === 'converting'} aria-busy={item.status === 'converting'}>
                                            {item.status === 'converting' ? <span className="spin" aria-hidden /> : '⚙️'}
                                            {item.status === 'done' ? 'Reconvert' : 'Convert now'}
                                        </button>
                                        {item.resultBlob && (
                                            <a className="ghost" role="button" href={URL.createObjectURL(item.resultBlob)} download={item.resultName}>
                                                ⬇️ Download
                                            </a>
                                        )}
                                        <button className="secondary" onClick={() => setExpanded(prev => ({ ...prev, [item.id]: !prev[item.id] }))}>
                                            {expanded[item.id] ? '▾ Hide Preview' : '▸ Show Preview'}
                                        </button>
                                        <button className="danger" onClick={() => removeItem(item.id)} title="Remove file">
                                            🗑️ Remove
                                        </button>
                                    </div>
                                </div>
                                {expanded[item.id] !== false && (
                                    <div className="grid">
                                        <div className="preview-col">
                                            {/* Improve separation between original and converted previews */}
                                            <div className="section-title"><strong>Original</strong></div>
                                            <div className="table-wrap">
                                                <TablePreview rows={item.preview} title="Uploaded preview (first 5 rows)" />
                                            </div>
                                        </div>
                                        {item.resultBlob && (
                                            <div className="preview-col">
                                                <div className="section-title"><strong>Converted</strong></div>
                                                <div className="table-wrap">
                                                    <TablePreview rows={convertRows(item.rows).slice(0, 5).map(r => Object.values(r))} title="Converted preview" />
                                                </div>
                                            </div>
                                        )}
                                        {item.status === 'converting' && (
                                            <div className="preview-col" style={{ gridColumn: '1 / -1' }}>
                                                {/* Loading placeholder while converting */}
                                                <div className="skeleton" aria-hidden />
                                            </div>
                                        )}
                                        {item.status === 'error' && (
                                            <div className="muted" style={{ gridColumn: '1 / -1', color: '#ffb3ae' }}>
                                                Conversion failed. Make sure the first column follows the pattern “name price unit” (e.g. «محصول 123 €»)
                                            </div>
                                        )}
                                    </div>
                                )}
                            </li>
                        ))}
                        {!items.length && (
                            <li className="empty">
                                <div className="icon">📂</div>
                                <div>No files listed.</div>
                                <div className="muted">Upload files in the Upload tab.</div>
                            </li>
                        )}
                    </ul>
                )}
                {tab === 'results' && (
                    <div>
                        {/* Improved discoverability: duplicate bulk download in Results */}
                        <div className="row" style={{ marginBottom: 12 }}>
                            <button onClick={downloadAllZip} disabled={!items.some(i => i.resultBlob)}>
                                ⬇️ Download All (ZIP)
                            </button>
                        </div>
                        <ul className="list">
                        {items.filter(i => i.resultBlob).map(item => (
                            <li className="item" key={item.id}>
                                <div className="row" style={{ marginBottom: 8 }}>
                                    <strong>{item.resultName}</strong>
                                    <div className="spacer" />
                                    <a className="ghost" role="button" href={URL.createObjectURL(item.resultBlob)} download={item.resultName}>
                                        ⬇️ Download
                                    </a>
                                </div>
                            </li>
                        ))}
                        {!items.filter(i => i.resultBlob).length && (
                            <li className="empty">
                                <div className="icon">📦</div>
                                <div>No results yet.</div>
                                <div className="muted">Convert files in the Files tab.</div>
                            </li>
                        )}
                        </ul>
                    </div>
                )}
                <div className="sr-only" aria-live="polite">{live}</div>
            </main>
            <div className="footer">Built with ❤️ – drag-and-drop, bulk convert, and ZIP export.</div>
        </div>
    )
}


