import { useState, useCallback, useRef } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

type FileStatus = 'idle' | 'loading' | 'success' | 'error'

interface FileState {
  status: FileStatus
  name: string | null
  rowCount: number | null
  error: string | null
}

const initialFileState: FileState = {
  status: 'idle',
  name: null,
  rowCount: null,
  error: null,
}

function parseExcel(file: File): Promise<{ rowCount: number }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = e.target?.result
        const wb = XLSX.read(data, { type: 'array' })
        const ws = wb.Sheets[wb.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]
        resolve({ rowCount: rows.length })
      } catch {
        reject(new Error('ファイルの読み込みに失敗しました'))
      }
    }
    reader.onerror = () => reject(new Error('ファイルの読み込みに失敗しました'))
    reader.readAsArrayBuffer(file)
  })
}

interface DropZoneProps {
  label: string
  fileKey: 'ep' | 'ev'
  accept: string
  state: FileState
  onFile: (file: File) => void
}

function DropZone({ label, fileKey, accept, state, onFile }: DropZoneProps) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault()
      setDragging(false)
      const file = e.dataTransfer.files[0]
      if (file) onFile(file)
    },
    [onFile],
  )

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file) onFile(file)
    e.target.value = ''
  }

  const statusIcon = () => {
    if (state.status === 'loading') return <span className="spinner" />
    if (state.status === 'success') return <span className="icon-success">✓</span>
    if (state.status === 'error') return <span className="icon-error">✕</span>
    return (
      <svg className="upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
        <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
      </svg>
    )
  }

  const tagColor = fileKey === 'ep' ? 'tag-ep' : 'tag-ev'

  return (
    <div
      className={`dropzone ${dragging ? 'dragging' : ''} ${state.status}`}
      onDragOver={(e) => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => e.key === 'Enter' && inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        className="hidden-input"
        onChange={handleChange}
      />

      <div className="dropzone-inner">
        <div className={`file-tag ${tagColor}`}>{label}</div>

        <div className="dropzone-icon">{statusIcon()}</div>

        {state.status === 'idle' && (
          <>
            <p className="dropzone-main">ファイルをドロップ</p>
            <p className="dropzone-sub">または クリックして選択</p>
            <p className="dropzone-hint">.xlsx / .xls</p>
          </>
        )}

        {state.status === 'loading' && (
          <p className="dropzone-main">読み込み中...</p>
        )}

        {state.status === 'success' && (
          <>
            <p className="dropzone-main success-text">{state.name}</p>
            <p className="dropzone-sub">{state.rowCount?.toLocaleString()} 行を検出</p>
            <p className="dropzone-hint replace-hint">クリックで差し替え</p>
          </>
        )}

        {state.status === 'error' && (
          <>
            <p className="dropzone-main error-text">読み込み失敗</p>
            <p className="dropzone-sub">{state.error}</p>
            <p className="dropzone-hint replace-hint">クリックで再試行</p>
          </>
        )}
      </div>
    </div>
  )
}

export default function App() {
  const [epState, setEpState] = useState<FileState>(initialFileState)
  const [evState, setEvState] = useState<FileState>(initialFileState)

  const handleFile = useCallback(
    (setter: React.Dispatch<React.SetStateAction<FileState>>) =>
      async (file: File) => {
        setter({ status: 'loading', name: file.name, rowCount: null, error: null })
        try {
          const { rowCount } = await parseExcel(file)
          setter({ status: 'success', name: file.name, rowCount, error: null })
        } catch (e) {
          setter({ status: 'error', name: file.name, rowCount: null, error: (e as Error).message })
        }
      },
    [],
  )

  const bothReady = epState.status === 'success' && evState.status === 'success'
  const anyLoaded = epState.status !== 'idle' || evState.status !== 'idle'

  const handleReset = () => {
    setEpState(initialFileState)
    setEvState(initialFileState)
  }

  return (
    <div className="app">
      <header className="header">
        <div className="header-inner">
          <div className="header-logo">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path strokeLinecap="round" strokeLinejoin="round" d="M7.5 14.25v2.25m3-4.5v4.5m3-6.75v6.75m3-9v9M6 20.25h12A2.25 2.25 0 0020.25 18V6A2.25 2.25 0 0018 3.75H6A2.25 2.25 0 003.75 6v12A2.25 2.25 0 006 20.25z" />
            </svg>
          </div>
          <div>
            <h1 className="header-title">データ比較ツール</h1>
            <p className="header-subtitle">EP・EV データを取り込んで比較分析</p>
          </div>
        </div>
      </header>

      <main className="main">
        <div className="step-label">
          <span className="step-badge">Step 1</span>
          比較するファイルをインポート
        </div>

        <div className="dropzone-grid">
          <DropZone
            label="EP データ"
            fileKey="ep"
            accept=".xlsx,.xls"
            state={epState}
            onFile={handleFile(setEpState)}
          />
          <div className="divider-col">
            <div className="divider-line" />
            <span className="divider-vs">VS</span>
            <div className="divider-line" />
          </div>
          <DropZone
            label="EV データ"
            fileKey="ev"
            accept=".xlsx,.xls"
            state={evState}
            onFile={handleFile(setEvState)}
          />
        </div>

        {anyLoaded && (
          <div className="summary-bar">
            <div className="summary-items">
              <SummaryItem label="EP" state={epState} color="#3b82f6" />
              <SummaryItem label="EV" state={evState} color="#10b981" />
            </div>
            <button className="reset-btn" onClick={handleReset}>
              リセット
            </button>
          </div>
        )}

        <button
          className={`analyze-btn ${bothReady ? 'active' : ''}`}
          disabled={!bothReady}
        >
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path strokeLinecap="round" strokeLinejoin="round" d="M3.75 3v11.25A2.25 2.25 0 006 16.5h2.25M3.75 3h-1.5m1.5 0h16.5m0 0h1.5m-1.5 0v11.25A2.25 2.25 0 0118 16.5h-2.25m-7.5 0h7.5m-7.5 0l-1 3m8.5-3l1 3m0 0l.5 1.5m-.5-1.5h-9.5m0 0l-.5 1.5" />
          </svg>
          比較分析を開始
        </button>

        {!bothReady && (
          <p className="analyze-hint">
            {epState.status === 'idle' && evState.status === 'idle'
              ? '両方のファイルをインポートしてください'
              : epState.status !== 'success'
              ? 'EP データをインポートしてください'
              : 'EV データをインポートしてください'}
          </p>
        )}
      </main>
    </div>
  )
}

function SummaryItem({ label, state, color }: { label: string; state: FileState; color: string }) {
  if (state.status === 'idle') return null
  return (
    <div className="summary-item">
      <span className="summary-dot" style={{ background: color }} />
      <span className="summary-label">{label}</span>
      <span className="summary-value">
        {state.status === 'loading' && '読み込み中...'}
        {state.status === 'success' && `${state.name} — ${state.rowCount?.toLocaleString()} 行`}
        {state.status === 'error' && `エラー: ${state.error}`}
      </span>
    </div>
  )
}
