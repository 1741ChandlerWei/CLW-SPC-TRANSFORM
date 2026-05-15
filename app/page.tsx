'use client'

import { useState, useRef } from 'react'

type Mode = 'transform' | 'compare'
type Status = 'idle' | 'loading' | 'success' | 'error'

export default function Page() {
  const [mode, setMode] = useState<Mode>('transform')
  const [status, setStatus] = useState<Status>('idle')
  const [message, setMessage] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)
  const fileARef = useRef<HTMLInputElement>(null)
  const fileBRef = useRef<HTMLInputElement>(null)

  async function handleTransform() {
    const file = fileRef.current?.files?.[0]
    if (!file) { setStatus('error'); setMessage('請選擇檔案'); return }

    setStatus('loading')
    setMessage('解析中...')

    const fd = new FormData()
    fd.append('file', file)

    try {
      const res = await fetch('/api/parse-excel', { method: 'POST', body: fd })
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: '未知錯誤' }))
        setStatus('error')
        setMessage(err.error || '解析失敗')
        return
      }
      const blob = await res.blob()
      const cd = res.headers.get('Content-Disposition') || ''
      const nameMatch = cd.match(/filename\*?=(?:UTF-8'')?(.+)/)
      const fileName = nameMatch ? decodeURIComponent(nameMatch[1].replace(/"/g, '')) : '整理結果.xlsx'
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = fileName
      a.click()
      URL.revokeObjectURL(url)
      setStatus('success')
      setMessage(`✅ 完成！已下載 ${fileName}`)
    } catch {
      setStatus('error')
      setMessage('連線失敗，請再試一次')
    }
  }

  async function handleCompare() {
    const fileA = fileARef.current?.files?.[0]
    const fileB = fileBRef.current?.files?.[0]
    if (!fileA || !fileB) { setStatus('error'); setMessage('請選擇兩個檔案'); return }

    setStatus('loading')
    setMessage('比對中...')

    const fd = new FormData()
    fd.append('fileA', fileA)
    fd.append('fileB', fileB)

    try {
      const res = await fetch('/api/compare-excel', { method: 'POST', body: fd })
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: '未知錯誤' }))
        setStatus('error')
        setMessage(err.error || '比對失敗')
        return
      }
      const blob = await res.blob()
      const cd = res.headers.get('Content-Disposition') || ''
      const nameMatch = cd.match(/filename\*?=(?:UTF-8'')?(.+)/)
      const fileName = nameMatch ? decodeURIComponent(nameMatch[1].replace(/"/g, '')) : '比對結果.xlsx'
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = fileName
      a.click()
      URL.revokeObjectURL(url)
      setStatus('success')
      setMessage(`✅ 完成！已下載 ${fileName}`)
    } catch {
      setStatus('error')
      setMessage('連線失敗，請再試一次')
    }
  }

  return (
    <main style={styles.main}>
      <div style={styles.card}>
        {/* Header */}
        <div style={styles.header}>
          <div style={styles.logo}>CLW</div>
          <div>
            <h1 style={styles.title}>規格表轉換工具</h1>
            <p style={styles.subtitle}>CLW-SPC-TRANSFORM ｜ Spec Parser & Comparator</p>
          </div>
        </div>

        {/* Mode Toggle */}
        <div style={styles.modeBar}>
          <button
            style={{ ...styles.modeBtn, ...(mode === 'transform' ? styles.modeBtnActive : {}) }}
            onClick={() => { setMode('transform'); setStatus('idle') }}
          >
            📄 整理規格表
          </button>
          <button
            style={{ ...styles.modeBtn, ...(mode === 'compare' ? styles.modeBtnActive : {}) }}
            onClick={() => { setMode('compare'); setStatus('idle') }}
          >
            🔍 版次比對
          </button>
        </div>

        {/* Transform Mode */}
        {mode === 'transform' && (
          <div style={styles.section}>
            <p style={styles.desc}>
              上傳客供規格表（.xlsm / .xlsx），自動整理成標準格式：<br />
              <span style={styles.tag}>規格總覽</span>
              <span style={styles.tag}>部件重量</span>
              <span style={styles.tag}>密度（英制/公制）</span>
            </p>
            <label style={styles.fileLabel}>
              <input
                ref={fileRef}
                type="file"
                accept=".xls,.xlsx,.xlsm"
                style={{ display: 'none' }}
                onChange={() => setStatus('idle')}
              />
              <div style={styles.fileBox}>
                <span style={styles.fileIcon}>📁</span>
                <span>點此選擇規格表檔案</span>
                <span style={styles.fileHint}>支援 .xlsm / .xlsx / .xls</span>
              </div>
            </label>
            {fileRef.current?.files?.[0] && (
              <p style={styles.fileName}>已選：{fileRef.current.files[0].name}</p>
            )}
            <button
              style={{ ...styles.btn, ...(status === 'loading' ? styles.btnDisabled : {}) }}
              onClick={handleTransform}
              disabled={status === 'loading'}
            >
              {status === 'loading' ? '處理中...' : '開始整理 → 下載 Excel'}
            </button>
          </div>
        )}

        {/* Compare Mode */}
        {mode === 'compare' && (
          <div style={styles.section}>
            <p style={styles.desc}>
              上傳兩個版次的規格表（順序 = 舊版 → 新版），自動比對差異，<br />
              標示新增 🟢、變更 🟡、移除 🔴
            </p>
            <div style={styles.compareRow}>
              <div style={styles.compareBox}>
                <p style={styles.compareLabel}>舊版（Version A）</p>
                <label style={styles.fileLabel}>
                  <input ref={fileARef} type="file" accept=".xls,.xlsx,.xlsm" style={{ display: 'none' }} onChange={() => setStatus('idle')} />
                  <div style={styles.fileBox}>
                    <span style={styles.fileIcon}>📁</span>
                    <span>選擇舊版檔案</span>
                  </div>
                </label>
              </div>
              <div style={styles.arrow}>→</div>
              <div style={styles.compareBox}>
                <p style={styles.compareLabel}>新版（Version B）</p>
                <label style={styles.fileLabel}>
                  <input ref={fileBRef} type="file" accept=".xls,.xlsx,.xlsm" style={{ display: 'none' }} onChange={() => setStatus('idle')} />
                  <div style={styles.fileBox}>
                    <span style={styles.fileIcon}>📁</span>
                    <span>選擇新版檔案</span>
                  </div>
                </label>
              </div>
            </div>
            <button
              style={{ ...styles.btn, ...(status === 'loading' ? styles.btnDisabled : {}) }}
              onClick={handleCompare}
              disabled={status === 'loading'}
            >
              {status === 'loading' ? '比對中...' : '開始比對 → 下載 Excel'}
            </button>
          </div>
        )}

        {/* Status */}
        {status !== 'idle' && message && (
          <div style={{
            ...styles.statusBox,
            background: status === 'success' ? '#e8f5e9' : status === 'error' ? '#ffebee' : '#e3f2fd',
            borderColor: status === 'success' ? '#a5d6a7' : status === 'error' ? '#ef9a9a' : '#90caf9',
          }}>
            {message}
          </div>
        )}

        <p style={styles.footer}>根莖葉工作室 ｜ CLW-SPC-TRANSFORM</p>
      </div>
    </main>
  )
}

const styles: Record<string, React.CSSProperties> = {
  main: {
    minHeight: '100vh',
    background: 'linear-gradient(135deg, #1a237e 0%, #0d47a1 50%, #01579b 100%)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '24px',
    fontFamily: "'Arial', 'Microsoft JhengHei', sans-serif",
  },
  card: {
    background: '#ffffff',
    borderRadius: '16px',
    boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
    padding: '40px',
    maxWidth: '680px',
    width: '100%',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
    marginBottom: '28px',
    paddingBottom: '20px',
    borderBottom: '2px solid #e3f2fd',
  },
  logo: {
    background: 'linear-gradient(135deg, #1565c0, #0288d1)',
    color: '#fff',
    fontWeight: 'bold',
    fontSize: '20px',
    width: '52px',
    height: '52px',
    borderRadius: '12px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  title: { margin: 0, fontSize: '22px', fontWeight: 700, color: '#1a237e' },
  subtitle: { margin: '4px 0 0', fontSize: '12px', color: '#78909c' },
  modeBar: {
    display: 'flex',
    gap: '8px',
    marginBottom: '24px',
    background: '#f5f7fa',
    padding: '4px',
    borderRadius: '10px',
  },
  modeBtn: {
    flex: 1,
    padding: '10px 16px',
    border: 'none',
    borderRadius: '8px',
    background: 'transparent',
    cursor: 'pointer',
    fontSize: '14px',
    fontWeight: 500,
    color: '#546e7a',
    transition: 'all 0.2s',
  },
  modeBtnActive: {
    background: '#1565c0',
    color: '#fff',
    boxShadow: '0 2px 8px rgba(21,101,192,0.4)',
  },
  section: { display: 'flex', flexDirection: 'column', gap: '16px' },
  desc: {
    margin: 0,
    fontSize: '14px',
    color: '#546e7a',
    lineHeight: 1.7,
    background: '#f5f7fa',
    padding: '12px 16px',
    borderRadius: '8px',
    borderLeft: '3px solid #1565c0',
  },
  tag: {
    display: 'inline-block',
    background: '#e3f2fd',
    color: '#1565c0',
    borderRadius: '4px',
    padding: '1px 8px',
    fontSize: '12px',
    fontWeight: 600,
    margin: '2px 4px 0 0',
  },
  fileLabel: { cursor: 'pointer' },
  fileBox: {
    border: '2px dashed #90caf9',
    borderRadius: '10px',
    padding: '20px',
    textAlign: 'center',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '6px',
    fontSize: '14px',
    color: '#546e7a',
    transition: 'border-color 0.2s',
    background: '#fafcff',
  },
  fileIcon: { fontSize: '28px' },
  fileHint: { fontSize: '11px', color: '#90a4ae' },
  fileName: { margin: 0, fontSize: '13px', color: '#1565c0', fontWeight: 500 },
  btn: {
    background: 'linear-gradient(135deg, #1565c0, #0288d1)',
    color: '#fff',
    border: 'none',
    borderRadius: '10px',
    padding: '14px 24px',
    fontSize: '15px',
    fontWeight: 600,
    cursor: 'pointer',
    boxShadow: '0 4px 16px rgba(21,101,192,0.3)',
  },
  btnDisabled: { opacity: 0.6, cursor: 'not-allowed' },
  compareRow: {
    display: 'flex',
    gap: '12px',
    alignItems: 'center',
  },
  compareBox: { flex: 1 },
  compareLabel: {
    margin: '0 0 8px',
    fontSize: '13px',
    fontWeight: 600,
    color: '#1565c0',
  },
  arrow: { fontSize: '24px', color: '#90a4ae', flexShrink: 0 },
  statusBox: {
    padding: '12px 16px',
    borderRadius: '8px',
    border: '1px solid',
    fontSize: '14px',
    fontWeight: 500,
  },
  footer: {
    marginTop: '24px',
    textAlign: 'center',
    fontSize: '11px',
    color: '#b0bec5',
  },
}
