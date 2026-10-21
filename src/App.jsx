import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { Copy, MapPinned, Search, Upload, X } from 'lucide-react'

const sampleRows = [
  { awb_nr: 'PH87493988835E', pkg_value: 845, last_location: 'Jeddah', last_scanned_by: 'Mohammed Ali', last_scanned_at: '2026-10-01T08:12:00', priority: 'High' },
  { awb_nr: 'PH89007388527E', pkg_value: 1120, last_location: 'Riyadh', last_scanned_by: 'Sara Ahmed', last_scanned_at: '2026-10-01T08:20:00', priority: 'Medium' },
  { awb_nr: 'PH61256112269E', pkg_value: 570, last_location: 'Dammam', last_scanned_by: 'Ahmed Salem', last_scanned_at: '2026-10-01T09:05:00', priority: 'High' },
  { awb_nr: 'PH76451233218E', pkg_value: 930, last_location: 'Jeddah', last_scanned_by: 'Sara Ahmed', last_scanned_at: '2026-10-01T09:48:00', priority: 'High' },
  { awb_nr: 'PH99163350226E', pkg_value: 640, last_location: 'Riyadh', last_scanned_by: 'Mohammed Ali', last_scanned_at: '2026-10-01T10:02:00', priority: 'Low' },
  { awb_nr: 'PH55491022786E', pkg_value: 760, last_location: 'Dammam', last_scanned_by: 'Nora Hassan', last_scanned_at: '2026-10-01T10:18:00', priority: 'Medium' },
  { awb_nr: 'PH24001198655E', pkg_value: 430, last_location: 'Madinah', last_scanned_by: 'Yousef Nasser', last_scanned_at: '2026-10-01T11:03:00', priority: 'Low' },
  { awb_nr: 'PH38944001176E', pkg_value: 1185, last_location: 'Jeddah', last_scanned_by: 'Nora Hassan', last_scanned_at: '2026-10-01T11:40:00', priority: 'High' },
  { awb_nr: 'PH48022499186E', pkg_value: 690, last_location: 'Madinah', last_scanned_by: 'Sara Ahmed', last_scanned_at: '2026-10-01T12:14:00', priority: 'Medium' },
  { awb_nr: 'PH73048718915E', pkg_value: 2150, last_location: 'Riyadh', last_scanned_by: 'Ahmed Salem', last_scanned_at: '2026-10-01T12:55:00', priority: 'High' },
  { awb_nr: 'PH13041084549E', pkg_value: 990, last_location: 'Dammam', last_scanned_by: 'Yousef Nasser', last_scanned_at: '2026-10-01T13:25:00', priority: 'Medium' },
  { awb_nr: 'PH66920310957E', pkg_value: 460, last_location: 'Madinah', last_scanned_by: 'Mohammed Ali', last_scanned_at: '2026-10-01T13:42:00', priority: 'Low' },
]

const requiredHeaders = ['awb_nr', 'pkg_value', 'last_location', 'last_scanned_by', 'last_scanned_at', 'priority']

const normalizeHeader = (value) => {
  const text = String(value ?? '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_')
  const aliases = {
    awb: 'awb_nr',
    awb_nr: 'awb_nr',
    package_value: 'pkg_value',
    pkg_value: 'pkg_value',
    pkgvalue: 'pkg_value',
    lastlocation: 'last_location',
    last_location: 'last_location',
    location: 'last_location',
    scanned_by: 'last_scanned_by',
    last_scanned_by: 'last_scanned_by',
    scanned_at: 'last_scanned_at',
    last_scanned_at: 'last_scanned_at',
    priority_level: 'priority',
    priority: 'priority',
  }

  return aliases[text] || text
}

const normalizeText = (value) => String(value ?? '').trim()

const getLocationSortRank = (value) => {
  const text = normalizeText(value).toUpperCase()
  const lmsgMatch = text.match(/LMSG\s*([0-9]+)/i)
  const lmsgNumber = lmsgMatch ? Number(lmsgMatch[1]) : Number.MAX_SAFE_INTEGER

  let statusRank = 99
  if (text.includes('HELD')) statusRank = 0
  else if (text.includes('FPDD')) statusRank = 1
  else if (text.includes('OFD')) statusRank = 2

  return { lmsgNumber, statusRank, label: text }
}

const sortLocations = (list) => {
  return [...list].sort((a, b) => {
    const aRank = getLocationSortRank(a.last_location ?? a)
    const bRank = getLocationSortRank(b.last_location ?? b)

    if (aRank.lmsgNumber !== bRank.lmsgNumber) {
      return aRank.lmsgNumber - bRank.lmsgNumber
    }

    if (aRank.statusRank !== bRank.statusRank) {
      return aRank.statusRank - bRank.statusRank
    }

    return aRank.label.localeCompare(bRank.label)
  })
}

const parseDelimitedText = (text) => {
  const trimmed = text.trim()
  if (!trimmed) return []

  const rows = []
  let currentValue = ''
  let row = []
  let inQuotes = false

  const pushCell = () => {
    row.push(currentValue)
    currentValue = ''
  }

  const pushRow = () => {
    if (row.length > 0 || currentValue) {
      pushCell()
      rows.push(row)
      row = []
    }
  }

  for (let idx = 0; idx < trimmed.length; idx += 1) {
    const char = trimmed[idx]

    if (char === '"') {
      if (inQuotes && trimmed[idx + 1] === '"') {
        currentValue += '"'
        idx += 1
      } else {
        inQuotes = !inQuotes
      }
      continue
    }

    if (char === '\n' || char === '\r') {
      if (inQuotes) {
        currentValue += char
      } else {
        pushRow()
      }
      continue
    }

    if (char === '\t' || char === ',') {
      if (!inQuotes) {
        pushCell()
        continue
      }
    }

    currentValue += char
  }

  if (currentValue || row.length) {
    pushRow()
  }

  return rows.filter((entry) => entry.some((cell) => normalizeText(cell).length > 0))
}

const normalizeRecord = (record) => {
  const normalized = {}

  Object.entries(record).forEach(([key, value]) => {
    const mappedKey = normalizeHeader(key)
    let mappedValue = value

    if (mappedKey === 'pkg_value') {
      const numericValue = Number(String(value ?? '').replace(/[^0-9.-]/g, ''))
      mappedValue = Number.isFinite(numericValue) ? numericValue : 0
    }

    if (mappedKey === 'priority') {
      mappedValue = String(value ?? '').trim() || 'Unknown'
    }

    normalized[mappedKey] = mappedValue
  })

  requiredHeaders.forEach((field) => {
    if (!(field in normalized)) {
      normalized[field] = field === 'pkg_value' ? 0 : ''
    }
  })

  return normalized
}

const parseImportedRows = (rows) => {
  if (!rows || rows.length < 2) return []

  const headerRow = rows[0].map((header) => normalizeHeader(header))
  const dataRows = rows.slice(1)

  return dataRows
    .filter((row) => row.some((cell) => normalizeText(cell).length > 0))
    .map((row) => {
      const entry = {}
      headerRow.forEach((header, index) => {
        entry[header] = row[index] ?? ''
      })
      return normalizeRecord(entry)
    })
}

const buildPivotRows = (rows) => {
  const grouped = new Map()

  rows.forEach((row) => {
    const location = normalizeText(row.last_location) || 'Unknown'
    const bucket = grouped.get(location) || {
      last_location: location,
      count: 0,
      awb_nr: [],
      records: [],
    }

    bucket.count += 1
    bucket.awb_nr.push(normalizeText(row.awb_nr))
    bucket.records.push(row)
    grouped.set(location, bucket)
  })

  return Array.from(grouped.values()).map((entry) => ({
    last_location: entry.last_location,
    count: entry.count,
    awb_nr: entry.awb_nr.filter(Boolean),
    records: entry.records,
  }))
}

const chunkArray = (items, chunkSize = 100) => {
  const chunks = []
  for (let index = 0; index < items.length; index += chunkSize) {
    chunks.push(items.slice(index, index + chunkSize))
  }
  return chunks
}

const makeTsv = (rows, headers) => {
  const headerLine = headers.join('\t')
  const body = rows.map((row) =>
    headers
      .map((header) => {
        const value = row[header]
        if (header === 'pkg_value') {
          return Number(value || 0).toFixed(2)
        }
        return String(value ?? '').replace(/\t|\n|\r/g, ' ')
      })
      .join('\t'),
  )

  return [headerLine, ...body].join('\n')
}

const copyText = async (text) => {
  if (navigator.clipboard && window.isSecureContext) {
    await navigator.clipboard.writeText(text)
    return
  }

  const helper = document.createElement('textarea')
  helper.value = text
  helper.setAttribute('readonly', '')
  helper.style.position = 'fixed'
  helper.style.left = '-9999px'
  document.body.appendChild(helper)
  helper.select()
  document.execCommand('copy')
  document.body.removeChild(helper)
}

const formatDate = (value) => {
  if (!value) return '—'
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return value
  return new Intl.DateTimeFormat('en-US', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  }).format(date)
}

const formatCurrency = (value) => {
  const numeric = Number(value) || 0
  return `${numeric.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
}

const BrandLogo = ({ className = '' }) => (
  <div className={`brand-mark ${className}`} aria-label="Brand logo">
    <span>P</span>
  </div>
)

function App() {
  const [rawData, setRawData] = useState(sampleRows)
  const [importText, setImportText] = useState('')
  const [toast, setToast] = useState('')
  const [pivotSearch, setPivotSearch] = useState('')
  const [selectedLocation, setSelectedLocation] = useState(null)
  const [modalSearch, setModalSearch] = useState('')
  const [summaryStatus, setSummaryStatus] = useState(() => Object.fromEntries(sampleRows.map((row) => [row.last_location, 'Pending'])))
  const [isProcessing, setIsProcessing] = useState(false)
  const [activeCopyKey, setActiveCopyKey] = useState(null)
  const [isServiceEnabled, setIsServiceEnabled] = useState(true)
  const [isServiceCheckLoading, setIsServiceCheckLoading] = useState(true)
  const fileInputRef = useRef(null)
  const copyResetTimerRef = useRef(null)

  const refreshServiceStatus = async () => {
    const apiBaseUrl = import.meta.env.VITE_API_URL || ''

    try {
      const response = await fetch(`${apiBaseUrl}/api/services/shipment`)
      if (!response.ok) {
        throw new Error('Service status request failed')
      }

      const data = await response.json()
      setIsServiceEnabled(Boolean(data.enabled))
    } catch (error) {
      console.error('Unable to fetch service status:', error)
      setIsServiceEnabled(true)
    } finally {
      setIsServiceCheckLoading(false)
    }
  }

  useEffect(() => {
    refreshServiceStatus()
  }, [])

  useEffect(() => {
    if (!toast) return undefined
    const timer = window.setTimeout(() => setToast(''), 2200)
    return () => window.clearTimeout(timer)
  }, [toast])

  useEffect(() => {
    return () => {
      if (copyResetTimerRef.current) {
        window.clearTimeout(copyResetTimerRef.current)
      }
    }
  }, [])

  const triggerCopyState = (id, duration = 1600) => {
    if (copyResetTimerRef.current) {
      window.clearTimeout(copyResetTimerRef.current)
    }

    setActiveCopyKey(id)
    copyResetTimerRef.current = window.setTimeout(() => setActiveCopyKey(null), duration)
  }

  const locationSummary = useMemo(() => {
    const grouped = new Map()

    rawData.forEach((row) => {
      const location = normalizeText(row.last_location) || 'Unknown'
      if (!grouped.has(location)) {
        grouped.set(location, { last_location: location, awb_nr: [] })
      }
      grouped.get(location).awb_nr.push(normalizeText(row.awb_nr))
    })

    return sortLocations(Array.from(grouped.values())).map((entry) => {
      const awbList = entry.awb_nr.filter(Boolean)
      const awbGroups = chunkArray(awbList, 100)

      return {
        ...entry,
        awb_nr: awbList,
        awbText: awbList.join(','),
        awbGroups,
      }
    })
  }, [rawData])

  const maxCopyColumns = Math.max(1, ...locationSummary.map((entry) => entry.awbGroups.length))

  const pivotRows = useMemo(() => {
    const built = sortLocations(buildPivotRows(rawData))
    const query = normalizeText(pivotSearch).toLowerCase()

    if (!query) return built
    return built.filter((row) =>
      normalizeText(row.last_location).toLowerCase().includes(query) ||
      String(row.count).includes(query),
    )
  }, [pivotSearch, rawData])

  const selectedRows = useMemo(() => {
    if (!selectedLocation) return []
    const query = normalizeText(modalSearch).toLowerCase()
    return rawData.filter((row) => {
      const matchesLocation = normalizeText(row.last_location) === normalizeText(selectedLocation)
      if (!matchesLocation) return false
      if (!query) return true
      return Object.values(row).some((value) => normalizeText(value).toLowerCase().includes(query))
    })
  }, [modalSearch, rawData, selectedLocation])

  const applyImportedRows = (rows) => {
    setIsProcessing(true)

    window.setTimeout(() => {
      const imported = parseImportedRows(rows)
      if (!imported.length) {
        setToast('No valid rows were found in the pasted data.')
        setIsProcessing(false)
        return
      }

      setRawData(imported)
      setImportText('')
      setPivotSearch('')
      setToast('Data imported successfully.')
      setIsProcessing(false)
    }, 2600 )
  }

  const handlePasteImport = () => {
    const rows = parseDelimitedText(importText)
    if (!rows.length) {
      setToast('The pasted data is empty or invalid.')
      return
    }
    applyImportedRows(rows)
  }

  const handleFileImport = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      if (file.name.toLowerCase().endsWith('.csv') || file.name.toLowerCase().endsWith('.txt')) {
        const text = await file.text()
        applyImportedRows(parseDelimitedText(text))
      } else {
        const buffer = await file.arrayBuffer()
        const workbook = XLSX.read(buffer, { type: 'array' })
        const worksheet = workbook.Sheets[workbook.SheetNames[0]]
        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false })
        if (!jsonRows.length) {
          setToast('The spreadsheet file has no rows to import.')
          return
        }

        const firstRow = Object.keys(jsonRows[0])
        const rawRows = jsonRows.map((row) => firstRow.map((header) => row[header]))
        applyImportedRows([firstRow, ...rawRows])
      }
    } catch (error) {
      console.error(error)
      setToast('Unable to read this file. Please use CSV or XLSX format.')
    } finally {
      event.target.value = ''
    }
  }

  const handleCopySummary = async (text, label, key) => {
    try {
      await copyText(text)
      triggerCopyState(key)
      setToast(label)
    } catch (error) {
      console.error(error)
      setToast('Copy failed.')
    }
  }

  const handleCopyPivot = async () => {
    const headers = ['last_location', 'Count of awb_nr']
    const rows = pivotRows.map((row) => ({ last_location: row.last_location, 'Count of awb_nr': row.count }))

    try {
      await copyText(makeTsv(rows, headers))
      triggerCopyState('pivot-copy')
      setToast('Pivot table copied to Excel.')
    } catch (error) {
      console.error(error)
      setToast('Unable to copy pivot table.')
    }
  }

  const handleCopyDrilldown = async () => {
    const headers = ['awb_nr', 'pkg_value', 'last_scanned_by', 'last_scanned_at', 'priority']
    const rows = selectedRows.map((row) => ({
      awb_nr: row.awb_nr,
      pkg_value: row.pkg_value,
      last_scanned_by: row.last_scanned_by,
      last_scanned_at: row.last_scanned_at,
      priority: row.priority,
    }))

    try {
      await copyText(makeTsv(rows, headers))
      triggerCopyState('modal-copy')
      setToast('Drill-down data copied to Excel.')
    } catch (error) {
      console.error(error)
      setToast('Unable to copy drill-down data.')
    }
  }

  const handleServiceToggle = async () => {
    const nextState = !isServiceEnabled
    const apiBaseUrl = import.meta.env.VITE_API_URL || ''

    try {
      const response = await fetch(`${apiBaseUrl}/api/services/shipment`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ enabled: nextState }),
      })

      if (!response.ok) {
        throw new Error('Failed to update service status')
      }

      const data = await response.json()
      setIsServiceEnabled(Boolean(data.enabled))
      setToast(data.enabled ? 'Service enabled successfully.' : 'Service disabled successfully.')
    } catch (error) {
      console.error(error)
      setToast('Unable to change service status right now.')
    }
  }

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900" dir="ltr">
      <div className="mx-auto max-w-7xl px-4 py-6 md:px-8 md:py-8">
        <header className="mb-6 rounded-3xl bg-gradient-to-r from-[#0f172a] via-[#1e293b] to-[#0b1220] p-6 text-white shadow-xl ring-1 ring-white/10">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div className="flex items-center gap-3">
              <div className="rounded-2xl bg-white/6 p-3 backdrop-blur-sm shadow-lg shadow-amber-500/10">
                <BrandLogo />
              </div>
              <div>
                <p className="text-xs font-medium uppercase tracking-[0.22em] text-amber-200/90">Shipment Data Hub</p>
                <h1 className="text-2xl font-black md:text-3xl">Pivot Table & Drill-down Processor</h1>
              </div>
            </div>

            <button
              type="button"
              onClick={handleServiceToggle}
              disabled={isServiceCheckLoading}
              className={`rounded-full px-4 py-2 text-sm font-bold transition admin-button ${
                isServiceEnabled
                  ? 'bg-rose-500 text-white hover:bg-rose-400'
                  : 'bg-emerald-500 text-white hover:bg-emerald-400'
              } disabled:cursor-not-allowed disabled:opacity-60`}
            >
              {isServiceCheckLoading ? 'جاري التحقق...' : isServiceEnabled ? 'إيقاف الخدمة' : 'تفعيل الخدمة'}
            </button>
          </div>
        </header>

        {isServiceCheckLoading ? (
          <section className="flex min-h-[420px] items-center justify-center rounded-3xl border border-slate-200 bg-white p-10 shadow-sm">
            <div className="text-center">
              <div className="mx-auto mb-4 h-16 w-16 animate-spin rounded-full border-4 border-slate-200 border-t-emerald-500" aria-label="Checking service status" />
              <p className="mb-2 text-sm font-bold uppercase tracking-[0.28em] text-slate-500">Checking</p>
              <h2 className="text-3xl font-black text-slate-800">جاري التحقق من حالة الخدمة...</h2>
            </div>
          </section>
        ) : !isServiceEnabled ? (
          <section className="flex min-h-[420px] items-center justify-center rounded-3xl border border-slate-200 bg-white p-10 shadow-sm">
            <div className="text-center">
              <div className="mx-auto mb-4 h-16 w-16 animate-spin rounded-full border-4 border-slate-200 border-t-rose-500" aria-label="Service is disabled" />
              <p className="mb-2 text-sm font-bold uppercase tracking-[0.28em] text-slate-500">Service Status</p>
              <h2 className="text-3xl font-black text-slate-800">لقد أوقف المطور هذه الخدمة</h2>
              <p className="mt-3 text-slate-600">تم إيقاف هذه الخدمة من الخادم، ويمكن إعادة تفعيلها من واجهة المطور لاحقًا.</p>
            </div>
          </section>
        ) : (
          <>
            {isProcessing && (
              <div className="mb-6 flex items-center justify-center rounded-3xl border border-amber-200 bg-amber-50/80 p-5 shadow-sm">
                <div className="processing-shell" aria-live="polite" aria-label="Processing data">
                  <div className="processing-ring" />
                  <BrandLogo className="processing-mark" />
                </div>
                <div className="ml-4 text-left">
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-amber-700">Processing</p>
                  <p className="text-lg font-black text-slate-800">Preparing shipment results...</p>
                </div>
              </div>
            )}

            <section className="grid gap-6 xl:grid-cols-[420px_minmax(0,1fr)]">
              <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="mb-4 flex items-center justify-between">
                  <h2 className="flex items-center gap-2 text-xl font-bold text-slate-800">
                    <Upload className="h-5 w-5 text-emerald-600" />
                    Paste Data
                  </h2>
                  <button type="button" onClick={() => setImportText('')} className="rounded-full p-2 text-slate-500 hover:bg-slate-100">
                    <X className="h-4 w-4" />
                  </button>
                </div>

                <textarea
                  value={importText}
                  onChange={(event) => setImportText(event.target.value)}
                  onPaste={(event) => {
                    const pasted = event.clipboardData.getData('text/plain')
                    if (pasted) {
                      event.preventDefault()
                      setImportText(pasted)
                      setTimeout(() => {
                        const rows = parseDelimitedText(pasted)
                        if (rows.length) applyImportedRows(rows)
                      }, 0)
                    }
                  }}
                  placeholder="Paste data copied from Excel or Google Sheets (CSV or TSV)"
                  className="min-h-[190px] w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm text-slate-700 outline-none transition focus:border-emerald-500 focus:bg-white"
                />

                <div className="mt-4 grid grid-cols-2 gap-3">
                  <button type="button" onClick={handlePasteImport} className="rounded-2xl bg-emerald-600 px-3 py-2.5 text-sm font-bold text-white transition hover:bg-emerald-500">
                    Import Paste
                  </button>
                  <button type="button" onClick={() => fileInputRef.current?.click()} className="rounded-2xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-bold text-slate-700 transition hover:bg-slate-50">
                    Upload CSV/XLSX
                  </button>
                  <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls,.txt" onChange={handleFileImport} className="hidden" />
                </div>

                <div className="mt-5 rounded-2xl border border-emerald-200 bg-emerald-50 p-3 text-sm text-emerald-900">
                  <p className="font-semibold">Required fields</p>
                  <p className="mt-2">awb_nr • pkg_value • last_location • last_scanned_by • last_scanned_at • priority</p>
                </div>

                <button type="button" onClick={() => setRawData(sampleRows)} className="mt-4 w-full rounded-2xl border border-emerald-200 bg-white px-3 py-2.5 text-sm font-bold text-emerald-700 transition hover:bg-emerald-50">
                  Restore Sample Data
                </button>
              </div>

              <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="mb-4 flex items-center justify-between gap-3">
                  <h2 className="flex items-center gap-2 text-xl font-bold text-slate-800">
                    <MapPinned className="h-5 w-5 text-emerald-600" />
                    Location Summary
                  </h2>
                  <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-bold text-emerald-700">
                    {locationSummary.length} locations
                  </span>
                </div>

                <div className="overflow-auto rounded-2xl border border-slate-200">
                  <table className="min-w-full border-collapse text-sm">
                    <thead>
                      <tr className="bg-emerald-700 text-white">
                        <th className="border border-emerald-800 px-3 py-3 text-left font-bold">last_location</th>
                        {Array.from({ length: maxCopyColumns }).map((_, index) => (
                          <th key={`copy-header-${index + 1}`} className="border border-emerald-800 px-3 py-3 text-left font-bold">
                            Copy {index + 1}
                          </th>
                        ))}
                        <th className="border border-emerald-800 px-3 py-3 text-left font-bold">Status</th>
                        <th className="border border-emerald-800 px-3 py-3 text-left font-bold">awb_nr</th>
                      </tr>
                    </thead>
                    <tbody>
                      {locationSummary.map((entry) => (
                        <tr key={entry.last_location} className="odd:bg-white even:bg-slate-50">
                          <td className="border border-slate-200 px-3 py-3 text-left font-medium text-slate-700">{entry.last_location}</td>

                          {entry.awbGroups.map((group, index) => (
                            <td key={`${entry.last_location}-copy-${index + 1}`} className="border border-slate-200 px-3 py-3 text-left align-top">
                              <button
                                type="button"
                                onClick={() => handleCopySummary(group.join(','), `Copied AWB group ${index + 1} for ${entry.last_location}.`, `summary-${entry.last_location}-${index}`)}
                                className={`copy-button ${activeCopyKey === `summary-${entry.last_location}-${index}` ? 'active' : ''}`}
                              >
                                <Copy className="h-3.5 w-3.5" />
                                Copy {index + 1}
                              </button>
                            </td>
                          ))}
                          {Array.from({ length: maxCopyColumns - entry.awbGroups.length }).map((_, index) => (
                            <td key={`${entry.last_location}-empty-${index}`} className="border border-slate-200 bg-slate-50 px-3 py-3 text-left" />
                          ))}
                          <td className="border border-slate-200 px-3 py-3 text-left">
                            <div className="flex items-center gap-3">
                              <label className="flex cursor-pointer items-center gap-1.5 text-xs text-slate-700">
                                <input
                                  type="radio"
                                  name={`status-${entry.last_location}`}
                                  checked={(summaryStatus[entry.last_location] ?? 'Pending') === 'Done'}
                                  onChange={() => setSummaryStatus((prev) => ({ ...prev, [entry.last_location]: 'Done' }))}
                                />
                                Done
                              </label>
                              <label className="flex cursor-pointer items-center gap-1.5 text-xs text-slate-700">
                                <input
                                  type="radio"
                                  name={`status-${entry.last_location}`}
                                  checked={(summaryStatus[entry.last_location] ?? 'Pending') === 'Pending'}
                                  onChange={() => setSummaryStatus((prev) => ({ ...prev, [entry.last_location]: 'Pending' }))}
                                />
                                Pending
                              </label>
                            </div>
                          </td>
                          <td className="border border-slate-200 px-3 py-3 text-left text-slate-700">{entry.awbText}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>

            <section className="mt-6 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                <div className="flex items-center gap-3">
                  <h2 className="text-xl font-bold text-slate-800">Immediate Pivot Table</h2>
                </div>

                <div className="flex flex-wrap items-center gap-3">
                  <div className="flex items-center gap-2 rounded-2xl border border-slate-200 bg-slate-50 px-3 py-2.5">
                    <Search className="h-4 w-4 text-slate-500" />
                    <input value={pivotSearch} onChange={(event) => setPivotSearch(event.target.value)} placeholder="Filter locations" className="w-44 border-0 bg-transparent text-sm text-slate-700 outline-none placeholder:text-slate-400" />
                  </div>
                  <button type="button" onClick={handleCopyPivot} className={`copy-button ${activeCopyKey === 'pivot-copy' ? 'active' : ''}`}>
                    <Copy className="h-4 w-4" />
                    Copy Table to Excel
                  </button>
                </div>
              </div>

              <div className="overflow-auto rounded-2xl border border-slate-200">
                <table className="min-w-full border-collapse text-sm">
                  <thead>
                    <tr className="bg-emerald-700 text-white">
                      <th className="border border-emerald-800 px-3 py-3 text-left font-bold">last_location</th>
                      <th className="border border-emerald-800 px-3 py-3 text-left font-bold">Count of awb_nr</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pivotRows.length ? (
                      pivotRows.map((row) => (
                        <tr key={row.last_location} className="odd:bg-white even:bg-slate-50">
                          <td className="border border-slate-200 px-3 py-3 text-left font-medium text-slate-700">{row.last_location}</td>
                          <td className="border border-slate-200 px-3 py-3 text-left">
                            <button type="button" onClick={() => setSelectedLocation(row.last_location)} className="cursor-pointer rounded-lg bg-emerald-50 px-2.5 py-1 font-bold text-emerald-700 transition hover:bg-emerald-100">
                              {row.count}
                            </button>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={2} className="px-4 py-8 text-center text-slate-500">No matching rows found.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </>
        )}
      </div>

      {selectedLocation && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 px-4 backdrop-blur-sm">
          <div className="w-full max-w-5xl rounded-3xl bg-white shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-200 px-5 py-4">
              <div>
                <p className="text-sm text-slate-500">Drill-down</p>
                <h3 className="text-xl font-black text-slate-800">Location: {selectedLocation}</h3>
              </div>
              <button type="button" onClick={() => setSelectedLocation(null)} className="rounded-full p-2 hover:bg-slate-100">
                <X className="h-5 w-5 text-slate-600" />
              </button>
            </div>

            <div className="flex items-center justify-between gap-3 border-b border-slate-200 bg-slate-50 px-5 py-4">
              <div className="flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-3 py-2.5">
                <Search className="h-4 w-4 text-slate-500" />
                <input value={modalSearch} onChange={(event) => setModalSearch(event.target.value)} placeholder="Search modal records" className="w-56 border-0 bg-transparent text-sm text-slate-700 outline-none placeholder:text-slate-400" />
              </div>
              <div className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-bold text-emerald-700">
                Records: {selectedRows.length}
              </div>
              <button type="button" onClick={handleCopyDrilldown} className={`copy-button ${activeCopyKey === 'modal-copy' ? 'active' : ''}`}>
                <Copy className="h-4 w-4" />
                Copy Modal Table to Excel
              </button>
            </div>

            <div className="max-h-[70vh] overflow-auto p-5">
              <table className="min-w-full border-collapse text-sm">
                <thead>
                  <tr className="bg-emerald-700 text-white">
                    <th className="border border-emerald-800 px-3 py-3 text-left font-bold">awb_nr</th>
                    <th className="border border-emerald-800 px-3 py-3 text-left font-bold">pkg_value</th>
                    <th className="border border-emerald-800 px-3 py-3 text-left font-bold">last_scanned_by</th>
                    <th className="border border-emerald-800 px-3 py-3 text-left font-bold">last_scanned_at</th>
                    <th className="border border-emerald-800 px-3 py-3 text-left font-bold">priority</th>
                  </tr>
                </thead>
                <tbody>
                  {selectedRows.length ? (
                    selectedRows.map((row) => (
                      <tr key={row.awb_nr} className="odd:bg-white even:bg-slate-50">
                        <td className="border border-slate-200 px-3 py-3 text-left">{row.awb_nr}</td>
                        <td className="border border-slate-200 px-3 py-3 text-left">{formatCurrency(row.pkg_value)}</td>
                        <td className="border border-slate-200 px-3 py-3 text-left">{row.last_scanned_by}</td>
                        <td className="border border-slate-200 px-3 py-3 text-left">{formatDate(row.last_scanned_at)}</td>
                        <td className="border border-slate-200 px-3 py-3 text-left">{row.priority}</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={5} className="px-4 py-8 text-center text-slate-500">No records for this location.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div className="fixed bottom-5 left-1/2 z-[60] -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm font-bold text-white shadow-lg">
          {toast}
        </div>
      )}
    </div>
  )
}

export default App
