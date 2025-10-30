import * as XLSX from 'xlsx'

const extract = (paramsArray) => {
    const [productName, quantity] = paramsArray
    const regex = new RegExp(/(.+?)\s*?(\d+)\s*?(.+?)/)
    const result = String(productName).match(regex)
    if (!result || (!result[1] && !result[2] && !result[3])) return null
    const [name, price, unit] = [String(result[1]), Number(result[2]), String(result[3])]
    return { name, price, unit, quantity: Number(quantity) }
}

const convertUnitStrToCharCode = (unit) => String(unit).charCodeAt(0)

const mapUnitCharCode = (code) => {
    const currencyMap = { 165: 'ژاپن', 36: 'امریکا', 8364: 'اروپا' }
    return currencyMap[code] ?? 'Unknown'
}

const aggregate = (json) => {
    const country = mapUnitCharCode(convertUnitStrToCharCode(json.unit))
    return [json.name, json.price, country].join(' ')
}

const serialize = (name) => ({ name })

export async function readWorkbookFromFile(file) {
    const data = await file.arrayBuffer()
    const wb = XLSX.read(data, { type: 'array' })
    return wb
}

export function firstSheetToRows(workbook) {
    const sheetName = workbook.SheetNames?.[0]
    if (!sheetName) return []
    const sheet = workbook.Sheets[sheetName]
    const json = XLSX.utils.sheet_to_json(sheet)
    return json.map(Object.values)
}

export function convertRows(rows) {
    return rows
        .map(extract)
        .filter(Boolean)
        .map(aggregate)
        .map(serialize)
}

export function toWorkbookFromJson(jsonRows) {
    const sheet = XLSX.utils.json_to_sheet(jsonRows, { skipHeader: true })
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, sheet, 'finalized')
    return book
}

export function workbookToBlob(workbook, filenameBase) {
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    return { blob, filename: `${filenameBase}.xlsx` }
}

export function previewRows(rows, max = 5) {
    return rows.slice(0, max)
}

// Small UX helper: count rows that match the expected pattern so we can communicate
// how many lines are recognized before conversion. This helps set expectations for users.
export function countValidRows(rows) {
    const regex = new RegExp(/(.+?)\s*?(\d+)\s*?(.+?)/)
    let count = 0
    for (const r of rows) {
        const productName = r?.[0]
        if (productName && String(productName).match(regex)) count++
    }
    return count
}


