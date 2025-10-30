const fs = require('fs')
const path = require('path')
const xlsx = require('xlsx')
const BEFORE_DIR = path.join(__dirname, "../before")

const beforeFilesPath = fs.readdirSync(BEFORE_DIR).map(localPath => path.join(BEFORE_DIR, localPath))


const extract = (paramsArray) => {
    const [productName, quantity] = paramsArray
    
    const regex = new RegExp(/(.+?)\s*?(\d+)\s*?(.+?)/)
    const result = String(productName).match(regex)
    if(!result || (!result[1] && !result[2] && !result[3])) return null
    
    const [name, price, unit] = [String(result[[1]]), Number(result[2]), String(result[3])]
    
    return { name, price, unit, quantity: Number(quantity) }
}

const convertFilesToJson = (pathArray) => {
    const result = []
    
    for(const before of pathArray) {
        const file = xlsx.readFile(path.resolve(before))
        const sheetName = file.SheetNames[0]
        if(!sheetName) continue
    
        const sheet = file.Sheets[sheetName]
        const toJson = xlsx.utils.sheet_to_json(sheet)
        result.push({ content: toJson.map(Object.values), filename: path.basename(before).replace(path.extname(before), '') })
    }
    return result
}

const convertUnitStrToCharCode = (unit) => {
    return String(unit).charCodeAt(0)
}

const mapUnitCharCode = (code) => {
    const currencyMap = {
        165: "ژاپن",
        36: "امریکا",
        8364: "اروپا"
    }
    return currencyMap[code] ?? "Unknown"
}

const aggregate = (json) => {
    const country = mapUnitCharCode(convertUnitStrToCharCode(json.unit))
    return [json.name, json.price, country].join(' ')
}

const convertToCsv = (json) => {
    const sheet = xlsx.utils.json_to_sheet(json, {skipHeader: true})
    const csv = xlsx.utils.sheet_to_csv(sheet)
    return csv
}


const convertToBook = (json) => {
    const sheet = xlsx.utils.json_to_sheet(json, {skipHeader: true})
    const book = xlsx.utils.book_new()
    xlsx.utils.book_append_sheet(book, sheet, "finalized")
    return book
}


const writeXlsxFile = (book, filename) => {
    const fullname = path.join(__dirname, "../after", filename + ".xlsx")
    xlsx.writeFile(book, fullname)
}

const serialize = (name) => ({ name })

const processed = convertFilesToJson(beforeFilesPath)


for(const { content, filename } of processed)  {
    const convertedJsonArray = content
            .map(extract)
            .map(aggregate)
            .map(serialize)
    
    const book = convertToBook(convertedJsonArray)
    writeXlsxFile(book, filename)
}
 



