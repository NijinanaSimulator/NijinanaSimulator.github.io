const appPage = document.body.dataset.page === "deck" ? "deck" : "board"
const byId = id => document.getElementById(id)
const escapeHTML = value => String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll("\"", "&quot;")
const copy = value => JSON.parse(JSON.stringify(value))
const uid = prefix => `${prefix}_${[...crypto.getRandomValues(new Uint32Array(4))].map(value => value.toString(16).padStart(8, "0")).join("")}`
const playerIds = ["p1", "p2"]
const nijiColors = ["R", "O", "Y", "G", "B", "I", "V"]
const textCardTypes = {other: "未設定", liver: "ライバーカード", event: "イベントカード", stage: "ステージカード", channel: "チャンネルカード"}
const textCardClasses = ["太陽", "月", "彗星", "星"]
const CARD_TAG_LIMITS = {text: 2000, count: 50, tag: 100}
function normalizeCardTags(raw) {
    const values = typeof raw === "string" ? raw.split(/[,，、]/u) : raw
    if (!Array.isArray(values) || values.some(value => typeof value !== "string") || values.join(",").length > CARD_TAG_LIMITS.text) throw new Error("タグはカンマ区切りで2000文字以内で入力してください。")
    const tags = [...new Set(values.map(value => value.replace(/[\s#＃]+/gu, "")).filter(Boolean))]
    if (tags.length > CARD_TAG_LIMITS.count || tags.some(tag => tag.length > CARD_TAG_LIMITS.tag)) throw new Error("タグは50個まで、各100文字以内で入力してください。")
    return tags
}
function validCardTags(tags) {
    if (!Array.isArray(tags)) return false
    try {return JSON.stringify(tags) === JSON.stringify(normalizeCardTags(tags))} catch {return false}
}
function liverCardTags(item) {return item && textCardType(item) === "liver" && Array.isArray(item.tags) ? item.tags : []}
function cardTagsHTML(item, className = "card-detail-tags") {
    const tags = liverCardTags(item)
    return tags.length ? `<div class="${escapeHTML(className)}">${tags.map(tag => `<span>#${escapeHTML(tag)}</span>`).join("")}</div>` : ""
}
// CARD_LIBRARY_START
// Offline XLSX import. Card effects remain plain text and never run as rules.
const XLSX_CARD_LIMITS = {fileBytes: 10 * 1024 * 1024, entryBytes: 20 * 1024 * 1024, totalBytes: 50 * 1024 * 1024, entries: 2000, rows: 20000, cells: 250000, cards: 1000}
const XLSX_CARD_HEADERS = ["ID", "カード名", "種類", "クラス", "レベル", "Power"]
const XLSX_CARD_HEADER_ALIASES = {ID: ["ID", "仮ID"], レベル: ["レベル", "Lv", "Level", "コスト", "Cost"]}
const XLSX_CARD_TYPES = {LIVER: "liver", ライバー: "liver", ライバーカード: "liver", EVENT: "event", イベント: "event", イベントカード: "event", STAGE: "stage", ステージ: "stage", ステージカード: "stage", CHANNEL: "channel", チャンネル: "channel", チャンネルカード: "channel"}
const XLSX_CARD_CLASSES = {SOL: "太陽", LUNA: "月", COMET: "彗星", STELLA: "星", 太陽: "太陽", 月: "月", 彗星: "彗星", 星: "星"}
const XLSX_CARD_COLOR_HEADERS = ["colors", "color", "色", "カラー", "ライバーの色"]
const XLSX_CARD_TAG_HEADERS = ["#タグ", "タグ", "tags", "tag", "ライバーのタグ"]
const XLSX_CARD_COLOR_ORDER = ["R", "O", "Y", "G", "B", "I", "V"]
const XLSX_CARD_COLORS = {R: "R", O: "O", Y: "Y", G: "G", B: "B", I: "I", V: "V", 赤: "R", 橙: "O", 黄: "Y", 緑: "G", 青: "B", 藍: "I", 紫: "V"}

function xlsxCardCellValue(cell) {
    return cell && typeof cell === "object" && cell.formula === true ? cell.value : cell
}

function xlsxCardHeaderColumns(values) {
    const labels = values.map(value => String(xlsxCardCellValue(value) ?? "").normalize("NFKC").trim().toLowerCase())
    const columns = XLSX_CARD_HEADERS.map(label => {
        const aliases = XLSX_CARD_HEADER_ALIASES[label] || [label]
        return aliases.map(alias => labels.indexOf(alias.toLowerCase())).find(index => index >= 0) ?? -1
    })
    return columns.every(index => index >= 0) ? columns : null
}

function normalizeXlsxCardColors(cell, fail) {
    const value = xlsxCardCellValue(cell)
    if (cell?.formula === true && (value === null || value === undefined)) fail("色の数式の計算結果がありません。Excel で計算して保存してください。")
    if (value === null || value === undefined || (typeof value === "string" && !value.trim())) return []
    if (typeof value !== "string" || value.length > 200) fail("色は R・O・Y・G・B・I・V または赤・橙・黄・緑・青・藍・紫で入力してください。")
    const tokens = value.normalize("NFKC").toUpperCase().replace(/[\s,、・/|]+/g, "")
    const colors = new Set()
    for (const token of tokens) {
        if (!Object.hasOwn(XLSX_CARD_COLORS, token)) fail(`色欄の「${token}」に対応していません。7色の記号または色名を使ってください。`)
        colors.add(XLSX_CARD_COLORS[token])
    }
    return XLSX_CARD_COLOR_ORDER.filter(color => colors.has(color))
}

// rows: [{rowNumber: 1, values: ["仮ID", "カード名", ...]}, ...].
// Formula cells may use {formula: true, value: cachedValueOrNull}; no formula runs.
function normalizeXlsxCardTags(cell, fail) {
    const value = xlsxCardCellValue(cell)
    if (cell?.formula === true && (value === null || value === undefined)) fail("タグの数式の計算結果がありません。Excelで計算して保存してください。")
    if (value === null || value === undefined || value === "") return []
    if (typeof value !== "string") fail("タグは文字列で入力してください。")
    try {return normalizeCardTags(value.replace(/[#＃]+/gu, ","))} catch (error) {fail(error.message)}
}
function normalizeCardLibraryRows(rows, sourceName, sheetName) {
    if (!Array.isArray(rows) || rows.length > XLSX_CARD_LIMITS.rows) throw new Error("Excel の行数が読み込み上限を超えています。")
    const headerIndex = rows.findIndex(row => row.rowNumber <= 30 && Array.isArray(row.values) && xlsxCardHeaderColumns(row.values))
    if (headerIndex < 0) throw new Error(`「${sheetName}」に必要な列（ID・カード名・種類・クラス・レベル・Power）が見つかりません。`)
    const columns = xlsxCardHeaderColumns(rows[headerIndex].values)
    const effectColumn = rows[headerIndex].values.findIndex(value => ["効果", "効果テキスト", "効果本文"].includes(String(xlsxCardCellValue(value) ?? "").trim()))
    const colorColumn = rows[headerIndex].values.findIndex(value => XLSX_CARD_COLOR_HEADERS.includes(String(xlsxCardCellValue(value) ?? "").normalize("NFKC").trim().toLowerCase()))
    const tagColumn = rows[headerIndex].values.findIndex(value => XLSX_CARD_TAG_HEADERS.includes(String(xlsxCardCellValue(value) ?? "").normalize("NFKC").trim().toLowerCase()))
    const cards = [], keys = new Set()
    for (const row of rows.slice(headerIndex + 1)) {
        const rowNumber = row.rowNumber
        const fail = message => { throw new Error(`「${sheetName}」${rowNumber} 行目：${message}`) }
        if (!Array.isArray(row.values) || !Number.isInteger(rowNumber) || rowNumber < 1) fail("行の形式が正しくありません。")
        const cells = columns.map(index => row.values[index])
        const values = cells.map(xlsxCardCellValue)
        const empty = value => value === null || value === undefined || (typeof value === "string" && !value.trim())
        if (values.every(empty)) continue
        // The supplied workbook has a formula-only FILTER placeholder below
        // its card rows. A row without a card name and type is not card data.
        if (empty(values[1]) && empty(values[2]) && cells[0]?.formula === true && values.slice(1).every(empty)) continue
        cells.forEach((cell, index) => {
            if (cell?.formula === true && empty(cell.value)) fail(`${XLSX_CARD_HEADERS[index]} の数式の計算結果がありません。Excel で計算して保存してください。`)
        })
        const readText = (value, label, maximum, required = false) => {
            if (empty(value)) {
                if (required) fail(`${label}が空欄です。`)
                return ""
            }
            if (typeof value !== "string" || value.length > maximum) fail(`${label}の文字列が正しくありません。`)
            return value.trim()
        }
        const key = readText(values[0], "ID", 100, true)
        if (!/^[A-Za-z0-9_-]+$/.test(key)) fail("IDには半角英数字・ハイフン・アンダーバーを使ってください。")
        if (["__proto__", "prototype", "constructor"].includes(key)) fail("このIDは使用できません。別のIDに変更してください。")
        if (keys.has(key)) fail(`ID「${key}」が重複しています。`)
        const name = readText(values[1], "カード名", 200, true)
        const rawType = readText(values[2], "種類", 40, true).toUpperCase()
        const rawClass = readText(values[3], "クラス", 40).toUpperCase()
        if (!Object.hasOwn(XLSX_CARD_TYPES, rawType)) fail(`種類「${rawType}」に対応していません。`)
        if (rawClass && !Object.hasOwn(XLSX_CARD_CLASSES, rawClass)) fail(`クラス「${rawClass}」に対応していません。`)
        const number = (value, label, maximum) => {
            if (empty(value)) return null
            if (typeof value === "string" && /^\d+(?:\.0+)?$/.test(value.trim())) value = Number(value.trim())
            if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > maximum) fail(`${label}は 0 ～ ${maximum} の整数にしてください。`)
            return value
        }
        const cardType = XLSX_CARD_TYPES[rawType]
        const cost = number(values[4], "レベル", 999)
        const power = number(values[5], "Power", 999999)
        const effectCell = effectColumn < 0 ? null : row.values[effectColumn]
        const effectValue = xlsxCardCellValue(effectCell)
        if (effectCell?.formula === true && (effectValue === null || effectValue === undefined)) fail("効果の数式の計算結果がありません。Excel で計算して保存してください。")
        const text = effectValue ?? ""
        if (typeof text !== "string" || text.length > 10000) fail("効果は 10000 文字以内の文章にしてください。")
        const colors = cardType === "liver" && colorColumn >= 0 ? normalizeXlsxCardColors(row.values[colorColumn], fail) : []
        const tags = cardType === "liver" && tagColumn >= 0 ? normalizeXlsxCardTags(row.values[tagColumn], fail) : []
        keys.add(key)
        cards.push({key, name, text, cardType, className: rawClass ? XLSX_CARD_CLASSES[rawClass] : "", colors, tags, cost: cardType === "channel" ? null : cost, power: cardType === "liver" ? power : null, sourceRow: rowNumber})
        if (cards.length > XLSX_CARD_LIMITS.cards) fail(`カードは ${XLSX_CARD_LIMITS.cards} 件まで読み込めます。`)
    }
    if (!cards.length) throw new Error(`「${sheetName}」に読み込めるカードがありません。`)
    return {source: String(sourceName || "カードリスト.xlsx"), sourceSheet: sheetName, cards}
}

function xlsxCardCrc32(bytes) {
    let crc = 0xffffffff
    for (const byte of bytes) {
        crc ^= byte
        for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0)
    }
    return (crc ^ 0xffffffff) >>> 0
}

async function xlsxCardZip(buffer) {
    const bytes = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : ArrayBuffer.isView(buffer) ? new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength) : null
    if (!bytes || bytes.byteLength < 22) throw new Error("Excel ファイルの形式が正しくないか、途中で切れています。")
    if (bytes.byteLength > XLSX_CARD_LIMITS.fileBytes) throw new Error("Excel ファイルは 10 MB 以下にしてください。")
    if ([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((byte, index) => bytes[index] === byte)) throw new Error("パスワード付き Excel や .xls 形式には対応していません。パスワードなしの .xlsx ファイルを選んでください。")
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    const range = (offset, length) => {
        if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(length) || offset < 0 || length < 0 || offset + length > bytes.length) throw new Error("Excel ファイルが途中で切れています。")
    }
    let end = -1
    for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65557); offset--) {
        if (view.getUint32(offset, true) === 0x06054b50 && offset + 22 + view.getUint16(offset + 20, true) === bytes.length) { end = offset; break }
    }
    if (end < 0) throw new Error("Excel ファイルの ZIP 情報が見つかりません。.xlsx ファイルを選んでください。")
    const count = view.getUint16(end + 10, true), directorySize = view.getUint32(end + 12, true), directoryOffset = view.getUint32(end + 16, true)
    if (count === 0xffff || directorySize === 0xffffffff || directoryOffset === 0xffffffff) throw new Error("ZIP64 形式の Excel ファイルには対応していません。")
    if (view.getUint16(end + 4, true) || view.getUint16(end + 6, true) || view.getUint16(end + 8, true) !== count) throw new Error("分割 ZIP 形式には対応していません。")
    if (count > XLSX_CARD_LIMITS.entries) throw new Error("Excel ファイル内の項目数が読み込み上限を超えています。")
    range(directoryOffset, directorySize)
    if (directoryOffset + directorySize > end) throw new Error("Excel ファイルの ZIP 情報が壊れています。")
    const entries = new Map(), decoder = new TextDecoder("utf-8", {fatal: true})
    let offset = directoryOffset, totalSize = 0, actualTotal = 0
    for (let index = 0; index < count; index++) {
        range(offset, 46)
        if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("Excel ファイルの ZIP 情報が壊れています。")
        const flags = view.getUint16(offset + 8, true), method = view.getUint16(offset + 10, true), crc = view.getUint32(offset + 16, true)
        const compressedSize = view.getUint32(offset + 20, true), size = view.getUint32(offset + 24, true), nameLength = view.getUint16(offset + 28, true), extraLength = view.getUint16(offset + 30, true), commentLength = view.getUint16(offset + 32, true), disk = view.getUint16(offset + 34, true), localOffset = view.getUint32(offset + 42, true)
        range(offset + 46, nameLength + extraLength + commentLength)
        if (flags & 0x41) throw new Error("パスワード付き Excel ファイルには対応していません。")
        if (![0, 8].includes(method)) throw new Error("この Excel ファイルの圧縮方式には対応していません。")
        if (size === 0xffffffff || compressedSize === 0xffffffff || localOffset === 0xffffffff || disk === 0xffff) throw new Error("ZIP64 形式の Excel ファイルには対応していません。")
        if (disk) throw new Error("分割 ZIP 形式には対応していません。")
        totalSize += size
        if (size > XLSX_CARD_LIMITS.entryBytes || totalSize > XLSX_CARD_LIMITS.totalBytes) throw new Error("Excel ファイルの展開サイズが読み込み上限を超えています。")
        let name
        try { name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength)) } catch { throw new Error("Excel ファイル内の名前を読み取れません。") }
        if (!name || name.includes("\\") || name.includes("\0") || name.startsWith("/") || name.split("/").some(part => part === ".." || part === ".") || /^[A-Za-z]:/.test(name) || entries.has(name)) throw new Error("Excel ファイル内のパスが正しくありません。")
        range(localOffset, 30)
        if (localOffset >= directoryOffset || view.getUint32(localOffset, true) !== 0x04034b50 || view.getUint16(localOffset + 6, true) !== flags || view.getUint16(localOffset + 8, true) !== method) throw new Error("Excel ファイルの ZIP 情報が一致しません。")
        const localNameLength = view.getUint16(localOffset + 26, true), localExtraLength = view.getUint16(localOffset + 28, true), start = localOffset + 30 + localNameLength + localExtraLength
        range(localOffset + 30, localNameLength + localExtraLength)
        if (localNameLength !== nameLength || bytes.subarray(localOffset + 30, localOffset + 30 + localNameLength).some((value, n) => value !== bytes[offset + 46 + n])) throw new Error("Excel ファイル内の名前が一致しません。")
        range(start, compressedSize)
        if (start + compressedSize > directoryOffset) throw new Error("Excel ファイルの圧縮データが壊れています。")
        entries.set(name, {method, crc, compressedSize, size, start})
        offset += 46 + nameLength + extraLength + commentLength
    }
    if (offset !== directoryOffset + directorySize) throw new Error("Excel ファイルの ZIP 情報が壊れています。")
    const cache = new Map()
    return {
        has: name => entries.has(name),
        async read(name) {
            if (cache.has(name)) return cache.get(name)
            const entry = entries.get(name)
            if (!entry) throw new Error(`Excel ファイル内の「${name}」が見つかりません。`)
            let result
            if (entry.method === 0) result = bytes.subarray(entry.start, entry.start + entry.compressedSize)
            else {
                if (typeof DecompressionStream !== "function") throw new Error("このブラウザは Excel の展開に対応していません。最新版のブラウザで開いてください。")
                let reader
                try {
                    reader = new Blob([bytes.subarray(entry.start, entry.start + entry.compressedSize)]).stream().pipeThrough(new DecompressionStream("deflate-raw")).getReader()
                    const chunks = []
                    let length = 0
                    while (true) {
                        const chunk = await reader.read()
                        if (chunk.done) break
                        length += chunk.value.length
                        if (length > entry.size || length > XLSX_CARD_LIMITS.entryBytes || actualTotal + length > XLSX_CARD_LIMITS.totalBytes) {
                            await reader.cancel()
                            throw new Error("Excel ファイルの展開サイズが読み込み上限を超えています。")
                        }
                        chunks.push(chunk.value)
                    }
                    result = new Uint8Array(length)
                    let position = 0
                    for (const chunk of chunks) { result.set(chunk, position); position += chunk.length }
                } catch (error) {
                    if (error instanceof Error && error.message.startsWith("Excel")) throw error
                    throw new Error("Excel ファイルの圧縮データを展開できません。ファイルが壊れていないか確認してください。")
                } finally { reader?.releaseLock() }
            }
            if (result.length !== entry.size || xlsxCardCrc32(result) !== entry.crc) throw new Error("Excel ファイルの内容が壊れているか、途中で切れています。")
            actualTotal += result.length
            let text
            try { text = decoder.decode(result) } catch { throw new Error("Excel ファイル内の文字を読み取れません。") }
            cache.set(name, text)
            return text
        }
    }
}

function xlsxCardXml(text) {
    if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error("外部定義を含む XML には対応していません。")
    const xml = new DOMParser().parseFromString(text, "application/xml")
    if (xml.getElementsByTagNameNS("*", "parsererror").length || xml.documentElement?.localName === "parsererror") throw new Error("Excel ファイル内の XML が壊れています。")
    return xml
}

function xlsxCardResolvePath(baseFile, target) {
    if (!target || /[\\\0?#]/.test(target) || /^[a-z][a-z\d+.-]*:/i.test(target) || target.startsWith("//")) throw new Error("Excel の外部参照には対応していません。")
    const parts = target.startsWith("/") ? [] : baseFile.split("/").slice(0, -1)
    for (const part of target.split("/")) {
        if (!part || part === ".") continue
        if (part === "..") {
            if (!parts.length) throw new Error("Excel ファイル内の参照先が正しくありません。")
            parts.pop()
        } else parts.push(part)
    }
    return parts.join("/")
}

function xlsxCardRelationships(xml, baseFile) {
    const relationships = new Map()
    for (const relation of xml.getElementsByTagNameNS("*", "Relationship")) {
        if ((relation.getAttribute("TargetMode") || "").toLowerCase() === "external") throw new Error("外部ファイルを参照する Excel には対応していません。")
        const id = relation.getAttribute("Id"), type = relation.getAttribute("Type") || "", target = xlsxCardResolvePath(baseFile, relation.getAttribute("Target"))
        if (!id || relationships.has(id)) throw new Error("Excel ファイル内の参照情報が正しくありません。")
        relationships.set(id, {type, target})
    }
    return relationships
}

function xlsxCardString(node) {
    if (!node) return ""
    let text = ""
    for (const child of node.children) {
        if (child.localName === "t") text += child.textContent
        else if (child.localName === "r") for (const run of child.children) if (run.localName === "t") text += run.textContent
    }
    return text
}

function xlsxCardSheetRows(xml, strings) {
    const rows = []
    let cellsSeen = 0
    for (const node of xml.getElementsByTagNameNS("*", "row")) {
        const rowNumber = Number(node.getAttribute("r"))
        if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > 1048576) throw new Error("Excel の行番号が正しくありません。")
        const values = []
        for (const cell of node.children) {
            if (cell.localName !== "c") continue
            if (++cellsSeen > XLSX_CARD_LIMITS.cells) throw new Error("Excel のセル数が読み込み上限を超えています。")
            const reference = /^([A-Z]{1,3})([1-9]\d*)$/.exec(cell.getAttribute("r") || "")
            if (!reference || Number(reference[2]) !== rowNumber) throw new Error("Excel のセル参照が正しくありません。")
            let column = 0
            for (const character of reference[1]) column = column * 26 + character.charCodeAt(0) - 64
            if (column > 16384 || Object.hasOwn(values, column - 1)) throw new Error("Excel のセル参照が正しくありません。")
            const children = [...cell.children], raw = children.find(child => child.localName === "v")?.textContent ?? null, type = cell.getAttribute("t")
            let value = null
            if (type === "inlineStr") value = xlsxCardString(children.find(child => child.localName === "is"))
            else if (type === "s" && raw !== null) {
                const index = Number(raw)
                if (!Number.isInteger(index) || index < 0 || index >= strings.length) throw new Error("Excel の文字列参照が正しくありません。")
                value = strings[index]
            } else if (raw !== null) {
                if (!type || type === "n") value = raw.trim() ? Number(raw) : null
                else if (type === "b") value = raw === "1"
                else if (type === "e") value = {error: raw}
                else value = raw
            }
            values[column - 1] = children.some(child => child.localName === "f") ? {formula: true, value} : value
        }
        rows.push({rowNumber, values})
        if (rows.length > XLSX_CARD_LIMITS.rows) throw new Error("Excel の行数が読み込み上限を超えています。")
    }
    return rows
}

async function parseCardLibraryXlsx(buffer, sourceName) {
    const zip = await xlsxCardZip(buffer)
    const packageRelations = xlsxCardRelationships(xlsxCardXml(await zip.read("_rels/.rels")), "")
    const officeDocuments = [...packageRelations.values()].filter(relation => relation.type.endsWith("/officeDocument"))
    if (officeDocuments.length !== 1) throw new Error("Excel のブック情報が見つかりません。.xlsx ファイルを選んでください。")
    const workbookPath = officeDocuments[0].target, slash = workbookPath.lastIndexOf("/"), workbookRelationsPath = `${workbookPath.slice(0, slash + 1)}_rels/${workbookPath.slice(slash + 1)}.rels`
    const workbook = xlsxCardXml(await zip.read(workbookPath))
    const relations = xlsxCardRelationships(xlsxCardXml(await zip.read(workbookRelationsPath)), workbookPath)
    const sharedStringsRelation = [...relations.values()].find(relation => relation.type.endsWith("/sharedStrings"))
    const strings = sharedStringsRelation ? [...xlsxCardXml(await zip.read(sharedStringsRelation.target)).getElementsByTagNameNS("*", "si")].map(xlsxCardString) : []
    const sheets = [...workbook.getElementsByTagNameNS("*", "sheet")].map(sheet => {
        const name = sheet.getAttribute("name"), id = sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || sheet.getAttributeNS("http://purl.oclc.org/ooxml/officeDocument/relationships", "id") || sheet.getAttribute("r:id")
        return {name, relation: relations.get(id)}
    }).filter(sheet => sheet.relation?.type.endsWith("/worksheet"))
    if (!sheets.length) throw new Error("Excel に読み込めるシートがありません。")
    const canonical = sheets.find(sheet => sheet.name === "全カード")
    if (canonical) return normalizeCardLibraryRows(xlsxCardSheetRows(xlsxCardXml(await zip.read(canonical.relation.target)), strings), sourceName, canonical.name)
    for (const sheet of sheets) {
        const rows = xlsxCardSheetRows(xlsxCardXml(await zip.read(sheet.relation.target)), strings)
        if (rows.some(row => row.rowNumber <= 30 && xlsxCardHeaderColumns(row.values))) return normalizeCardLibraryRows(rows, sourceName, sheet.name)
    }
    throw new Error("カード一覧の列（ID・カード名・種類・クラス・レベル・Power）が見つかりません。")
}

let cardLibrary = {source: "", sourceSheet: "", cards: []}
const libraryUi = {query: "", cardType: "", className: "", page: 0, selected: new Set(), effectKey: ""}
const libraryPageSize = 30
let libraryByKey = new Map()
let libraryImporting = false
function setCardLibrary(library) {
    if (replay.mode === "replay") return false
    try {
        if (!library || typeof library.source !== "string" || !library.source || library.source.length > 200 || typeof library.sourceSheet !== "string" || library.sourceSheet.length > 100 || !Array.isArray(library.cards) || !library.cards.length || library.cards.length > 1000) throw new Error("カードリストの形式を確認してください。")
        const seen = new Set()
        const cards = library.cards.map(card => {
            if (!card || typeof card.key !== "string" || !/^[A-Za-z0-9_-]{1,100}$/.test(card.key) || ["__proto__", "constructor", "prototype"].includes(card.key) || seen.has(card.key)) throw new Error("カードのIDが不正か重複しています。")
            seen.add(card.key)
            if (typeof card.name !== "string" || !card.name.trim() || card.name.length > 200 || !["liver", "event", "stage", "channel"].includes(card.cardType) || !["", ...textCardClasses].includes(card.className)) throw new Error(`${card.key}の名前・種別・クラスを確認してください。`)
            if (card.text !== undefined && (typeof card.text !== "string" || card.text.length > 10000)) throw new Error(`${card.key}の効果は10000文字以内の文章にしてください。`)
            if (card.colors !== undefined && (!Array.isArray(card.colors) || card.colors.length > nijiColors.length || card.colors.some(color => !nijiColors.includes(color)) || new Set(card.colors).size !== card.colors.length)) throw new Error(`${card.key}の色を確認してください。`)
            if (card.tags !== undefined && !validCardTags(card.tags)) throw new Error(`${card.key}のタグを確認してください。`)
            for (const [field, max] of [["cost", 999], ["power", 999999]]) {
                if (card[field] !== null && (!Number.isSafeInteger(card[field]) || card[field] < 0 || card[field] > max)) throw new Error(`${card.key}の数値を確認してください。`)
            }
            if (!Number.isSafeInteger(card.sourceRow) || card.sourceRow < 1 || card.sourceRow > 1048576) throw new Error(`${card.key}の行番号を確認してください。`)
            return {key: card.key, name: card.name.trim(), text: card.text ?? "", cardType: card.cardType, className: card.className,
                cost: card.cardType === "channel" ? null : card.cost, power: card.cardType === "liver" ? card.power : null,
                colors: card.cardType === "liver" ? nijiColors.filter(color => card.colors?.includes(color)) : [], tags: card.cardType === "liver" ? [...(card.tags || [])] : [], sourceRow: card.sourceRow}
        })
        cardLibrary = {source: library.source, sourceSheet: library.sourceSheet, cards}
        libraryByKey = new Map(cards.map(card => [card.key, card]))
        return true
    } catch (error) {notify(error.message, true); return false}
}
async function readCardLibraryFile(file) {
    if (!file || libraryImporting || replay.mode === "replay") return false
    if (!/\.xlsx$/i.test(file.name) || file.size > 10 * 1024 * 1024) {notify("10MB以内の.xlsxファイルを選んでください。", true); return false}
    libraryImporting = true
    notify("Excelのカード情報を読み取っています。")
    try {
        const library = await parseCardLibraryXlsx(await file.arrayBuffer(), file.name)
        if (!setCardLibrary(library)) return false
        showCardLibrary(true)
        return true
    } catch (error) {notify(`Excelを読み込めませんでした。${error.message}`, true); return false}
    finally {libraryImporting = false}
}
function registeredLibraryKeys() {
    return new Set(Object.values(state.catalog).map(item => item.sourceCardKey).filter(Boolean))
}
function registeredLibraryDefinitions() {
    const definitions = new Map()
    for (const [id, item] of Object.entries(state.catalog)) {
        if (!item.sourceCardKey || isColorDefinition(item)) continue
        if (!definitions.has(item.sourceCardKey)) definitions.set(item.sourceCardKey, [])
        definitions.get(item.sourceCardKey).push({id, item})
    }
    return definitions
}
function libraryCardRegistration(card, registrations = registeredLibraryDefinitions()) {
    const definitions = registrations.get(card.key) || []
    const effectDefinitions = card.text?.trim() ? definitions.filter(({item}) => !item.text?.trim()) : []
    const colorDefinitions = card.cardType === "liver" && card.colors?.length
        ? definitions.filter(({item}) => textCardType(item) === "liver" && !item.colors?.length) : []
    const tagDefinitions = card.cardType === "liver" && card.tags?.length
        ? definitions.filter(({item}) => textCardType(item) === "liver" && !item.tags?.length) : []
    const additions = [["effect", effectDefinitions], ["color", colorDefinitions], ["tag", tagDefinitions]].filter(([, targets]) => targets.length).map(([name]) => name)
    const mode = !definitions.length ? "new" : !additions.length ? "registered" : additions.length === 1 ? additions[0] : "details"
    const targets = new Set([...effectDefinitions, ...colorDefinitions, ...tagDefinitions])
    return {mode, additions, definitions: mode === "registered" ? definitions : definitions.filter(item => targets.has(item)), effectDefinitions, colorDefinitions, tagDefinitions}
}
function isLibraryCardEligible(card, registrations) {
    return libraryCardRegistration(card, registrations).mode !== "registered"
}
function librarySelectionCounts(registrations = registeredLibraryDefinitions()) {
    const counts = {new: 0, effect: 0, color: 0, tag: 0}
    for (const key of libraryUi.selected) {
        const card = libraryByKey.get(key)
        if (!card) continue
        const {mode, effectDefinitions, colorDefinitions, tagDefinitions} = libraryCardRegistration(card, registrations)
        if (mode === "new") counts.new++
        if (effectDefinitions.length) counts.effect++
        if (colorDefinitions.length) counts.color++
        if (tagDefinitions.length) counts.tag++
    }
    return counts
}
function libraryColorTags(card) {
    const colors = liverCardColors(card)
    return colors.length ? `<span class="library-color-tags">${colors.map(color => `<span class="library-color-tag"><span class="card-color-swatch color-${color.toLowerCase()}" aria-hidden="true"></span>${colorNames[color]}</span>`).join("")}</span>` : "—"
}
function filteredLibraryCards() {
    const normalize = value => String(value).normalize("NFKC").toLocaleLowerCase("ja")
    const terms = normalize(libraryUi.query).trim().split(/\s+/).filter(Boolean)
    return cardLibrary.cards.filter(card => (!libraryUi.cardType || card.cardType === libraryUi.cardType)
        && (!libraryUi.className || card.className === libraryUi.className)
        && terms.every(term => normalize(`${card.name} ${card.key} ${liverCardTags(card).map(tag => `#${tag}`).join(" ")} ${card.text || ""}`).includes(term)))
}
function libraryCardSummary(item) {
    const parts = [item.sourceCardKey, textCardTypes[textCardType(item)], item.className]
    if (textCardType(item) !== "channel" && item.cost !== undefined && item.cost !== null) parts.push(`レベル ${item.cost}`)
    if (textCardType(item) === "liver" && item.power !== undefined && item.power !== null) parts.push(`パワー ${item.power}`)
    return parts.filter(Boolean).join(" / ")
}
function showCardLibrary(selectAll = false) {
    if (replay.mode === "replay") return false
    const registrations = registeredLibraryDefinitions()
    Object.assign(libraryUi, {query: "", cardType: "", className: "", page: 0, effectKey: "", selected: new Set(selectAll ? cardLibrary.cards.filter(card => isLibraryCardEligible(card, registrations)).map(card => card.key) : [])})
    if (!cardLibrary.cards.length) {
        openModal("Excelからカードを取り込む", `<div class="library-entry"><h3>カードリストのExcelを選択</h3><p>カード名・種別・クラス・色・レベル・パワー・タグ・効果を読み取り、一覧で確認してから登録できます。</p><button class="primary" data-action="pick-card-library">Excelファイルを選択</button></div><div class="modal-footer"><button data-action="catalog">カード登録へ戻る</button></div>`, "card-library")
        return true
    }
    const types = Object.entries(textCardTypes).filter(([value]) => value !== "other").map(([value, label]) => `<option value="${value}">${label}</option>`).join("")
    const classes = textCardClasses.map(value => `<option value="${value}">${value}</option>`).join("")
    openModal("取り込み内容を確認", `<div class="card-library"><div class="modal-toolbar"><strong>${escapeHTML(cardLibrary.source)}</strong><span>${cardLibrary.cards.length}種類</span><button data-action="pick-card-library">別のExcelを選択</button></div><p class="help-note">使うカードを選んで登録してください。登録済みのカードには、未入力の効果・色・タグを追加できます。設定済みの内容は上書きしません。ExcelのIDを使います。</p><div class="library-filters"><label>カード名・ID・タグ・効果<input id="librarySearch" type="search" maxlength="200" placeholder="カード名・ID・タグ・効果で検索"></label><label>カード種別<select id="libraryType"><option value="">すべて</option>${types}</select></label><label>クラス<select id="libraryClass"><option value="">すべて</option>${classes}</select></label></div><div id="libraryResults"></div></div>`, "card-library")
    renderLibraryResults()
    return true
}
function renderLibraryResults(preserveScroll = false) {
    if (modalKind !== "card-library" || !cardLibrary.cards.length) return
    const previousScroll = preserveScroll ? document.querySelector(".library-list")?.scrollTop || 0 : 0
    const registrations = registeredLibraryDefinitions()
    for (const key of libraryUi.selected) if (!libraryByKey.has(key) || !isLibraryCardEligible(libraryByKey.get(key), registrations)) libraryUi.selected.delete(key)
    const cards = filteredLibraryCards()
    const pages = Math.max(1, Math.ceil(cards.length / libraryPageSize))
    libraryUi.page = integer(libraryUi.page, 0, pages - 1)
    const visible = cards.slice(libraryUi.page * libraryPageSize, (libraryUi.page + 1) * libraryPageSize)
    const eligible = visible.some(card => isLibraryCardEligible(card, registrations))
    const counts = librarySelectionCounts(registrations)
    const countSummary = [counts.new ? `新規${counts.new}種類` : "", counts.effect ? `効果追加${counts.effect}種類` : "", counts.color ? `色追加${counts.color}種類` : "", counts.tag ? `タグ追加${counts.tag}種類` : ""].filter(Boolean).join("・") || "0種類"
    const registerLabel = counts.effect || counts.color || counts.tag ? [counts.new ? `新規${counts.new}種類を登録` : "", counts.effect ? `${counts.effect}種類に効果を追加` : "", counts.color ? `${counts.color}種類に色を追加` : "", counts.tag ? `${counts.tag}種類にタグを追加` : ""].filter(Boolean).join("・") : `選択した${counts.new}種類を登録`
    const rows = visible.map(card => {
        const {mode, additions} = libraryCardRegistration(card, registrations)
        const added = mode === "registered", inputId = `librarySelect-${card.key}`
        const status = added ? "登録済み" : mode === "new" ? "新規登録" : `登録済み · ${additions.map(field => ({effect: "効果", color: "色", tag: "タグ"})[field]).join("・")}を追加`
        const effectButton = card.text.trim() || liverCardTags(card).length ? `<button type="button" class="small library-effect-button" data-action="library-effect" data-library-key="${escapeHTML(card.key)}" aria-expanded="${libraryUi.effectKey === card.key}" aria-controls="libraryEffectDetails">${libraryUi.effectKey === card.key ? "詳細を閉じる" : "詳細を見る"}</button>` : ""
        return `<div class="library-row${added ? " is-registered" : mode !== "new" ? " needs-effect" : ""}"><input id="${escapeHTML(inputId)}" type="checkbox" data-library-key="${escapeHTML(card.key)}" aria-label="${escapeHTML(card.name)} (${escapeHTML(card.key)})を選択" ${added ? "disabled checked" : libraryUi.selected.has(card.key) ? "checked" : ""}><div class="library-name"><label for="${escapeHTML(inputId)}"><strong>${escapeHTML(card.name)}</strong><small>${escapeHTML(card.key)} · ${status}</small></label>${effectButton}</div><span class="library-type" data-label="種別">${escapeHTML(textCardTypes[card.cardType])}</span><span class="library-class" data-label="クラス">${escapeHTML(card.className || "—")}</span><span class="library-colors" data-label="色">${libraryColorTags(card)}</span><span class="library-cost" data-label="レベル">${card.cardType === "channel" ? "なし" : card.cost ?? "—"}</span><span class="library-power" data-label="パワー">${card.cardType === "liver" ? card.power ?? "—" : "—"}</span></div>`
    }).join("")
    const effectCard = visible.find(card => card.key === libraryUi.effectKey)
    if (!effectCard) libraryUi.effectKey = ""
    const details = effectCard ? `<section id="libraryEffectDetails" class="library-effect-details"><h3>${escapeHTML(effectCard.name)} <small>${escapeHTML(effectCard.key)}</small></h3>${cardTagsHTML(effectCard)}<p class="description" style="white-space:pre-wrap">${escapeHTML(effectCard.text)}</p></section>` : `<div id="libraryEffectDetails"></div>`
    byId("libraryResults").innerHTML = `<div class="library-summary" role="status"><span>${cards.length}件${cards.length ? `中 ${libraryUi.page * libraryPageSize + 1}〜${libraryUi.page * libraryPageSize + visible.length}件` : ""}</span><span>${countSummary}を選択</span></div><div class="library-list"><div class="library-list-head" aria-hidden="true"><span></span><span>カード名 / ID</span><span>種別</span><span>クラス</span><span>色</span><span>レベル</span><span>パワー</span></div>${rows || `<p class="library-empty">条件に合うカードがありません。</p>`}</div>${details}<div class="library-pagination"><button data-action="library-page" data-delta="-1" ${libraryUi.page === 0 ? "disabled" : ""}>前のページ</button><span>${libraryUi.page + 1} / ${pages}</span><button data-action="library-page" data-delta="1" ${libraryUi.page + 1 >= pages ? "disabled" : ""}>次のページ</button></div><div class="modal-footer"><button data-action="catalog">カード登録へ戻る</button><button data-action="library-select-page" ${eligible ? "" : "disabled"}>表示中を選択</button><button data-action="library-clear" ${libraryUi.selected.size ? "" : "disabled"}>選択解除</button><button data-action="library-register" class="primary" ${libraryUi.selected.size ? "" : "disabled"}>${registerLabel}</button></div>`
    const list = document.querySelector(".library-list")
    if (list) list.scrollTop = previousScroll
}
function selectLibraryCard(key, selected) {
    if (replay.mode === "replay" || !libraryByKey.has(key) || !isLibraryCardEligible(libraryByKey.get(key))) return false
    if (selected) libraryUi.selected.add(key)
    else libraryUi.selected.delete(key)
    renderLibraryResults(true)
    return true
}
function toggleLibraryEffect(key) {
    if (replay.mode === "replay" || (!libraryByKey.get(key)?.text.trim() && !liverCardTags(libraryByKey.get(key)).length)) return false
    libraryUi.effectKey = libraryUi.effectKey === key ? "" : key
    renderLibraryResults(true)
    return true
}
function registerLibraryCards(keys = [...libraryUi.selected]) {
    if (replay.mode === "replay") return false
    if (!Array.isArray(keys) || keys.some(key => !libraryByKey.has(key))) {
        notify("カードリストから追加するカードを選んでください。", true)
        return false
    }
    const registrations = registeredLibraryDefinitions()
    const cards = [...new Set(keys)].map(key => libraryByKey.get(key)).map(card => ({card, ...libraryCardRegistration(card, registrations)})).filter(item => item.mode !== "registered")
    if (!cards.length) {notify("新規登録または効果・色・タグを追加するカードを選んでください。設定済みの内容は上書きしません。"); return 0}
    const newCards = cards.filter(item => item.mode === "new"), effectCards = cards.filter(item => item.effectDefinitions.length), colorCards = cards.filter(item => item.colorDefinitions.length), tagCards = cards.filter(item => item.tagDefinitions.length)
    const currentCount = Object.values(state.catalog).filter(item => !isColorDefinition(item)).length
    if (newCards.length && currentCount + newCards.length > 500) {
        notify(`登録は500種類までです。追加できるのはあと${500 - currentCount}種類です。`, true)
        return false
    }
    const changes = [newCards.length ? `${newCards.length}種類を登録` : "", effectCards.length ? `${effectCards.length}種類に効果を追加` : "", colorCards.length ? `${colorCards.length}種類に色を追加` : "", tagCards.length ? `${tagCards.length}種類にタグを追加` : ""].filter(Boolean).join("・")
    const added = transact(`カードリストから${changes}`, () => {
        for (const {card: item} of newCards) {
            state.catalog[uid("definition")] = {name: item.name, text: item.text, asset: "", cardType: item.cardType,
                className: item.className, cost: item.cardType === "channel" ? null : item.cost,
                power: item.cardType === "liver" ? item.power : null, colors: [...item.colors], tags: [...item.tags], sourceCardKey: item.key}
        }
        for (const {card, effectDefinitions} of effectCards) {
            for (const {id} of effectDefinitions) state.catalog[id].text = card.text
        }
        for (const {card, colorDefinitions} of colorCards) {
            for (const {id} of colorDefinitions) state.catalog[id].colors = [...card.colors]
        }
        for (const {card, tagDefinitions} of tagCards) {
            for (const {id} of tagDefinitions) state.catalog[id].tags = [...card.tags]
        }
    })
    if (!added) return false
    libraryUi.selected.clear()
    showCatalog()
    notify(`${changes}しました。${newCards.length ? "「デッキ編集」で使えます。" : ""}`)
    return cards.length
}

// CARD_LIBRARY_END
const colorNames = {R: "赤", O: "橙", Y: "黄", G: "緑", B: "青", I: "藍", V: "紫"}
const phases = ["スタート", "メイン", "アクション", "セット", "エンド"]
const defaultZones = () => [
    {id: "deck", label: "山札", mode: "pile", hidden: true},
    {id: "field", label: "フィールド", mode: "row", hidden: false},
    {id: "set", label: "セット", mode: "pile", hidden: true},
    {id: "channel", label: "チャンネル", mode: "row", hidden: false},
    {id: "stage", label: "ステージ", mode: "row", hidden: false},
    {id: "discard", label: "ログ", mode: "pile", hidden: false},
    {id: "colorline", label: "カラーライン", mode: "row", hidden: true},
    ...nijiColors.map(color => ({id: `niji_${color.toLowerCase()}`, label: `にじ ${color}`, mode: "row", hidden: false})),
    {id: "hand", label: "手札", mode: "hand", hidden: false}
]
// Upgrade earlier saves to the current playmat and single-card slots.
function upgradeLayout(candidate) {
    if (candidate.layoutVersion !== 2) {
        for (const zone of defaultZones()) {
            if (candidate.zones.some(existing => existing.id === zone.id)) continue
            candidate.zones.push(zone)
            for (const player of playerIds) candidate.locations[player][zone.id] = []
        }
        const discard = candidate.zones.find(zone => zone.id === "discard")
        if (["捨札（仮）", "捨札"].includes(discard.label)) discard.label = "ログ"
        candidate.phase = 0
        candidate.layoutVersion = 2
    }
    const restoreCard = (id, player) => {
        const card = candidate.cards[id]
        const color = isColorDefinition(candidate.catalog[card.definition])
        candidate.locations[color ? card.colorOwner : player][color ? "colorline" : "discard"].push(id)
        card.faceDown = color
        card.rotated = false
    }
    if (candidate.rulesVersion !== 1) {
        const standardIds = new Set(defaultZones().map(zone => zone.id))
        for (const zone of candidate.zones.filter(zone => !standardIds.has(zone.id))) {
            for (const player of playerIds) {
                for (const id of candidate.locations[player][zone.id]) restoreCard(id, player)
                delete candidate.locations[player][zone.id]
            }
        }
        candidate.zones = candidate.zones.filter(zone => standardIds.has(zone.id))
    }
    if (candidate.colorSetupVersion !== 1) {
        for (const player of playerIds) {
            const previousCards = candidate.locations[player].colorline.filter(id => !isColorDefinition(candidate.catalog[candidate.cards[id].definition]))
            candidate.locations[player].colorline = candidate.locations[player].colorline.filter(id => isColorDefinition(candidate.catalog[candidate.cards[id].definition]))
            for (const id of previousCards) restoreCard(id, player)
        }
    }
    initializeColorCards(candidate)
    if (candidate.rulesVersion !== 1) {
        for (const player of playerIds) {
            for (const zone of ["set", "stage", "channel"]) {
                const ids = candidate.locations[player][zone]
                const regularIds = ids.filter(id => !isColorDefinition(candidate.catalog[candidate.cards[id].definition]))
                const retained = regularIds.at(-1)
                candidate.locations[player][zone] = retained ? [retained] : []
                for (const id of ids) if (id !== retained) restoreCard(id, player)
            }
        }
        candidate.rulesVersion = 1
    }
    initializeChannelConfiguration(candidate)
    for (const player of playerIds) for (const id of candidate.locations[player].hand) candidate.cards[id].rotated = false
    return candidate
}
function isColorDefinition(item) {return item?.kind === "color"}
function initializeChannelConfiguration(candidate) {
    if (!candidate.channelDefinitions) candidate.channelDefinitions = Object.fromEntries(playerIds.map(player => [player, candidate.cards[candidate.locations[player].channel[0]]?.definition || ""]))
    for (const player of playerIds) candidate.decklists[player] = candidate.decklists[player].filter(item => item.definition !== candidate.channelDefinitions[player])
    return candidate
}
function colorForCard(id) {
    const card = state.cards[id]
    const item = card && state.catalog[card.definition]
    return isColorDefinition(item) ? item.color : null
}
function initializeColorCards(candidate) {
    for (const color of nijiColors) {
        let definitionId = Object.keys(candidate.catalog).find(id => isColorDefinition(candidate.catalog[id]) && candidate.catalog[id].color === color)
        if (!definitionId) {
            definitionId = uid("color_definition")
            candidate.catalog[definitionId] = {name: `${colorNames[color]}のカラーカード`, text: "カラーラインでは裏向きで開始します。表裏の変更とにじエリアへの移動を手動で操作できます。", asset: "", kind: "color", color}
        }
        for (const player of playerIds) {
            const exists = Object.values(candidate.cards).some(card => card.colorOwner === player && isColorDefinition(candidate.catalog[card.definition]) && candidate.catalog[card.definition].color === color)
            if (exists) continue
            const id = uid("color_card")
            candidate.cards[id] = {definition: definitionId, colorOwner: player, faceDown: true, rotated: false, counter: 0, note: ""}
            candidate.locations[player].colorline.push(id)
        }
    }
    candidate.colorSetupVersion = 1
    return candidate
}
function resetColorCards(player) {
    if (!playerIds.includes(player)) throw new Error("Invalid color card owner.")
    initializeColorCards(state)
    const ids = nijiColors.map(color => Object.keys(state.cards).find(id => state.cards[id].colorOwner === player && colorForCard(id) === color))
    moveRaw(ids, player, "colorline", "first")
    for (const id of ids) Object.assign(state.cards[id], {faceDown: true, rotated: false, counter: 0, note: ""})
}
function blankState() {
    const zones = defaultZones()
    const players = {}
    const locations = {}
    const decklists = {}
    for (const [index, player] of playerIds.entries()) {
        players[player] = {name: `プレイヤー ${index + 1}`, meters: [{label: "カウンター A", value: 0}, {label: "カウンター B", value: 0}]}
        locations[player] = Object.fromEntries(zones.map(zone => [zone.id, []]))
        decklists[player] = []
    }
    return initializeColorCards({version: 1, layoutVersion: 2, rulesVersion: 1, channelDefinitions: {p1: "", p2: ""}, phase: 0, showOpponent: false, active: "p1", turn: 1, zones, players, locations, cards: {}, catalog: {}, decklists, log: []})
}
let state = blankState()
let assets = {}
let undoStack = []
let redoStack = []
let database = null
let saveTimer = null
let saveVersion = 0
let toastTimer = null
let modalKind = ""
let inspectContext = null
let previewCardId = null
let previewReturnContext = null
let editorPlayer = "p1"
const ui = {selection: new Set(), tab: "operation", drawCount: 1, peekCount: 3, discardCount: 1, initialCount: 5, bothHands: false, movePlayer: "p1", moveZone: "field", deckMoveZone: "field", position: "last", dragging: false, touched: false}
// REPLAY_START
const REPLAY_MAX_OPERATIONS = 200
const REPLAY_MAX_BYTES = 100000000
const replay = {data: null, recording: false, mode: "manual", playing: false, index: 0, stepSeconds: 1, timer: null, manual: null, epoch: 0, dismissedCutInIndex: -1, cutInKey: "", cutInEditor: null}

function replayInspectionView() {
    if (!byId("modal").open) return null
    if (modalKind === "zone" && inspectContext?.zone === "deck") {
        const {player, limit} = inspectContext
        return {player, kind: limit === null ? "all" : "top", cards: [...inspectorIds()]}
    }
    if (modalKind === "preview" && previewCardId) {
        const location = where(previewCardId)
        if (location?.zone === "deck") return {player: location.player, kind: "card", cards: [previewCardId]}
    }
    return null
}

function replayFrame(label, delay) {
    const snapshot = copy(state)
    snapshot.log = []
    return {label, delay, state: snapshot, view: {bothHands: ui.bothHands, inspection: replayInspectionView()}}
}
function collectReplayAssets(target) {
    for (const item of Object.values(state.catalog)) if (item.asset) target[item.asset] = assets[item.asset]
}
function validateReplayRecord(record) {
    const object = value => value !== null && typeof value === "object" && !Array.isArray(value)
    if (!object(record) || record.format !== "nijinana-replay" || record.formatVersion !== 1) throw new Error("対応していないリプレイ形式です。")
    if (typeof record.createdAt !== "string" || record.createdAt.length > 40 || !Number.isFinite(Date.parse(record.createdAt))) throw new Error("リプレイの日時が不正です。")
    if (!Array.isArray(record.frames) || !record.frames.length || record.frames.length > REPLAY_MAX_OPERATIONS + 1) throw new Error("リプレイの操作数が不正です。")
    if (new Blob([JSON.stringify(record)]).size > REPLAY_MAX_BYTES) throw new Error("リプレイは100MBまでです。")
    for (const [index, frame] of record.frames.entries()) {
        if (!object(frame) || typeof frame.label !== "string" || frame.label.length > 300 || !Number.isInteger(frame.delay) || frame.delay < 0 || frame.delay > 60000 || (index === 0 && frame.delay !== 0)) throw new Error("リプレイの操作が不正です。")
        if (!object(frame.view) || typeof frame.view.bothHands !== "boolean") throw new Error("リプレイの表示設定が不正です。")
        const cutIn = frame.view.cutIn
        if (cutIn !== undefined && cutIn !== null && (!object(cutIn) || index === 0 || typeof cutIn.text !== "string" || !cutIn.text.length || cutIn.text.length > 300 || cutIn.text !== cutIn.text.trim())) throw new Error("リプレイのカットインが不正です。")
        if (frame.state?.layoutVersion !== 2 || frame.state?.rulesVersion !== 1 || frame.state?.colorSetupVersion !== 1 || !frame.state?.channelDefinitions) throw new Error("対応していない盤面形式です。")
        checkState(frame.state, record.assets, index === 0)
        const inspection = frame.view.inspection
        if (inspection !== undefined && inspection !== null) {
            if (!object(inspection) || !playerIds.includes(inspection.player) || !["top", "all", "card"].includes(inspection.kind) || !Array.isArray(inspection.cards)) throw new Error("リプレイの確認カードが不正です。")
            const deck = frame.state.locations[inspection.player].deck
            if (inspection.cards.length > deck.length || (inspection.kind === "card" && inspection.cards.length !== 1) || new Set(inspection.cards).size !== inspection.cards.length || inspection.cards.some(id => typeof id !== "string" || !Object.hasOwn(frame.state.cards, id) || !deck.includes(id))) throw new Error("リプレイの確認カードが不正です。")
        }
    }
    return true
}
function startReplayRecording() {
    if (replay.mode !== "manual" || replay.recording) return false
    try {
        closeModal(false)
        const data = {format: "nijinana-replay", formatVersion: 1, createdAt: new Date().toISOString(), assets: {}, frames: [replayFrame("開始時の盤面", 0)]}
        collectReplayAssets(data.assets)
        validateRecord(buildRecord(data))
        replay.data = data
        replay.recording = true
        replay.index = 0
        replay.epoch += 1
        ui.touched = true
        renderReplayControls()
        scheduleSave()
        notify("記録を開始しました。盤面を操作したら「記録終了」を押してください。")
        return true
    } catch (error) {notify(`記録を開始できませんでした。${error.message}`, true); return false}
}
function captureReplayFrame(label, cutIn = null) {
    if (!replay.recording || replay.mode !== "manual") return false
    try {
        const frame = replayFrame(label, 1500)
        if (cutIn) frame.view.cutIn = copy(cutIn)
        const previous = replay.data.frames.at(-1)
        const comparable = entry => [entry.state, {bothHands: entry.view.bothHands, inspection: entry.view.inspection || null}]
        if (!cutIn && JSON.stringify(comparable(frame)) === JSON.stringify(comparable(previous))) return false
        const sharedAssets = {...replay.data.assets}
        collectReplayAssets(sharedAssets)
        const next = {...replay.data, assets: sharedAssets, frames: [...replay.data.frames, frame]}
        if (new Blob([JSON.stringify(buildRecord(next))]).size > REPLAY_MAX_BYTES) {
            stopReplayRecording()
            notify("容量の上限に達したため記録を終了しました。直前の操作は記録に含みません。", true)
            return false
        }
        replay.data = next
        if (next.frames.length >= REPLAY_MAX_OPERATIONS + 1) {
            stopReplayRecording()
            notify("200操作を記録したため、記録を終了しました。")
        }
        renderReplayControls()
        scheduleSave()
        return true
    } catch (error) {
        stopReplayRecording()
        notify(`記録を終了しました。${error.message}`, true)
        return false
    }
}
function stopReplayRecording() {
    if (!replay.recording) return false
    replay.recording = false
    replay.epoch += 1
    renderReplayControls()
    scheduleSave()
    return true
}
function pauseReplay() {
    clearTimeout(replay.timer)
    replay.timer = null
    replay.playing = false
    renderReplayControls()
}
function showReplayFrame(index) {
    replay.index = Math.max(0, Math.min(replay.data.frames.length - 1, Math.trunc(Number(index)) || 0))
    replay.dismissedCutInIndex = -1
    const frame = replay.data.frames[replay.index]
    state = copy(frame.state)
    assets = replay.data.assets
    ui.bothHands = frame.view.bothHands
    ui.selection.clear()
    ui.movePlayer = state.active
    render()
}
function enterReplay() {
    if (!replay.data || replay.recording || replay.mode !== "manual") return false
    clearTimeout(saveTimer)
    replay.manual = {state, assets, undoStack, redoStack, editorPlayer, ui: {...ui, selection: new Set(ui.selection)}}
    replay.mode = "replay"
    replay.epoch += 1
    replay.playing = false
    ui.touched = true
    ui.dragging = false
    closeModal()
    showReplayFrame(0)
    return true
}
function scheduleReplayStep() {
    clearTimeout(replay.timer)
    if (!replay.playing || replay.mode !== "replay") return
    if (replay.index >= replay.data.frames.length - 1) {pauseReplay(); return}
    const delay = replay.stepSeconds * 1000
    replay.timer = setTimeout(() => {
        if (!replay.playing || replay.mode !== "replay") return
        showReplayFrame(replay.index + 1)
        scheduleReplayStep()
    }, delay)
}
function toggleReplayPlayback() {
    if (replay.mode !== "replay" || replay.data.frames.length < 2) return
    if (replay.playing) {pauseReplay(); return}
    if (replay.index === replay.data.frames.length - 1) showReplayFrame(0)
    replay.playing = true
    renderReplayControls()
    scheduleReplayStep()
}
function seekReplay(index) {
    if (replay.mode !== "replay") return
    pauseReplay()
    showReplayFrame(index)
}
function setReplayStepSeconds(value) {
    const seconds = Number(value)
    if (![0.5, 1, 1.5, 2, 3, 5].includes(seconds)) return
    replay.stepSeconds = seconds
    renderReplayControls()
    if (replay.playing) scheduleReplayStep()
}
function leaveReplay() {
    if (replay.mode !== "replay") return false
    pauseReplay()
    const manual = replay.manual
    state = manual.state
    assets = manual.assets
    undoStack = manual.undoStack
    redoStack = manual.redoStack
    editorPlayer = manual.editorPlayer
    Object.assign(ui, manual.ui)
    ui.selection = new Set(manual.ui.selection)
    ui.touched = true
    replay.mode = "manual"
    replay.manual = null
    replay.epoch += 1
    closeModal()
    render()
    scheduleSave()
    return true
}
function buildReplayRecord() {
    if (!replay.data) throw new Error("記録したリプレイがありません。")
    return copy(replay.data)
}
function saveReplay() {
    try {
        const record = buildReplayRecord()
        validateReplayRecord(record)
        const blob = new Blob([JSON.stringify(record)], {type: "application/json"})
        const url = URL.createObjectURL(blob)
        const anchor = document.createElement("a")
        anchor.href = url
        anchor.download = `nijinana_replay_${new Date().toISOString().slice(0, 19).replaceAll(":", "-")}.json`
        anchor.click()
        setTimeout(() => URL.revokeObjectURL(url), 15000)
        notify("リプレイを保存しました。「リプレイ読込」で再生できます。")
    } catch (error) {notify(`保存できませんでした。${error.message}`, true)}
}
async function readReplayFile(file) {
    if (!file) return false
    if (replay.recording || replay.mode !== "manual") {notify("記録・再生を終了してから読み込んでください。"); return false}
    const epoch = replay.epoch
    try {
        if (file.size > REPLAY_MAX_BYTES) throw new Error("リプレイは100MBまでです。")
        const record = JSON.parse(await file.text())
        validateReplayRecord(record)
        if (epoch !== replay.epoch || replay.recording || replay.mode !== "manual") throw new Error("操作状態が変わったため、読み込みを中止しました。")
        validateRecord(buildRecord(record))
        replay.data = copy(record)
        replay.index = 0
        replay.epoch += 1
        ui.touched = true
        renderReplayControls()
        scheduleSave()
        notify("リプレイを読み込みました。「再生」で確認できます。")
        return true
    } catch (error) {notify(`読み込めませんでした。${error.message}`, true); return false}
}
function restoreReplayFromBoard(record) {
    replay.data = record.replay ? copy({...record.replay, assets: record.assets}) : null
    replay.recording = false
    replay.index = 0
    replay.epoch += 1
}

function dismissReplayCutIn() {
    if (replay.mode !== "replay") return
    replay.dismissedCutInIndex = replay.index
    renderReplayCutIn()
}
function cutInText(value) {
    if (typeof value !== "string") throw new Error("文章を入力してください。")
    const text = value.trim()
    if (!text || text.length > 300) throw new Error("文章は1〜300文字で入力してください。")
    return text
}
function cutInLabel(text) {return `カットイン：${text.replace(/\s+/g, " ").slice(0, 70)}`}
function commitReplayCutInFrames(frames, index) {
    const next = {...replay.data, frames}
    validateReplayRecord(next)
    validateRecord(buildRecord(next))
    pauseReplay()
    replay.data = next
    replay.epoch += 1
    ui.touched = true
    showReplayFrame(index)
    scheduleSave(true)
    return true
}
function addReplayTextCutIn(value) {
    if (!replay.recording && replay.mode !== "replay") return false
    try {
        const text = cutInText(value)
        if (replay.data.frames.length >= REPLAY_MAX_OPERATIONS + 1) throw new Error("1つのリプレイは200操作までです。")
        if (replay.recording) return captureReplayFrame(cutInLabel(text), {text})
        const frame = copy(replay.data.frames[replay.index])
        frame.label = cutInLabel(text)
        frame.delay = 1500
        frame.view.cutIn = {text}
        const frames = [...replay.data.frames]
        frames.splice(replay.index + 1, 0, frame)
        return commitReplayCutInFrames(frames, replay.index + 1)
    } catch (error) {notify(`追加できませんでした。${error.message}`, true); return false}
}
function editReplayTextCutIn(value) {
    if (replay.mode !== "replay" || !replay.data.frames[replay.index]?.view.cutIn) return false
    try {
        const text = cutInText(value)
        const frame = copy(replay.data.frames[replay.index])
        frame.label = cutInLabel(text)
        frame.view.cutIn = {text}
        const frames = [...replay.data.frames]
        frames[replay.index] = frame
        return commitReplayCutInFrames(frames, replay.index)
    } catch (error) {notify(`変更できませんでした。${error.message}`, true); return false}
}
function deleteReplayTextCutIn() {
    if (replay.mode !== "replay" || replay.index === 0 || !replay.data.frames[replay.index]?.view.cutIn) return false
    try {
        const frames = replay.data.frames.filter((_, index) => index !== replay.index)
        return commitReplayCutInFrames(frames, replay.index - 1)
    } catch (error) {notify(`削除できませんでした。${error.message}`, true); return false}
}
function openReplayCutInEditor(edit = false) {
    if (!replay.recording && replay.mode !== "replay") return
    const current = replay.data.frames[replay.index]
    if (edit && (replay.mode !== "replay" || !current?.view.cutIn)) return
    if (replay.mode === "replay") pauseReplay()
    replay.cutInEditor = {edit, data: replay.data, index: replay.index, mode: replay.mode}
    const text = edit ? current.view.cutIn.text : ""
    const position = replay.recording ? "いまの場面の後" : replay.index === 0 ? "最初の操作の前" : `ステップ${replay.index}の後`
    openModal(edit ? "テキストカットインを編集" : "テキストカットインを追加", `<div class="cut-in-editor"><p>${edit ? "このカットインの文章を変更します。" : `${position}に、文章だけの1ステップを挿入します。`}</p><label for="replayCutInText">表示する文章</label><textarea id="replayCutInText" rows="4" maxlength="300" placeholder="例：ここで手札を補充します">${escapeHTML(text)}</textarea><p class="help-note">300文字まで。改行できます。表示時間は「1ステップ」の設定に従います。</p><div class="modal-footer"><button data-action="close-modal">キャンセル</button><button data-action="replay-save-cutin" class="primary">${edit ? "変更を保存" : "追加する"}</button></div></div>`, "replay-cutin")
    byId("replayCutInText").focus?.()
}
function saveReplayCutInFromEditor() {
    const editor = replay.cutInEditor
    if (modalKind !== "replay-cutin" || !editor) return false
    if (editor.data !== replay.data || editor.mode !== replay.mode || (replay.mode === "replay" && editor.index !== replay.index)) {
        notify("再生位置が変わりました。編集画面を開き直してください。", true)
        return false
    }
    const result = editor.edit ? editReplayTextCutIn(byId("replayCutInText").value) : addReplayTextCutIn(byId("replayCutInText").value)
    if (result) {
        closeModal(false)
        replay.cutInEditor = null
        notify(editor.edit ? "カットインを変更しました。" : "カットインを追加しました。")
    }
    return result
}

const replayInspectionPageSize = 24
const replayInspectionUi = {key: "", page: 0}

function syncReplayInspectionPage(inspection) {
    const key = JSON.stringify([replay.index, inspection])
    if (replayInspectionUi.key !== key) {
        replayInspectionUi.key = key
        replayInspectionUi.page = 0
    }
    const pages = Math.max(1, Math.ceil(inspection.cards.length / replayInspectionPageSize))
    replayInspectionUi.page = integer(replayInspectionUi.page, 0, pages - 1, 0)
    return pages
}
function replayInspectionHTML(inspection) {
    if (!inspection || !Array.isArray(inspection.cards)) return ""
    const pages = syncReplayInspectionPage(inspection)
    const start = replayInspectionUi.page * replayInspectionPageSize
    const playerName = state.players[inspection.player]?.name || "プレイヤー"
    const topCard = state.locations[inspection.player]?.deck[0]
    const kindLabel = {top: "上から確認", all: "山札の一覧", card: "カードの確認"}[inspection.kind] || "カードの確認"
    const cards = inspection.cards.slice(start, start + replayInspectionPageSize).map((id, index) => {
        const card = state.cards[id]
        const item = card ? displayCardDefinition(id) : undefined
        const name = item?.name || "カード"
        const image = item?.asset ? cachedCardImageURL(item.asset) : null
        const face = image
            ? liverImageFaceHTML(item, `<img src="${escapeHTML(image)}" alt="${escapeHTML(name)}" draggable="false" loading="lazy" decoding="async">`, card.powerOverride)
            : textCardFaceHTML(item || {name}, "replay-inspection-text")
        return `<figure class="replay-inspection-card"><span class="replay-inspection-order">${start + index + 1}${id === topCard ? " / 先頭" : ""}</span><div class="replay-inspection-face">${face}</div><figcaption>${escapeHTML(name)}</figcaption></figure>`
    }).join("")
    const pagination = pages > 1 ? `<div class="replay-inspection-pagination"><span class="replay-inspection-range">${start + 1}〜${Math.min(start + replayInspectionPageSize, inspection.cards.length)}枚目 / 全${inspection.cards.length}枚</span><button class="small" data-action="replay-inspection-prev" ${replayInspectionUi.page === 0 ? "disabled" : ""}>前へ</button><span>${replayInspectionUi.page + 1} / ${pages}</span><button class="small" data-action="replay-inspection-next" ${replayInspectionUi.page + 1 >= pages ? "disabled" : ""}>次へ</button></div>` : ""
    return `<section class="replay-inspection" aria-label="山札で確認したカード"><div class="replay-inspection-heading"><strong>山札で確認したカード</strong><span>${escapeHTML(playerName)} · ${kindLabel} · ${inspection.cards.length}枚</span>${inspection.cards.length ? `<span class="replay-inspection-direction">${inspection.cards[start] === topCard ? "左が先頭" : "左から順に表示"}</span>` : ""}</div>${cards ? `<div class="replay-inspection-cards">${cards}</div>` : `<p class="replay-inspection-empty">確認したカードはすべて移動しました。</p>`}${pagination}</section>`
}

function positionReplayCutIn() {
    const overlay = byId("replayCutIn")
    const bar = byId("replayBar")
    if (!overlay || overlay.hidden || typeof overlay.style?.setProperty !== "function" || typeof bar?.getBoundingClientRect !== "function") return
    const bottom = bar.getBoundingClientRect().bottom
    if (Number.isFinite(bottom)) overlay.style.setProperty("--replay-cut-in-top", `${Math.max(10, Math.ceil(bottom) + 14)}px`)
}

function renderReplayCutIn() {
    const overlay = byId("replayCutIn")
    if (!overlay) return
    const frame = replay.data?.frames[replay.index]
    const previousFrame = replay.data?.frames[replay.index - 1]
    const phaseText = frame && previousFrame && (frame.state.phase ?? 0) !== (previousFrame.state.phase ?? 0)
        ? `${phases[frame.state.phase ?? 0]}フェイズ` : ""
    const cutIn = frame?.view?.cutIn
    const inspection = frame?.view?.inspection
    const hasText = typeof cutIn?.text === "string" && cutIn.text.trim().length > 0
    const hasCards = Array.isArray(inspection?.cards) && inspection.cards.length > 0
    if (replay.mode !== "replay" || replay.dismissedCutInIndex === replay.index || (!hasText && !hasCards && !phaseText)) {
        overlay.hidden = true
        overlay.innerHTML = ""
        replay.cutInKey = ""
        replayInspectionUi.key = ""
        replayInspectionUi.page = 0
        return
    }
    if (hasText || !hasCards) {
        replayInspectionUi.key = ""
        replayInspectionUi.page = 0
    } else syncReplayInspectionPage(inspection)
    const key = JSON.stringify([replay.index, hasText ? cutIn : [phaseText, inspection], hasText ? null : replayInspectionUi.page])
    if (replay.cutInKey !== key) {
        const textOnly = hasText || !hasCards
        const content = hasText
            ? `<p class="replay-cut-in-message">${escapeHTML(cutIn.text)}</p>`
            : `${phaseText ? `<p class="replay-cut-in-message${hasCards ? " replay-phase-message" : ""}">${escapeHTML(phaseText)}</p>` : ""}${hasCards ? replayInspectionHTML(inspection) : ""}`
        overlay.innerHTML = `<section class="replay-cut-in-panel ${textOnly ? "replay-cut-in-text" : "replay-cut-in-cards"}" aria-label="${hasText ? "テキストカットイン" : phaseText ? "フェイズのカットイン" : "カードのカットイン"}"><button class="replay-cut-in-close small" data-action="replay-dismiss-cutin" aria-label="カットインを閉じる">閉じる</button>${content}</section>`
        replay.cutInKey = key
    }
    overlay.hidden = false
    positionReplayCutIn()
}

function handleReplayInspectionAction(action, button) {
    if (!["replay-inspection-prev", "replay-inspection-next"].includes(action)) return false
    const frame = replay.data?.frames[replay.index]
    const inspection = frame?.view?.inspection
    if (replay.mode !== "replay" || replay.dismissedCutInIndex === replay.index || frame?.view?.cutIn || !Array.isArray(inspection?.cards) || button?.disabled) return true
    const pages = syncReplayInspectionPage(inspection)
    const delta = action === "replay-inspection-next" ? 1 : -1
    replayInspectionUi.page = integer(replayInspectionUi.page + delta, 0, pages - 1, 0)
    renderReplayCutIn()
    return true
}

function renderReplayControls() {
    const bar = byId("replayBar")
    if (!bar) {renderReplayCutIn(); return}
    const frames = replay.data?.frames || []
    const operationCount = Math.max(0, frames.length - 1)
    const hasReplay = frames.length > 0
    const replayMode = replay.mode === "replay"
    const index = Math.max(0, Math.min(operationCount, Number(replay.index) || 0))
    bar.className = `replay-bar${replay.recording ? " is-recording" : ""}${replayMode ? " is-playing-replay" : ""}`

    if (replayMode) {
        const frameLabel = index === 0 ? "開始時の盤面" : (frames[index]?.label || `操作 ${index}`)
        const hasCutIn = !!frames[index]?.view?.cutIn
        bar.innerHTML = `<div class="replay-topline"><div class="replay-heading"><span class="replay-mark" aria-hidden="true">▶</span><strong>リプレイ</strong><span class="replay-state">${replay.playing ? "再生中" : "停止中"}</span></div><div class="replay-actions"><button data-action="replay-add-cutin" class="small">テキストを挿入</button>${hasCutIn ? `<button data-action="replay-edit-cutin" class="small">テキストを編集</button><button data-action="replay-delete-cutin" class="small danger">テキストを削除</button>` : ""}<button data-action="replay-save" class="small">リプレイ保存</button><button data-action="replay-exit" class="primary small">操作に戻る</button></div></div>
            <div class="replay-playback"><div class="replay-transport"><button data-action="replay-prev" class="small" ${index === 0 ? "disabled" : ""}>前へ</button><button data-action="replay-play-pause" class="primary small" ${operationCount === 0 ? "disabled" : ""}>${replay.playing ? "一時停止" : "再生"}</button><button data-action="replay-next" class="small" ${index >= operationCount ? "disabled" : ""}>次へ</button></div><div class="replay-timeline"><input id="replaySeek" type="range" min="0" max="${operationCount}" step="1" value="${index}" aria-label="再生位置" aria-valuetext="${index} / ${operationCount}：${escapeHTML(frameLabel)}" ${operationCount === 0 ? "disabled" : ""}><span class="replay-position">${index}<span> / ${operationCount}</span></span></div><label class="replay-step-time" for="replayStepSeconds">1ステップ<select id="replayStepSeconds">${[0.5, 1, 1.5, 2, 3, 5].map(seconds => `<option value="${seconds}" ${Number(replay.stepSeconds ?? 1) === seconds ? "selected" : ""}>${seconds.toFixed(1)}秒</option>`).join("")}</select></label></div>
            <div class="replay-caption"><span class="replay-operation" title="${escapeHTML(frameLabel)}">${escapeHTML(frameLabel)}</span><span class="replay-return-note">「操作に戻る」で元の盤面へ</span></div>`
        renderReplayCutIn()
        return
    }

    if (replay.recording) {
        bar.innerHTML = `<div class="replay-topline"><div class="replay-heading"><span class="replay-record-dot" aria-hidden="true"></span><strong>記録中</strong><span class="replay-state" role="status">${operationCount}操作</span></div><div class="replay-actions"><button data-action="replay-add-cutin" class="small">テキストカットイン</button><button data-action="replay-stop-recording" class="small replay-stop">記録終了</button></div></div>`
        renderReplayCutIn()
        return
    }

    bar.innerHTML = `<div class="replay-topline"><div class="replay-heading"><span class="replay-mark" aria-hidden="true">▶</span><strong>リプレイ</strong>${hasReplay ? `<span class="replay-state">${operationCount}操作</span>` : `<span class="replay-state">盤面の操作を記録</span>`}</div><div class="replay-actions"><button data-action="replay-record" class="small"><span class="replay-record-dot" aria-hidden="true"></span>記録開始</button><button data-action="replay-enter" class="small" ${hasReplay ? "" : "disabled"}>再生</button><button data-action="replay-save" class="small" ${hasReplay ? "" : "disabled"}>リプレイ保存</button><button data-action="replay-load" class="small">リプレイ読込</button></div></div>`
    renderReplayCutIn()
}

// REPLAY_END
function visiblePlayerIds() {return state.showOpponent ? playerIds : ["p1"]}
function isZoneVisible(player, zone) {return player === "p1" || state.showOpponent || zone === "colorline" || zone.startsWith("niji_")}
function normalizePlayerView() {
    state.showOpponent = state.showOpponent === true
    if (!state.showOpponent) {
        state.active = "p1"
        editorPlayer = "p1"
        ui.selection = new Set(selectedIds().filter(id => {
            const place = where(id)
            return place && isZoneVisible(place.player, place.zone)
        }))
    }
    if (!playerIds.includes(ui.movePlayer)) ui.movePlayer = "p1"
    if (!isZoneVisible(ui.movePlayer, ui.moveZone)) ui.moveZone = "colorline"
}
function notify(message, error = false) {
    const element = byId("toast")
    element.textContent = message
    element.classList.toggle("error", error)
    element.style.display = "block"
    clearTimeout(toastTimer)
    toastTimer = setTimeout(() => {element.style.display = "none"}, error ? 6500 : 3300)
}
function zoneById(id) {return state.zones.find(zone => zone.id === id)}
function where(id) {
    for (const player of playerIds) {
        for (const zone of state.zones) {
            const index = state.locations[player][zone.id].indexOf(id)
            if (index >= 0) return {player, zone: zone.id, index}
        }
    }
    return null
}
function definition(id) {return state.catalog[state.cards[id].definition]}
function displayCardDefinition(id) {
    const item = definition(id)
    const power = state.cards[id].powerOverride
    return textCardType(item) === "liver" && power !== undefined && power !== null ? {...item, power} : item
}
function setSelectedPower(value) {
    if (replay.mode === "replay") return false
    const ids = selectedIds()
    if (!ids.length || !ids.every(id => !colorForCard(id) && textCardType(definition(id)) === "liver")) return false
    if (value !== null && !((typeof value === "number" && Number.isSafeInteger(value)) || (typeof value === "string" && /^\d+$/.test(value.trim())))) {
        notify("パワーは0〜999999の整数で入力してください。", true)
        return false
    }
    const power = value === null ? null : Number(value)
    if (power !== null && (!Number.isSafeInteger(power) || power < 0 || power > 999999)) {
        notify("パワーは0〜999999の整数で入力してください。", true)
        return false
    }
    if (ids.every(id => power === null ? state.cards[id].powerOverride == null : state.cards[id].powerOverride === power)) return true
    return transact(power === null ? "選択したライバーのパワーを元に戻す" : `選択したライバーのパワーを${power}に変更`, () => {
        for (const id of ids) {
            if (power === null) delete state.cards[id].powerOverride
            else state.cards[id].powerOverride = power
        }
    })
}
function visible(id, place = where(id)) {
    const card = state.cards[id]
    return !!card && !!place && !card.faceDown && !(place.zone === "hand" && place.player !== state.active && !ui.bothHands)
}
function selectedIds() {return [...ui.selection].filter(id => Object.hasOwn(state.cards, id))}
function integer(value, minimum, maximum, fallback = minimum) {
    const number = Number(value)
    return Number.isFinite(number) ? Math.min(maximum, Math.max(minimum, Math.trunc(number))) : fallback
}
function logAction(label) {
    state.log.push({time: new Date().toISOString(), label})
    state.log = state.log.slice(-200)
}
function transact(label, operation, {recordLog = true} = {}) {
    if (replay.mode === "replay") return false
    const previous = copy(state)
    try {
        operation()
        checkState(state, assets)
        if (appPage === "deck") {ui.touched = true; render(); scheduleSave(); return true}
        undoStack.push(previous)
        if (undoStack.length > 80) undoStack.shift()
        redoStack = []
        if (recordLog) logAction(label)
        captureReplayFrame(label)
        ui.touched = true
        ui.selection = new Set(selectedIds())
        render()
        scheduleSave()
        return true
    } catch (error) {
        state = previous
        console.error("Operation failed", error)
        render()
        notify(`操作を取り消しました。${error.message}`, true)
        return false
    }
}
function clearHistory() {
    if (!state.log.length) return false
    return transact("履歴を初期化", () => {state.log = []}, {recordLog: false})
}
function undo() {
    if (replay.mode === "replay") return
    if (!undoStack.length) return
    redoStack.push(copy(state))
    state = undoStack.pop()
    captureReplayFrame("操作を元に戻す")
    ui.selection.clear()
    ui.touched = true
    render()
    scheduleSave()
}
function redo() {
    if (replay.mode === "replay") return
    if (!redoStack.length) return
    undoStack.push(copy(state))
    state = redoStack.pop()
    captureReplayFrame("操作をやり直す")
    ui.selection.clear()
    ui.touched = true
    render()
    scheduleSave()
}
function randomBelow(maximum) {
    const buffer = new Uint32Array(1)
    const limit = Math.floor(4294967296 / maximum) * maximum
    do {crypto.getRandomValues(buffer)} while (buffer[0] >= limit)
    return buffer[0] % maximum
}
function shuffleArray(array) {
    for (let index = array.length - 1; index > 0; index -= 1) {
        const target = randomBelow(index + 1)
        const temporary = array[index]
        array[index] = array[target]
        array[target] = temporary
    }
}
function moveRaw(ids, player, destination, position = "last") {
    const valid = [...new Set(ids)].filter(id => Object.hasOwn(state.cards, id))
    if (destination === "channel" || valid.some(id => where(id)?.zone === "channel")) throw new Error("チャンネルはデッキ編集で設定してください。")
    if (!playerIds.includes(player) || !zoneById(destination)) throw new Error("移動先のゾーンがありません。")
    const target = state.locations[player][destination]
    if (destination === "set" || destination === "stage") {
        for (const id of valid) {
            const place = where(id)
            state.locations[place.player][place.zone].splice(place.index, 1)
            for (const previousId of target.splice(0)) {
                state.locations[player].discard.push(previousId)
                state.cards[previousId].faceDown = false
                state.cards[previousId].rotated = false
            }
            target.push(id)
            state.cards[id].faceDown = zoneById(destination).hidden
            if (zoneById(destination).mode === "pile") state.cards[id].rotated = false
        }
        return
    }
    for (const id of valid) {
        const place = where(id)
        state.locations[place.player][place.zone].splice(place.index, 1)
    }
    if (position === "first") target.unshift(...valid)
    else target.push(...valid)
    for (const id of valid) {
        state.cards[id].faceDown = zoneById(destination).hidden
        if (destination === "hand" || zoneById(destination).mode === "pile") state.cards[id].rotated = false
    }
}
function reorderZone(ids, player, zone, index, finalPosition = false) {
    if (!playerIds.includes(player) || !zoneById(zone)) return false
    const current = state.locations[player][zone]
    const chosen = new Set(ids)
    const moving = current.filter(id => chosen.has(id))
    if (!moving.length || moving.length !== chosen.size) return false
    const boundary = integer(index, 0, current.length, current.length)
    const remaining = current.filter(id => !chosen.has(id))
    const insertion = finalPosition ? Math.min(boundary, remaining.length)
        : boundary - current.slice(0, boundary).filter(id => chosen.has(id)).length
    const next = [...remaining.slice(0, insertion), ...moving, ...remaining.slice(insertion)]
    if (next.every((id, position) => id === current[position])) return false
    return transact(`${state.players[player].name}の${zoneById(zone).label}を並べ替え`, () => {
        state.locations[player][zone] = next
    })
}
function moveCards(ids, player, destination, position = "last") {
    if (!ids.length || !playerIds.includes(player) || !zoneById(destination)) return
    if (destination === "channel" || ids.some(id => where(id)?.zone === "channel")) return notify("チャンネルはデッキ編集で設定してください。")
    if (ids.some(id => colorForCard(id)) && destination !== "colorline" && !destination.startsWith("niji_")) return notify("カラーカードはカラーラインかにじエリアへ移動してください。")
    if (ids.every(id => where(id)?.player === player && where(id)?.zone === destination)) {
        return reorderZone(ids, player, destination, position === "first" ? 0 : state.locations[player][destination].length)
    }
    transact(`${ids.length}枚を${state.players[player].name}の${zoneById(destination).label}へ移動`, () => moveRaw(ids, player, destination, position))
}
function dropCardsAt(ids, player, destination, boundary, allowPileReorder = false) {
    if (!ids.length || !playerIds.includes(player) || !zoneById(destination)) return false
    if (destination === "channel" || ids.some(id => where(id)?.zone === "channel")) {
        notify("チャンネルはデッキ編集で設定してください。")
        return false
    }
    if (ids.some(id => colorForCard(id)) && destination !== "colorline" && !destination.startsWith("niji_")) {
        notify("カラーカードはカラーラインかにじエリアへ移動してください。")
        return false
    }
    const zone = zoneById(destination)
    const current = state.locations[player][destination]
    if (ids.every(id => where(id)?.player === player && where(id)?.zone === destination)) {
        if ((zone.mode === "pile" && !allowPileReorder) || destination === "set" || destination === "stage") return false
        return reorderZone(ids, player, destination, boundary)
    }
    if (destination === "set" || destination === "stage") return transact(`${ids.length}枚を${state.players[player].name}の${zone.label}へ移動`, () => moveRaw(ids, player, destination))
    const bounded = integer(boundary, 0, current.length, current.length)
    const insertion = bounded - current.slice(0, bounded).filter(id => ids.includes(id)).length
    const retained = ids.filter(id => where(id)?.player === player && where(id)?.zone === destination)
        .map(id => [id, {faceDown: state.cards[id].faceDown, rotated: state.cards[id].rotated}])
    return transact(`${ids.length}枚を${state.players[player].name}の${zone.label}へ移動`, () => {
        moveRaw(ids, player, destination)
        const moved = current.splice(current.length - ids.length, ids.length)
        current.splice(insertion, 0, ...moved)
        for (const [id, appearance] of retained) Object.assign(state.cards[id], appearance)
    })
}
function moveDeckSelection(destination, rotated = false, sourceCard = null, position = "last") {
    if (!zoneById(destination) || destination === "channel" || destination === "colorline" || destination.startsWith("niji_")) return
    const inspectingDeck = modalKind === "zone" && inspectContext?.zone === "deck"
    const displayed = inspectingDeck ? new Set(inspectContext.limit === null ? state.locations[inspectContext.player].deck : inspectContext.snapshot) : null
    const moving = (sourceCard ? [sourceCard] : selectedIds()).filter(id => where(id)?.zone === "deck" && !colorForCard(id) && (!displayed || (displayed.has(id) && where(id).player === inspectContext.player)))
    if (!moving.length) return notify("山札のカードを選択してください。")
    const grouped = playerIds.map(player => [player, moving.filter(id => where(id).player === player)])
    const inspection = inspectingDeck ? inspectContext : null
    const previousSnapshot = inspection?.snapshot
    const removeFromInspection = previousSnapshot && (destination !== "deck" || position === "last")
    if (!transact(`山札の${moving.length}枚を${rotated && destination === "field" ? "横向きで" : ""}${zoneById(destination).label}へ移動`, () => {
        for (const [player, ids] of grouped) if (ids.length) {
            moveRaw(ids, player, destination, position)
            if (destination === "field") for (const id of ids) state.cards[id].rotated = rotated === true
        }
        if (removeFromInspection) inspection.snapshot = previousSnapshot.filter(id => !moving.includes(id))
    })) {
        if (removeFromInspection) {
            inspection.snapshot = previousSnapshot
            if (inspectContext === inspection && modalKind === "zone") renderZoneInspector()
        }
        return
    }
    ui.selection.clear()
    if (sourceCard && modalKind === "preview") closeModal()
    render()
}
function draw(player = state.active, requested = ui.drawCount) {
    const count = Math.min(integer(requested, 1, 500), state.locations[player].deck.length)
    if (!count) return notify("山札が空です。")
    transact(`${state.players[player].name}が${count}枚ドロー`, () => moveRaw(state.locations[player].deck.slice(0, count), player, "hand"))
    if (count < requested) notify(`山札の残り${count}枚を引きました。勝敗は判定しません。`)
}
function shuffleDeck(player = state.active) {
    if (!state.locations[player].deck.length) return notify("山札が空です。")
    transact(`${state.players[player].name}の山札をシャッフル`, () => {
        shuffleArray(state.locations[player].deck)
        for (const id of state.locations[player].deck) state.cards[id].faceDown = true
    })
}
function operateSelectedDeck(operation, player, requested) {
    if (!playerIds.includes(player) || !selectedIds().some(id => where(id)?.player === player && where(id)?.zone === "deck")) return
    const count = integer(requested, 1, 500)
    const deck = state.locations[player].deck
    if (!deck.length) return notify("山札が空です。")
    if (operation === "peek") return inspectZone(player, "deck", count)
    if (operation === "draw") draw(player, count)
    else if (operation === "discard") {
        const moving = deck.slice(0, count)
        if (!transact(`${state.players[player].name}の山札の上から${moving.length}枚をログへ移動`, () => moveRaw(moving, player, "discard"))) return
    } else return
    ui.selection.clear()
    const top = state.locations[player].deck[0]
    if (top) ui.selection.add(top)
    ui.movePlayer = player
    refreshSelectionUI()
}
function shuffleColorCards() {
    const lines = playerIds.map(player => {
        const line = state.locations[player].colorline
        return {player, line, colors: line.filter(id => colorForCard(id))}
    }).filter(({colors}) => colors.length >= 2)
    if (!lines.length) {notify("カラーラインにカラーカードが2枚以上必要です。"); return false}
    return transact("両側のカラーカードをシャッフル", () => {
        for (const {player, line, colors} of lines) {
            shuffleArray(colors)
            let index = 0
            state.locations[player].colorline = line.map(id => colorForCard(id) ? colors[index++] : id)
        }
    })
}
function readyField() {
    transact(`${state.players[state.active].name}のフィールドをすべて縦向きに変更`, () => {
        for (const zone of state.zones.filter(zone => zone.id === "field")) {
            for (const id of state.locations[state.active][zone.id]) state.cards[id].rotated = false
        }
    })
}
function resetBoard() {
    const restored = transact("両側の盤面を初期状態に戻す", () => {
        for (const player of playerIds) createDeckRaw(player, state.decklists[player])
        for (const player of playerIds) resetColorCards(player)
        state.turn = 1
        state.phase = 0
        state.active = "p1"
    })
    if (!restored) return
    ui.selection.clear()
    ui.tab = "operation"
    ui.movePlayer = "p1"
    ui.moveZone = "field"
    ui.position = "last"
    ui.dragging = false
    closeModal()
    render()
    notify("初期盤面に戻しました。「元に戻す」で取り消せます。")
}
function redeal() {
    const player = state.active
    if (!state.zones.some(zone => zone.id !== "channel" && state.locations[player][zone.id].some(id => !colorForCard(id)))) return notify("引き直せるカードがありません。先にデッキを登録してください。")
    transact(`${state.players[player].name}の初期配置を作成（${ui.initialCount}枚指定）`, () => {
        const ids = state.zones.filter(zone => zone.id !== "channel").flatMap(zone => state.locations[player][zone.id].filter(id => !colorForCard(id)))
        moveRaw(ids, player, "deck")
        for (const id of ids) {
            state.cards[id].counter = 0
            state.cards[id].note = ""
            state.cards[id].rotated = false
            delete state.cards[id].powerOverride
        }
        shuffleArray(ids)
        let index = 0
        state.locations[player].deck = state.locations[player].deck.map(id => colorForCard(id) ? id : ids[index++])
        moveRaw(ids.slice(0, ui.initialCount), player, "hand")
        for (const meter of state.players[player].meters) meter.value = 0
    })
}
function createDeckRaw(player, list, channelDefinition = state.channelDefinitions[player] || "") {
    if (list.some(item => isColorDefinition(state.catalog[item.definition]))) throw new Error("Color cards cannot enter a deck list.")
    if (channelDefinition && (!state.catalog[channelDefinition] || isColorDefinition(state.catalog[channelDefinition]))) throw new Error("Invalid channel definition.")
    const mainList = list.filter(item => item.definition !== channelDefinition && textCardType(state.catalog[item.definition]) !== "channel")
    for (const zone of state.zones) {
        const colors = state.locations[player][zone.id].filter(id => colorForCard(id))
        for (const id of state.locations[player][zone.id]) if (!colorForCard(id)) delete state.cards[id]
        state.locations[player][zone.id] = colors
    }
    state.decklists[player] = copy(mainList)
    state.channelDefinitions[player] = channelDefinition
    for (const item of mainList) {
        for (let index = 0; index < item.count; index += 1) {
            const id = uid("card")
            state.cards[id] = {definition: item.definition, faceDown: true, rotated: false, counter: 0, note: ""}
            state.locations[player].deck.push(id)
        }
    }
    if (channelDefinition) {
        const id = uid("card")
        state.cards[id] = {definition: channelDefinition, faceDown: false, rotated: false, counter: 0, note: ""}
        state.locations[player].channel.push(id)
    }
    for (const meter of state.players[player].meters) meter.value = 0
}
function demo() {
    if ((Object.values(state.catalog).some(item => !isColorDefinition(item)) || Object.entries(state.cards).some(([id, card]) => !card.faceDown || card.note || card.counter || card.rotated || where(id)?.zone !== "colorline")) && !confirm("現在のカード登録と盤面を動作確認用デモに置き換えます。必要なら先にJSON保存してください。")) return
    transact("仮カードによる動作確認用デモを作成", () => {
        const showOpponent = state.showOpponent
        state = blankState()
        state.showOpponent = showOpponent
        const list = []
        for (let index = 1; index <= 6; index += 1) {
            const id = uid("definition")
            state.catalog[id] = {name: `仮カード ${String(index).padStart(2, "0")}`, text: "動作確認用の仮カードです。公式カードの名前や効果ではありません。", asset: ""}
            list.push({definition: id, count: 3})
        }
        const channelDefinition = uid("definition")
        state.catalog[channelDefinition] = {name: "仮チャンネル", text: "動作確認用の横長カードです。公式カードではありません。", asset: ""}
        for (const player of visiblePlayerIds()) {
            createDeckRaw(player, list, channelDefinition)
            shuffleArray(state.locations[player].deck)
            moveRaw(state.locations[player].deck.slice(0, 4), player, "hand")
            moveRaw(state.locations[player].deck.slice(0, 1), player, "field")
            moveRaw(state.locations[player].deck.slice(0, 1), player, "set")
            moveRaw(state.locations[player].deck.slice(0, 1), player, "niji_r")
        }
    })
    notify("仮カードを各エリアに置きました。デモの枚数や配置は動作確認用です。")
}
// LIVER_CARD_DETAILS_START
function textCardType(item) {
    return item.cardType ?? (item.isLiver ? "liver" : "other")
}
function liverCardColors(item) {
    if (textCardType(item) !== "liver" || isColorDefinition(item) || !Array.isArray(item.colors)) return []
    return nijiColors.filter(color => item.colors.includes(color))
}
function liverColorBandsHTML(item) {
    const colors = liverCardColors(item)
    if (!colors.length) return ""
    const description = `ライバーの色：${colors.map(color => colorNames[color]).join("・")}`
    return `<span class="liver-color-bands" role="img" aria-label="${escapeHTML(description)}" title="${escapeHTML(description)}">${colors.map(color => `<span class="liver-color-band color-${color.toLowerCase()}" style="grid-row:${nijiColors.indexOf(color) + 1}" aria-hidden="true"></span>`).join("")}</span>`
}
function cardTypeSideLabel(item) {
    const type = textCardType(item)
    return type === "event" ? "EVENT" : type === "stage" ? "STAGE" : ""
}
function cardTypeSideLabelHTML(item) {
    const label = cardTypeSideLabel(item)
    return label ? `<span class="card-type-side-label" aria-label="${label === "EVENT" ? "イベントカード" : "ステージカード"}">${label}</span>` : ""
}
function textCardFaceHTML(item, className = "card-body", powerOverride = undefined) {
    const type = textCardType(item)
    const cost = type === "channel" || item.cost === undefined || item.cost === null ? "" : `<span class="text-card-stat text-card-cost" title="レベル ${escapeHTML(item.cost)}">${escapeHTML(item.cost)}</span>`
    const classNameText = (item.className || "").trim()
    const initial = Array.from(classNameText)[0] || ""
    const cardClass = initial ? `<span class="text-card-stat text-card-class" title="クラス ${escapeHTML(classNameText)}">${escapeHTML(initial)}</span>` : ""
    const powerValue = powerOverride === undefined ? item.power : powerOverride
    const power = type === "liver" && powerValue !== undefined && powerValue !== null ? `<span class="text-card-stat text-card-power" title="パワー ${escapeHTML(powerValue)}">${escapeHTML(powerValue)}</span>` : ""
    return `<span class="${className} text-card-face">${cost}${cardClass}<span class="card-symbol" aria-hidden="true">◇</span><strong class="text-card-name">${escapeHTML(item.name)}</strong>${liverColorBandsHTML(item)}${cardTypeSideLabelHTML(item)}${power}</span>`
}
function liverImageFaceHTML(item, imageHtml, powerOverride = undefined) {
    const type = textCardType(item)
    if (!["liver", "event", "stage"].includes(type) || isColorDefinition(item)) return imageHtml
    // The original image already contains its base power. Show a badge only
    // when the card instance has an explicit value, including zero.
    const power = type === "liver" && typeof powerOverride === "number" && Number.isFinite(powerOverride)
        ? `<span class="liver-image-power" title="パワー ${escapeHTML(powerOverride)}">${escapeHTML(powerOverride)}</span>` : ""
    return `<span class="liver-image-face">${imageHtml}${liverColorBandsHTML(item)}${cardTypeSideLabelHTML(item)}${power}</span>`
}
function cardDetailFields(item = {}, definitionId = "") {
    const type = textCardType(item)
    const inputId = field => definitionId ? `cardDetail_${definitionId}_${field}` : {cost: "newCardCost", className: "newCardClass", cardType: "newCardType", power: "newCardPower", colors: "newCardColors", tags: "newCardTags"}[field]
    const attributes = field => `id="${inputId(field)}"${definitionId ? ` data-catalog-detail="${field}" data-definition="${definitionId}"` : ""}`
    const classValue = item.className || ""
    const oldClass = classValue && !textCardClasses.includes(classValue) ? `<option value="${escapeHTML(classValue)}" selected disabled>旧設定：${escapeHTML(classValue)}</option>` : ""
    const classes = ["", ...textCardClasses].map(value => `<option value="${value}" ${classValue === value ? "selected" : ""}>${value || "未設定"}</option>`).join("")
    const types = Object.entries(textCardTypes).map(([value, label]) => `<option value="${value}" ${type === value ? "selected" : ""}>${label}</option>`).join("")
    const selectedColors = Array.isArray(item.colors) ? item.colors : []
    const colors = nijiColors.map(color => {
        const colorId = definitionId ? `cardDetail_${definitionId}_color_${color}` : `newCardColor_${color}`
        const data = definitionId ? `data-catalog-color="${color}" data-definition="${definitionId}"` : `data-new-card-color="${color}"`
        return `<label class="card-color-choice"><input id="${colorId}" type="checkbox" ${data} value="${color}" aria-label="${colorNames[color]}（${color}）" ${selectedColors.includes(color) ? "checked" : ""} ${type === "liver" ? "" : "disabled"}><span class="card-color-swatch color-${color.toLowerCase()}" aria-hidden="true"></span><span>${colorNames[color]}</span></label>`
    }).join("")
    return `<div class="card-detail-fields ${definitionId ? "catalog-card-details" : "new-text-card-details"}"><label>カード種別<select ${attributes("cardType")}>${types}</select></label><label>レベル<input ${attributes("cost")} type="number" min="0" max="999" step="1" value="${type === "channel" ? "" : (item.cost ?? "")}" placeholder="${type === "channel" ? "レベルなし" : "未設定"}" ${type === "channel" ? "disabled" : ""}></label><label>クラス<select ${attributes("className")}>${oldClass}${classes}</select></label><label>パワー<input ${attributes("power")} type="number" min="0" max="999999" step="1" value="${item.power ?? ""}" placeholder="未設定" ${type === "liver" ? "" : "disabled"}></label></div><label class="card-tag-field">タグ（カンマ区切り・#なし）<input ${attributes("tags")} type="text" maxlength="2000" value="${escapeHTML((item.tags || []).join(", "))}" placeholder="囚人, JK" aria-label="ライバーのタグ（カンマ区切り・#なし）" ${type === "liver" ? "" : "disabled"}></label><fieldset id="${inputId("colors")}" class="card-color-fields" ${type === "liver" ? "" : "disabled"}><legend>ライバーの色 <span>複数選択可</span></legend><div class="card-color-choices">${colors}</div></fieldset>`
}
function readCardDetailColors(definitionId = "") {
    const selector = definitionId ? `[data-catalog-color][data-definition="${definitionId}"]:checked` : "[data-new-card-color]:checked"
    return parseCardDetail("colors", [...document.querySelectorAll(selector)].map(element => element.dataset.catalogColor || element.dataset.newCardColor))
}
function syncCardDetailFields(item, definitionId = "") {
    const type = textCardType(item)
    const inputId = field => definitionId ? `cardDetail_${definitionId}_${field}` : {cost: "newCardCost", power: "newCardPower", colors: "newCardColors", tags: "newCardTags"}[field]
    const cost = byId(inputId("cost")), power = byId(inputId("power")), colors = byId(inputId("colors")), tags = byId(inputId("tags"))
    if (cost) {
        cost.disabled = type === "channel"
        cost.placeholder = type === "channel" ? "レベルなし" : "未設定"
        if (type === "channel") cost.value = ""
        else if (definitionId) cost.value = item.cost ?? ""
    }
    if (power) power.disabled = type !== "liver"
    if (colors) colors.disabled = type !== "liver"
    if (tags) tags.disabled = type !== "liver"
    for (const color of nijiColors) {
        const input = byId(definitionId ? `cardDetail_${definitionId}_color_${color}` : `newCardColor_${color}`)
        if (input) input.disabled = type !== "liver"
    }
}
function parseCardDetail(field, raw) {
    if (field === "tags") return normalizeCardTags(raw)
    if (field === "colors") {
        if (!Array.isArray(raw) || raw.length > nijiColors.length || new Set(raw).size !== raw.length || raw.some(color => !nijiColors.includes(color))) throw new Error("ライバーの色は7色から選んでください。")
        return nijiColors.filter(color => raw.includes(color))
    }
    if (typeof raw !== "string") throw new Error("入力値が正しくありません。")
    const value = raw.trim()
    if (field === "className") {
        if (value !== "" && !textCardClasses.includes(value)) throw new Error("クラスは太陽・月・彗星・星から選んでください。")
        return value
    }
    if (field === "cardType") {
        if (!Object.hasOwn(textCardTypes, value)) throw new Error("カード種別を選んでください。")
        return value
    }
    if (!["cost", "power"].includes(field)) throw new Error("編集できない項目です。")
    if (value === "") return null
    const maximum = field === "cost" ? 999 : 999999
    if (!/^\d+$/.test(value) || !Number.isSafeInteger(Number(value)) || Number(value) > maximum) throw new Error(`${field === "cost" ? "レベル" : "パワー"}は0〜${maximum}の整数で入力してください。`)
    return Number(value)
}
function updateCatalogCardDetail(id, field, raw) {
    const item = state.catalog[id]
    if (replay.mode === "replay" || !Object.hasOwn(state.catalog, id) || isColorDefinition(item) || !["cost", "className", "cardType", "power", "colors", "tags"].includes(field) || (field === "cost" && textCardType(item) === "channel") || (["power", "colors", "tags"].includes(field) && textCardType(item) !== "liver")) return false
    try {
        const value = parseCardDetail(field, raw)
        if (field === "cardType" && (value === "channel") !== (textCardType(item) === "channel") &&
            (savedDeckUsesDefinition(id) || deckManagerUsesDefinition(id) || playerIds.some(player => state.channelDefinitions[player] === id || state.decklists[player].some(entry => entry.definition === id)))) {
            throw new Error("デッキからこのカードを外してから、チャンネルとの種別変更を行ってください。")
        }
        const unchanged = ["colors", "tags"].includes(field) ? JSON.stringify(item[field] || []) === JSON.stringify(value) : item[field] === value
        if (unchanged && !(field === "cardType" && (item.isLiver !== undefined || (value === "channel" && item.cost !== null)))) return true
        return transact("カードの情報を変更", () => {
            state.catalog[id][field] = value
            if (field === "cardType") {
                delete state.catalog[id].isLiver
                if (value === "channel") state.catalog[id].cost = null
            }
        })
    } catch (error) {notify(error.message, true); return false}
}

// LIVER_CARD_DETAILS_END
// CARD_IMAGE_CACHE_START
// Render short Blob URLs while keeping the original embedded images for saves.
const cardImageURLs = new Map()

function releaseCardImageURL(assetId) {
    const cached = cardImageURLs.get(assetId)
    if (!cached) return
    try { URL.revokeObjectURL(cached.url) } catch {}
    cardImageURLs.delete(assetId)
}

function cachedCardImageURL(assetId) {
    const data = typeof assetId === "string" && Object.hasOwn(assets, assetId) ? assets[assetId] : ""
    const cached = cardImageURLs.get(assetId)
    if (cached && cached.data === data) return cached.url
    if (cached) releaseCardImageURL(assetId)
    if (typeof data !== "string" || !data) return ""
    // Some non-browser runtimes lack these APIs. The embedded image still works.
    if (typeof atob !== "function" || typeof Blob !== "function" || typeof URL === "undefined" || typeof URL.createObjectURL !== "function" || typeof URL.revokeObjectURL !== "function") return data
    const match = data.length <= 4000000 && /^data:(image\/(?:png|jpeg|webp));base64,([A-Za-z0-9+/]+={0,2})$/.exec(data)
    if (!match) return data
    try {
        const decoded = atob(match[2])
        const bytes = new Uint8Array(decoded.length)
        for (let index = 0; index < decoded.length; index++) bytes[index] = decoded.charCodeAt(index)
        const url = URL.createObjectURL(new Blob([bytes], {type: match[1]}))
        cardImageURLs.set(assetId, {data, url})
        return url
    } catch {
        return data
    }
}

function pruneCardImageURLs() {
    for (const [assetId, cached] of cardImageURLs) {
        if (!Object.hasOwn(assets, assetId) || assets[assetId] !== cached.data) releaseCardImageURL(assetId)
    }
}

// CARD_IMAGE_CACHE_END
function cardHTML(id, options = {}) {
    const card = state.cards[id]
    const item = displayCardDefinition(id)
    const place = options.location || where(id)
    const hidden = options.forceHidden || (!options.reveal && !visible(id, place))
    const color = colorForCard(id)
    const rotated = card.rotated && !options.preview && place?.zone !== "hand"
    const classes = ["card", color ? "color-card" : "", !color && !hidden && !item.asset ? "text-card" : "", color && !hidden ? `color-${color.toLowerCase()}` : "", hidden ? "face-down" : "", rotated ? "rotated" : "", ui.selection.has(id) && !options.preview ? "selected" : ""].filter(Boolean).join(" ")
    const title = hidden ? color ? "裏向きのカラーカード" : "裏向きのカード" : item.name
    let content = hidden ? `<span class="back-mark">7</span>` : item.asset ? `<img src="${cachedCardImageURL(item.asset)}" alt="${escapeHTML(title)}" draggable="false" loading="${options.preview ? "eager" : "lazy"}" decoding="async">` : textCardFaceHTML(item)
    if (!hidden && item.asset) content = liverImageFaceHTML(item, content, card.powerOverride)
    if (color && !hidden) content = `<span class="color-face"><small>COLOR</small><strong>${color}</strong><span>${escapeHTML(colorNames[color])}</span></span>`
    if (!options.preview) {
        if (card.counter) content += `<span class="counter-badge">${card.counter}</span>`
        if (rotated) content += `<span class="rotated-label">横</span>`
        if (card.note) content += `<span class="note-badge"></span>`
    }
    if (options.preview) return `<div class="${classes}">${content}</div>`
    const doubleClickHint = ["deck", "discard"].includes(place?.zone) && !options.reveal ? "ダブルクリックで一覧を表示" : color && place?.zone === "colorline" ? "ダブルクリックで表裏を反転" : "ダブルクリックで内容を確認"
    return `<button class="${classes}" draggable="${place?.zone !== "channel"}" data-card-id="${id}" aria-label="${escapeHTML(title)}" aria-pressed="${ui.selection.has(id)}" title="${escapeHTML(title)}${doubleClickHint ? `。${doubleClickHint}` : ""}">${content}</button>`
}
function zoneHTML(player, zone) {
    const ids = state.locations[player][zone.id]
    const displayed = zone.mode === "pile" ? ids.slice(0, 1) : ids
    const names = {deck: "DECK", field: "FIELD", set: "SET AREA", channel: "CHANNEL", stage: "STAGE AREA", discard: "LOG", colorline: "COLOR LINE", hand: "HAND"}
    const isNiji = zone.id.startsWith("niji_")
    const location = {player, zone: zone.id}
    let content = zone.id === "colorline" ? "" : displayed.map(id => cardHTML(id, {location})).join("")
    if (zone.id === "colorline") content = Array.from({length: Math.max(7, displayed.length)}, (_, index) => `<div class="colorline-slot"><span class="slot-number">${index + 1}</span>${displayed[index] ? cardHTML(displayed[index], {location}) : `<span class="slot-placeholder" aria-hidden="true"></span>`}</div>`).join("")
    if (isNiji && !displayed.length) content = `<span class="zone-empty niji-letter" aria-hidden="true">${escapeHTML(zone.id.slice(-1).toUpperCase())}</span>`
    const dropTarget = zone.id === "channel" ? "" : `data-drop-player="${player}" data-drop-zone="${zone.id}"`
    const zoneAction = ["deck", "discard"].includes(zone.id) ? `<button class="mobile-only" data-action="inspect" data-player="${player}" data-zone="${zone.id}" aria-label="${escapeHTML(zone.label)}の一覧">一覧</button>` : ["set", "stage", "field", "hand", "colorline"].includes(zone.id) ? "" : zone.id === "channel" ? `<button data-action="deck-editor" data-player="${player}" aria-label="チャンネルを設定">設定</button>` : `<button data-action="inspect" data-player="${player}" data-zone="${zone.id}" aria-label="${escapeHTML(zone.label)}の一覧">一覧</button>`
    const listHint = ["deck", "discard"].includes(zone.id) ? ` title="ダブルクリックで一覧を表示"` : ""
    return `<section class="zone zone-${zone.id} ${zone.mode === "hand" ? "hand-zone" : ""}" aria-label="${escapeHTML(zone.label)}"${listHint} ${dropTarget}><div class="zone-head"><span class="zone-title">${names[zone.id] ? `<span class="zone-en">${names[zone.id]}</span>` : ""}${escapeHTML(zone.label)}<span class="zone-count">${ids.length}</span></span>${zoneAction}</div><div class="zone-cards ${zone.mode === "pile" ? "pile" : ""}">${content}</div></section>`
}
function playmatHTML(player) {
    const standard = new Set(defaultZones().map(zone => zone.id))
    const zone = id => zoneById(id) ? zoneHTML(player, zoneById(id)) : ""
    const extras = state.zones.filter(item => !standard.has(item.id))
    return `<div class="playmat-scroll"><div class="playmat">${["colorline", "field", "deck", "set", "channel", "stage", "discard", "hand"].map(zone).join("")}<div class="niji-area"><div class="niji-heading"><span>NIJI AREA</span><span>にじエリア</span></div><div class="niji-grid">${nijiColors.map(color => zone(`niji_${color.toLowerCase()}`)).join("")}</div></div></div></div>${extras.length ? `<details class="extra-zones" open><summary>追加ゾーン <span>${extras.length}</span></summary><div class="zone-grid">${extras.map(item => zoneHTML(player, item)).join("")}</div></details>` : ""}`
}
function opponentColorsHTML() {
    return `<section class="player-panel opponent-color-panel" data-player-panel="p2"><div class="player-heading"><div><span class="name">${escapeHTML(state.players.p2.name)}</span><span class="badge">相手のカラー</span></div></div><div class="opponent-colors-scroll"><div class="opponent-colors">${zoneHTML("p2", zoneById("colorline"))}<div class="niji-area"><div class="niji-heading"><span>NIJI AREA</span><span>相手のにじエリア</span></div><div class="niji-grid">${nijiColors.map(color => zoneHTML("p2", zoneById(`niji_${color.toLowerCase()}`))).join("")}</div></div></div></div></section>`
}
function renderBoard() {
    byId("board").innerHTML = (!state.showOpponent ? opponentColorsHTML() : "") + [...visiblePlayerIds()].reverse().map(player => {
        const data = state.players[player]
        const meters = data.meters.map((meter, index) => `<div class="meter"><span>${escapeHTML(meter.label)}</span><button data-action="meter" data-player="${player}" data-index="${index}" data-delta="-1" aria-label="減らす">−</button><output>${meter.value}</output><button data-action="meter" data-player="${player}" data-index="${index}" data-delta="1" aria-label="増やす">＋</button></div>`).join("")
        return `<section class="player-panel ${player === state.active ? "active-player" : ""}" data-player-panel="${player}"><div class="player-heading"><div><span class="name">${escapeHTML(data.name)}</span><span class="badge">${player === state.active ? "操作側" : "相手側"}</span></div><div class="meters">${meters}</div></div>${playmatHTML(player)}</section>`
    }).join("")
    byId("phaseStrip").innerHTML = phases.map((phase, index) => `<button class="phase-button ${state.phase === index ? "is-current" : ""}" data-action="phase" data-phase="${index}" aria-pressed="${state.phase === index}"><span>0${index + 1}</span>${phase}<small>フェイズ</small></button>`).join("")
}
// MOBILE_CONTROLS_START
const mobileViewport = globalThis.matchMedia?.("(max-width:800px), (max-width:999px) and (hover:none) and (pointer:coarse)")
const mobileUi = {panelOpen: false, multiple: false}
function mobileSelectionTools(allowMultiple = false) {
    const ids = selectedIds()
    const canPreview = ids.length === 1 && (modalKind === "zone" || where(ids[0])?.zone !== "deck")
    return `<div class="mobile-only mobile-selection-tools"><button data-action="preview-selected" data-mobile-preview ${canPreview ? "" : "disabled"}>内容を確認</button><button data-action="clear-selection" data-mobile-clear ${ids.length ? "" : "disabled"}>選択を解除</button>${allowMultiple ? `<button data-action="mobile-multiple" aria-pressed="${mobileUi.multiple}">複数選択：${mobileUi.multiple ? "オン" : "オフ"}</button>` : ""}</div>`
}
function syncMobileControls() {
    if (appPage === "deck") return
    const enabled = !!mobileViewport?.matches
    if (!enabled || replay.mode === "replay") mobileUi.panelOpen = false
    document.body.classList.toggle("mobile-panel-open", enabled && mobileUi.panelOpen)
    byId("sidePanel").dataset.mobileTab = ui.tab
    byId("mobilePanelTitle").textContent = {operation: "基本操作", selection: "選択したカード", history: "履歴"}[ui.tab] || "操作"
    const count = selectedIds().length
    byId("mobileSelectionCount").textContent = `${count}枚${mobileUi.multiple ? " · 複数" : ""}`
    for (const button of byId("mobileNav").querySelectorAll("[data-tab]")) {
        button.disabled = replay.mode === "replay"
        button.setAttribute("aria-expanded", String(enabled && mobileUi.panelOpen && ui.tab === button.dataset.tab))
    }
    const id = selectedIds()[0]
    for (const button of document.querySelectorAll("[data-mobile-preview]")) button.disabled = count !== 1 || (modalKind !== "zone" && where(id)?.zone === "deck")
    for (const button of document.querySelectorAll("[data-mobile-clear]")) button.disabled = !count
}
function setMobilePanel(open, tab = ui.tab, restoreFocus = true) {
    if (appPage === "deck") return
    const hadPanel = mobileUi.panelOpen
    if (open && (!mobileViewport?.matches || replay.mode === "replay")) return
    if (open) {
        ui.tab = ["operation", "selection", "history"].includes(tab) ? tab : "operation"
        mobileUi.panelOpen = true
        renderSideHead()
        renderSideContent()
        byId("sidePanel").scrollTop = 0
        byId("sidePanel").querySelector('[data-action="mobile-close"]').focus({preventScroll: true})
    } else {
        mobileUi.panelOpen = false
        syncMobileControls()
        if (hadPanel && restoreFocus) byId("mobileNav").querySelector(`[data-tab="${ui.tab}"]`)?.focus({preventScroll: true})
    }
}
mobileViewport?.addEventListener("change", () => {
    if (appPage === "deck") return
    if (modalKind === "zone") renderZoneInspector()
    const focusedInside = byId("sidePanel").contains(document.activeElement)
    syncMobileControls()
    if (mobileViewport.matches && focusedInside) byId("mobileNav").querySelector('[data-tab="operation"]').focus({preventScroll: true})
})
// MOBILE_CONTROLS_END
// SIDE_PANEL_UI_START
function renderSideHead() {
    byId("sideHead").innerHTML = `<h3>${state.showOpponent ? "操作するプレイヤー" : "一人回し"}</h3>${state.showOpponent ? `<div class="player-switch">${visiblePlayerIds().map(player => `<button class="${state.active === player ? "current" : ""}" data-action="active" data-player="${player}">${escapeHTML(state.players[player].name)}</button>`).join("")}</div>` : `<div>${escapeHTML(state.players.p1.name)}</div>`}<button class="small full ghost" style="margin-top:9px" data-action="toggle-opponent" aria-pressed="${state.showOpponent}">${state.showOpponent ? "相手はカラーだけ表示" : "相手の全フィールドを表示"}</button><div class="turn-line"><span>ターン表示 <strong>${state.turn}</strong></span><button class="small" data-action="next-turn">次へ</button></div><div class="help-note">「次へ」では操作側を切り替えません。</div><div class="undo-row"><button data-action="undo" ${undoStack.length ? "" : "disabled"}>元に戻す</button><button data-action="redo" ${redoStack.length ? "" : "disabled"}>やり直す</button></div>`
    byId("sideTabs").innerHTML = [["operation", "基本操作"], ["selection", `選択 ${selectedIds().length}`], ["history", "履歴"]].map(([value, label]) => `<button class="${ui.tab === value ? "active" : ""}" data-action="tab" data-tab="${value}">${label}</button>`).join("")
}
function renderOperation() {
    return `<section class="control-group"><h3>山札と手札</h3><div class="control-row"><input id="drawCount" aria-label="引く枚数" type="number" min="1" max="500" value="${ui.drawCount}"><button data-action="draw" class="primary grow">枚引く</button></div><button data-action="shuffle" class="full">シャッフル</button><div class="control-row"><input id="peekCount" aria-label="確認する枚数" type="number" min="1" max="500" value="${ui.peekCount}"><button data-action="peek" class="grow">上から確認</button></div></section>
        <section class="control-group board-controls"><h3>盤面操作</h3><button data-action="ready" class="full">自分のフィールドを全て縦向きにする</button><button data-action="reset-colors" class="full">カラーカードの初期化</button><button data-action="shuffle-colors" class="full">カラーカードのシャッフル</button><div class="control-row"><input id="initialCount" aria-label="ランダムに引き直す枚数" type="number" min="0" max="500" value="${ui.initialCount}"><button data-action="redeal" class="grow">枚をランダムに引き直す</button></div><button data-action="reset-board" class="full">初期盤面に戻す</button></section>
        ${state.showOpponent ? `<div class="control-row"><label><input id="bothHands" type="checkbox" ${ui.bothHands ? "checked" : ""}> 両方の手札を表示する</label></div>` : ""}`
}
function destinationOptions(value, player = ui.movePlayer, colorsOnly = false, rotatedField = false) {
    const zones = state.zones.filter(zone => zone.id !== "channel" && isZoneVisible(player, zone.id)
        && (zone.id === "colorline" || zone.id.startsWith("niji_")) === colorsOnly)
        .flatMap(zone => rotatedField && zone.id === "field" ? [zone, {...zone, id: "field-rotated", label: `${zone.label}（横向き）`}] : [zone])
    const selected = zones.some(zone => zone.id === value) ? value : (zones.find(zone => zone.id === "field") || zones[0])?.id
    return zones.map(zone => `<option value="${zone.id}" ${zone.id === selected ? "selected" : ""}>${escapeHTML(zone.label)}</option>`).join("")
}
function colorSelectionHTML(ids) {
    if (!ids.length || !ids.every(id => colorForCard(id))) return ""
    const owners = [...new Set(ids.map(id => state.cards[id].colorOwner || where(id).player))]
    const allOnLine = ids.every(id => where(id)?.zone === "colorline")
    const flipLabel = ids.every(id => state.cards[id].faceDown) ? "表にする" : ids.every(id => !state.cards[id].faceDown) ? "裏にする" : "表裏を反転"
    return `<section class="control-group color-controls"><h3>${owners.length === 1 && owners[0] === "p2" ? "相手の" : ""}カラーカード</h3><button class="full primary" data-action="flip">${flipLabel}</button><button class="full" data-action="color-to-niji">同じ色のにじエリアへ</button>${allOnLine ? "" : `<button class="full" data-action="color-to-line">裏向きでカラーラインへ</button>`}</section>`
}
function selectedPowerHTML(ids) {
    if (!ids.length || !ids.every(id => textCardType(definition(id)) === "liver" && !colorForCard(id))) return ""
    const powers = ids.map(id => displayCardDefinition(id).power ?? null)
    const same = powers.every(power => power === powers[0])
    const value = same && powers[0] !== null ? String(powers[0]) : ""
    return `<div class="selection-power-control"><label for="selectedPower">パワー</label><div class="control-row"><input id="selectedPower" aria-label="パワー" type="number" min="0" max="999999" step="1" value="${escapeHTML(value)}" placeholder="${same ? "未設定" : "複数の値"}"><button data-action="set-selected-power" class="grow">パワーを変更</button></div><button data-action="reset-selected-power" class="small full">元のパワーに戻す</button></div>`
}
function selectedDeckControls(ids) {
    const players = [...new Set(ids.map(id => where(id).player))]
    const operations = [["draw", "drawCount", "ドローする枚数", "枚ドロー"], ["peek", "peekCount", "上から確認する枚数", "枚確認"], ["discard", "discardCount", "上からログへ送る枚数", "枚ログへ"]]
    return players.map(player => `<section class="control-group"><h3>${players.length > 1 ? `${escapeHTML(state.players[player].name)}の` : ""}山札の操作</h3>${operations.map(([operation, key, label, button]) => `<div class="control-row deck-operation-row">${operation === "draw" ? "" : "<span>上から</span>"}<input id="selectedDeck_${player}_${operation}" data-deck-count="${key}" aria-label="${label}" type="number" min="1" max="500" value="${ui[key]}"><button data-action="selected-deck-${operation}" data-player="${player}" class="${operation === "draw" ? "primary " : ""}grow">${button}</button></div>`).join("")}</section>`).join("")
}
function renderSelection() {
    const ids = selectedIds()
    if (!ids.length) return `<div class="empty-help">カードをクリックすると、移動・表裏・向き・カウンター・メモを操作できます。</div><p class="help-note">ダブルクリックで、カラーラインのカラーカードは表裏を反転し、山札以外のカードは内容を確認できます。</p>`
    const id = ids[0]
    const item = definition(id)
    const card = state.cards[id]
    const place = where(id)
    const revealSelection = place?.player === state.active && place.zone !== "deck" && (place.zone === "set" || !!colorForCard(id))
    const selectionCard = cardHTML(id, {preview: true, reveal: revealSelection, location: place})
    const title = ids.length === 1 ? visible(id, place) || revealSelection ? item.name : "裏向きのカード" : `${ids.length}枚を選択中`
    if (ids.every(id => where(id)?.zone === "deck")) return `<div class="selected-preview">${selectionCard}<div class="selection-title">山札<br><span class="tag">${state.locations[place.player].deck.length}枚</span></div></div>${selectedDeckControls(ids)}`
    if (ids.every(id => where(id)?.zone === "channel")) return `<div class="selected-preview">${selectionCard}<div class="selection-title">${escapeHTML(title)}</div></div><div class="control-group"><button class="full primary" data-action="deck-editor" data-player="${where(id).player}">チャンネルを設定</button></div>`
    if (ids.every(id => colorForCard(id))) return `<div class="selected-preview">${selectionCard}<div class="selection-title">${escapeHTML(title)}<br><span class="tag">${ids.length}枚を選択</span></div></div>${colorSelectionHTML(ids)}<div class="control-group"><h3>そのほかの操作</h3><div class="choice-grid"><button data-action="rotate">向きを変更</button></div><p class="help-note">ドラッグでもカラーラインとにじエリアの間を移動できます。</p></div>`
    if (!visiblePlayerIds().includes(ui.movePlayer)) ui.movePlayer = "p1"
    if (!zoneById(ui.moveZone) || ui.moveZone === "channel" || ui.moveZone === "colorline" || ui.moveZone.startsWith("niji_")) ui.moveZone = "field"
    return `<div class="selected-preview">${selectionCard}<div class="selection-title">${escapeHTML(title)}<br><span class="tag">${ids.length === 1 ? "1枚を選択" : "まとめて操作"}</span></div></div>${colorSelectionHTML(ids)}<div class="control-group"><h3>状態を変更</h3><div class="choice-grid">${ids.some(id => where(id)?.zone !== "hand") ? `<button data-action="rotate">向きを変更</button>` : ""}<button data-action="flip">表・裏を反転</button></div><div class="control-row"><button data-action="card-counter" data-delta="-1">−1</button><span class="grow" style="text-align:center">${ids.length === 1 ? `カウンター ${card.counter}` : "カウンター"}</span><button data-action="card-counter" data-delta="1">＋1</button></div><button data-action="reset-counter" class="small full" style="margin-top:6px">カウンターを0にする</button>${selectedPowerHTML(ids)}</div><div class="control-group"><h3>移動先</h3><select id="movePlayer" class="full">${visiblePlayerIds().map(player => `<option value="${player}" ${player === ui.movePlayer ? "selected" : ""}>${escapeHTML(state.players[player].name)}</option>`).join("")}</select><div class="control-row"><select id="moveZone" class="grow">${destinationOptions(ui.moveZone)}</select></div><div class="control-row"><select id="movePosition" class="grow"><option value="last" ${ui.position === "last" ? "selected" : ""}>下・末尾へ</option><option value="first" ${ui.position === "first" ? "selected" : ""}>上・先頭へ</option></select><button data-action="move-selected" class="primary">移動</button></div></div>${ids.length === 1 ? `<div class="control-group"><h3>このカードの一時メモ</h3><textarea id="cardNote" maxlength="2000" placeholder="効果や一時的な状態を記録">${escapeHTML(card.note)}</textarea><button class="small full" style="margin-top:5px" data-action="save-note">メモを反映</button></div>` : ""}`
}
function historyHTML() {
    return `<div class="history-toolbar"><button class="small full" data-action="clear-history" ${state.log.length ? "" : "disabled"}>履歴を初期化</button></div>${state.log.length ? [...state.log].reverse().map(item => `<div class="log-item"><time>${escapeHTML(new Date(item.time).toLocaleTimeString())}</time>${escapeHTML(item.label)}</div>`).join("") : `<div class="empty-help">履歴はありません。</div>`}`
}
function renderSideContent() {
    if (ui.tab === "operation") byId("sideContent").innerHTML = renderOperation()
    else if (ui.tab === "selection") byId("sideContent").innerHTML = mobileSelectionTools(true) + renderSelection()
    else byId("sideContent").innerHTML = historyHTML()
    syncMobileControls()
}

// SIDE_PANEL_UI_END
function refreshSelectionUI() {
    for (const element of document.querySelectorAll("[data-card-id]")) {
        const selected = ui.selection.has(element.dataset.cardId)
        element.classList.toggle("selected", selected)
        element.setAttribute("aria-pressed", String(selected))
    }
    renderSideHead()
    renderSideContent()
    if (byId("inspectSelectedCount")) byId("inspectSelectedCount").textContent = `${selectedIds().length}枚選択`
}
function render() {
    if (typeof renderLocalDeckEntry === "function" && renderLocalDeckEntry()) return
    if (appPage === "deck") {pruneCardImageURLs(); renderDeckPage(); return}
    pruneCardImageURLs()
    normalizePlayerView()
    ui.selection = new Set(selectedIds())
    if (!zoneById(ui.moveZone)) ui.moveZone = "hand"
    renderBoard()
    renderSideHead()
    renderSideContent()
    if (byId("modal").open && modalKind === "zone") renderZoneInspector()
    renderReplayControls()
    const replayMode = replay.mode === "replay"
    document.body?.classList.toggle("replay-mode", replayMode)
    for (const id of ["board", "phaseStrip", "sidePanel"]) if (byId(id)) byId(id).inert = replayMode
    for (const button of document.querySelectorAll(".header-actions button")) button.disabled = replayMode && !["save", "help"].includes(button.dataset.action)
}
function openModal(title, html, kind) {
    if (mobileUi.panelOpen) setMobilePanel(false, ui.tab, false)
    if (kind !== "preview") previewReturnContext = null
    byId("modalTitle").textContent = title
    byId("modalContent").innerHTML = html
    const back = byId("modalBack")
    back.hidden = kind !== "preview" || !previewReturnContext
    if (!back.hidden) back.textContent = previewReturnContext.limit === null ? "一覧に戻る" : "上から確認に戻る"
    modalKind = kind
    if (!byId("modal").open) byId("modal").showModal()
}
function closeModal(recordInspection = true) {
    if (appPage === "deck" && deckManagerPendingLeave) {finishDeckManagerLeave(false); return}
    if (modalKind === "deck-export") clearDeckExportPreview()
    byId("modal").close()
    modalKind = ""
    inspectContext = null
    previewCardId = null
    previewReturnContext = null
    replay.cutInEditor = null
    if (appPage === "deck") renderDeckPage(); else if (recordInspection) captureReplayFrame("山札の確認を終了")
}
function preview(id) {
    if (!Object.hasOwn(state.cards, id)) return
    previewReturnContext = modalKind === "zone" ? inspectContext : null
    previewCardId = id
    const item = definition(id)
    const card = state.cards[id]
    openModal("カードの内容を確認", `<div class="big-preview">${cardHTML(id, {reveal: true, preview: true})}<div><h2>${escapeHTML(item.name)}</h2>${cardTagsHTML(item)}<p class="description">${escapeHTML(item.text || "カードの説明は未登録です。")}</p>${card.note ? `<h3 style="margin-top:20px">このカードの一時メモ</h3><p class="description">${escapeHTML(card.note)}</p>` : ""}${where(id)?.zone === "deck" ? `<div style="margin-top:20px">${deckMoveControls(id)}</div>` : ""}</div></div>`, "preview")
    captureReplayFrame(where(id)?.zone === "deck" ? `${state.players[where(id).player].name}の山札のカードを確認` : "山札の確認を終了")
}
function returnToInspector() {
    if (modalKind !== "preview" || !previewReturnContext || inspectContext !== previewReturnContext) return false
    const {player, zone, limit} = inspectContext
    previewReturnContext = null
    previewCardId = null
    modalKind = "zone"
    byId("modalBack").hidden = true
    renderZoneInspector()
    if (zone === "deck") captureReplayFrame(`${state.players[player].name}の山札の${limit === null ? "一覧" : "上から確認"}に戻る`)
    return true
}
// ZONE_INSPECTOR_UI_START
const inspectorPageSize = 24
function currentInspectorPageSize() {
    return inspectContext?.zone === "deck" && !mobileViewport?.matches ? 50 : inspectorPageSize
}

function inspectZone(player, zone, limit = null) {
    const snapshot = limit === null ? null : state.locations[player][zone].slice(0, limit)
    inspectContext = {player, zone, limit, snapshot, originalCount: snapshot === null ? 0 : snapshot.length, page: 0}
    ui.selection.clear()
    openModal("ゾーンの内容を確認", "", "zone")
    captureReplayFrame(zone === "deck" ? `${state.players[player].name}の山札を${limit === null ? "すべて確認" : `上から${snapshot.length}枚確認`}` : "山札の確認を終了")
    render()
}
function deckMoveControls(sourceCard = "") {
    const source = sourceCard ? ` data-source-card="${sourceCard}"` : ""
    const player = sourceCard ? where(sourceCard).player : inspectContext.player
    return `<section class="control-group deck-move-controls"><h3>移動先</h3><div class="control-row"><select id="deckMoveDestination" class="grow" aria-label="移動先">${destinationOptions(ui.deckMoveZone, player, false, true)}</select></div><div class="control-row"><select id="deckMovePosition" class="grow" aria-label="移動する順番"><option value="last" ${ui.position === "last" ? "selected" : ""}>下・末尾へ</option><option value="first" ${ui.position === "first" ? "selected" : ""}>上・先頭へ</option></select><button data-action="deck-move"${source} class="primary">移動</button></div></section>`
}
function inspectorIds() {
    if (!inspectContext || !zoneById(inspectContext.zone)) return []
    const all = state.locations[inspectContext.player][inspectContext.zone]
    if (inspectContext.limit === null) return all
    const observed = new Set(inspectContext.snapshot)
    return all.filter(id => observed.has(id))
}
function visibleInspectorIds(ids = inspectorIds()) {
    if (!inspectContext) return []
    const pages = Math.max(1, Math.ceil(ids.length / currentInspectorPageSize()))
    inspectContext.page = integer(inspectContext.page, 0, pages - 1, 0)
    const start = inspectContext.page * currentInspectorPageSize()
    return ids.slice(start, start + currentInspectorPageSize())
}
function renderZoneInspector() {
    if (!inspectContext || !zoneById(inspectContext.zone)) return closeModal()
    const {player, zone, limit} = inspectContext
    const all = state.locations[player][zone]
    const ids = inspectorIds()
    const visible = visibleInspectorIds(ids)
    const start = inspectContext.page * currentInspectorPageSize()
    const pages = Math.max(1, Math.ceil(ids.length / currentInspectorPageSize()))
    const moves = zone === "deck" ? deckMoveControls() : `<select id="inspectDestination">${destinationOptions(ui.moveZone, player, zone === "colorline" || (ids.length > 0 && ids.every(id => colorForCard(id))))}</select><button data-action="inspect-move" data-position="last" class="primary">同じプレイヤーの移動先へ</button><button data-action="inspect-bottom">山札の下へ</button><button data-action="inspect-top">山札の上へ</button>`
    const summary = limit === null ? `全${all.length}枚` : `確認した${inspectContext.originalCount}枚のうち残り${ids.length}枚（ゾーン内${all.length}枚）`
    const range = ids.length ? `${start + 1}〜${start + visible.length}枚目を表示` : ""
    const rows = visible.map((id, index) => `<div class="inspect-item"><span class="position-label">${start + index + 1}${id === all[0] ? " / 先頭" : ""}</span>${cardHTML(id, {reveal: true, location: {player, zone}})}<div class="inspect-name">${escapeHTML(definition(id).name)}</div></div>`).join("")
    const deckDrop = zone === "deck" && limit === null ? ` data-drop-player="${player}" data-drop-zone="deck" data-inspector-drop="true"` : ""
    const deckOrder = zone === "deck" && limit === null ? `<div class="deck-order-controls"><label>山札内の位置（上から）<input id="inspectDeckPosition" type="number" min="1" max="${all.length}" value="1"></label><button data-action="inspect-reorder" class="small" ${all.length ? "" : "disabled"}>選択カードを移動</button></div>` : ""
    byId("modalTitle").textContent = `${state.players[player].name} / ${zoneById(zone).label}`
    byId("modalContent").innerHTML = `<div class="modal-toolbar"><span>${summary}</span><button class="small" data-action="inspect-select-all" ${visible.length ? "" : "disabled"}>表示中をすべて選択</button>${limit !== null ? `<button class="small" data-action="inspect-show-all">すべて見る</button>` : ""}</div>${mobileSelectionTools()}${zone === "colorline" ? `<p class="help-note">カラーカードはダブルクリックで表裏を反転します（この一覧では表面を表示します）。</p>` : ""}<div class="zone-inspect-grid"${deckDrop}>${rows || `<p class="muted">カードがありません。</p>`}</div><div class="zone-inspect-pagination"><span class="zone-inspect-range">${range}</span><button class="small" data-action="inspect-page-prev" ${inspectContext.page === 0 ? "disabled" : ""}>前へ</button><span>${inspectContext.page + 1} / ${pages}</span><button class="small" data-action="inspect-page-next" ${inspectContext.page + 1 >= pages ? "disabled" : ""}>次へ</button></div>${deckOrder}<div class="inspect-actions"><span id="inspectSelectedCount">${selectedIds().length}枚選択</span>${moves}</div>`
}
function handleInspectorAction(action, button) {
    if (!["inspect-page-prev", "inspect-page-next"].includes(action)) return false
    if (!inspectContext || modalKind !== "zone" || button?.disabled) return true
    visibleInspectorIds()
    inspectContext.page += action === "inspect-page-next" ? 1 : -1
    renderZoneInspector()
    return true
}

// ZONE_INSPECTOR_UI_END
function miniature(item) {
    return item.asset ? `<img class="mini-card" src="${cachedCardImageURL(item.asset)}" alt="" draggable="false" loading="lazy" decoding="async">` : `<span class="mini-card">◇</span>`
}
// CATALOG_UI_START
const catalogUi = {query: "", cardType: "", page: 0, editing: ""}
const catalogPageSize = 12
function filteredCatalogEntries() {
    const normalize = value => String(value).normalize("NFKC").toLocaleLowerCase("ja")
    const terms = normalize(catalogUi.query).trim().split(/\s+/).filter(Boolean)
    return Object.entries(state.catalog).filter(([, item]) => !isColorDefinition(item)
        && (!catalogUi.cardType || textCardType(item) === catalogUi.cardType)
        && terms.every(term => normalize(`${item.name} ${item.sourceCardKey || ""} ${liverCardTags(item).map(tag => `#${tag}`).join(" ")} ${item.text || ""}`).includes(term)))
}
function catalogResultsHTML() {
    const entries = filteredCatalogEntries()
    const pages = Math.max(1, Math.ceil(entries.length / catalogPageSize))
    catalogUi.page = integer(catalogUi.page, 0, pages - 1, 0)
    const start = catalogUi.page * catalogPageSize
    const visible = entries.slice(start, start + catalogPageSize)
    if (!visible.some(([id]) => id === catalogUi.editing)) catalogUi.editing = ""
    const usedDefinitions = new Set(Object.values(state.cards).map(card => card.definition))
    for (const player of playerIds) {
        for (const entry of state.decklists[player]) usedDefinitions.add(entry.definition)
        if (state.channelDefinitions[player]) usedDefinitions.add(state.channelDefinitions[player])
    }
    const rows = visible.map(([id, item]) => {
        const editing = catalogUi.editing === id
        const excerpt = Array.from(String(item.text || "").replace(/\s+/g, " ")).slice(0, 90).join("")
        return `<article class="catalog-entry"><div class="catalog-summary-row">${miniature(item)}<div class="catalog-card-info"><strong>${escapeHTML(item.name)}</strong><div class="catalog-card-summary">${escapeHTML(libraryCardSummary(item))}</div>${excerpt ? `<p class="catalog-effect-excerpt">${escapeHTML(excerpt)}</p>` : ""}</div><div class="catalog-row-actions"><button class="small" data-action="catalog-edit" data-definition="${id}" aria-expanded="${editing}">${editing ? "編集を閉じる" : "編集"}</button><button class="small" data-action="delete-definition" data-definition="${id}" ${usedDefinitions.has(id) || savedDeckUsesDefinition(id) || deckManagerUsesDefinition(id) ? "disabled" : ""}>削除</button></div></div>${editing ? `<div class="catalog-editor"><label>カード名<input type="text" maxlength="200" value="${escapeHTML(item.name)}" data-catalog-name="${id}" aria-label="カード名"></label><label>ID（任意）<input type="text" maxlength="100" value="${escapeHTML(item.sourceCardKey || "")}" data-catalog-id="${id}" aria-label="カードID" autocomplete="off"></label>${cardDetailFields(item, id)}<label>効果・説明<textarea maxlength="10000" data-catalog-text="${id}" aria-label="効果・説明" rows="5">${escapeHTML(item.text || "")}</textarea></label><p class="help-note">入力欄を離れると反映します。</p></div>` : ""}</article>`
    }).join("")
    const total = Object.values(state.catalog).filter(item => !isColorDefinition(item)).length
    return `<div class="catalog-summary" role="status">登録済み ${total}種類${entries.length ? ` · ${entries.length}件中 ${start + 1}〜${Math.min(start + catalogPageSize, entries.length)}件` : " · 該当するカードはありません"}</div><div class="catalog-list">${rows || `<p class="muted">${total ? "検索条件を変えてください。" : "画像・文字カードやExcelからカードを追加してください。"}</p>`}</div><div class="catalog-pagination"><button class="small" data-action="catalog-page" data-delta="-1" ${catalogUi.page === 0 ? "disabled" : ""}>前へ</button><span>${catalogUi.page + 1} / ${pages}</span><button class="small" data-action="catalog-page" data-delta="1" ${catalogUi.page + 1 >= pages ? "disabled" : ""}>次へ</button></div>`
}
function renderCatalogResults() {
    const results = byId("catalogResults")
    if (results) results.innerHTML = catalogResultsHTML()
}
function showCatalog() {
    const types = Object.entries(textCardTypes).map(([value, label]) => `<option value="${value}" ${catalogUi.cardType === value ? "selected" : ""}>${label}</option>`).join("")
    openModal("カード登録", `<div class="modal-toolbar"><button data-action="pick-images">カード画像を複数追加</button><span class="muted">PNG / JPEG / WebP · 1ファイル10MBまで</span></div><section class="catalog-new-card"><h3>画像なしの文字カードを追加</h3><div class="control-row"><input id="newCardName" type="text" maxlength="200" placeholder="カード名" class="grow"><button data-action="add-text-card">追加</button></div><div class="control-row"><label class="grow">ID（任意）<input id="newCardId" type="text" maxlength="100" placeholder="CARD-001" class="full" autocomplete="off"></label></div>${cardDetailFields()}<textarea id="newCardText" maxlength="10000" placeholder="効果・説明" rows="3"></textarea></section><div class="catalog-filters"><label>カード名・ID・タグ・効果<input id="catalogSearch" type="search" maxlength="200" value="${escapeHTML(catalogUi.query)}" placeholder="カード名・ID・タグ・効果で検索" autocomplete="off"></label><label>カード種別<select id="catalogType"><option value="">すべて</option>${types}</select></label></div><div id="catalogResults">${catalogResultsHTML()}</div><div class="modal-footer catalog-footer"><div class="catalog-import-actions"><button data-action="pick-card-library">Excelを取り込む</button>${cardLibrary.cards.length ? `<button data-action="card-library">読み取った一覧を開く</button>` : ""}</div><button data-action="catalog-return-decks" class="primary">${appPage === "deck" ? "デッキに戻る" : "デッキ管理へ"}</button><button data-action="close-modal">閉じる</button></div>`, "catalog")
}
function handleCatalogInput(element) {
    if (element.id !== "catalogSearch") return false
    catalogUi.query = element.value
    catalogUi.page = 0
    renderCatalogResults()
    return true
}
function handleCatalogChange(element) {
    if (element.id !== "catalogType") return false
    catalogUi.cardType = Object.hasOwn(textCardTypes, element.value) ? element.value : ""
    catalogUi.page = 0
    renderCatalogResults()
    return true
}
function handleCatalogAction(action, button) {
    if (action === "catalog-edit") {
        catalogUi.editing = catalogUi.editing === button.dataset.definition ? "" : button.dataset.definition
        const list = document.querySelector(".catalog-list")
        const scrollTop = list?.scrollTop || 0
        renderCatalogResults()
        const updated = document.querySelector(".catalog-list")
        if (updated) updated.scrollTop = scrollTop
        return true
    }
    if (action === "catalog-page") {
        if (!button.disabled) {
            catalogUi.page += Number(button.dataset.delta) === -1 ? -1 : 1
            catalogUi.editing = ""
            renderCatalogResults()
        }
        return true
    }
    return false
}

// CATALOG_UI_END
function parseCardId(raw, exceptDefinition = "") {
    const value = String(raw ?? "").trim()
    if (!value) return ""
    if (!/^[A-Za-z0-9_-]{1,100}$/.test(value)) throw new Error("IDは100文字以内の半角英数字・ハイフン・アンダーバーで入力してください。")
    if (["__proto__", "prototype", "constructor"].includes(value)) throw new Error("このIDは使用できません。別のIDを入力してください。")
    if (Object.entries(state.catalog).some(([id, item]) => id !== exceptDefinition && item.sourceCardKey === value)) throw new Error("同じIDのカードが既に登録されています。")
    return value
}

function updateCatalogCardId(definitionId, raw) {
    if (replay.mode === "replay" || !Object.hasOwn(state.catalog, definitionId) || isColorDefinition(state.catalog[definitionId])) return false
    try {
        const value = parseCardId(raw, definitionId)
        if ((state.catalog[definitionId].sourceCardKey || "") === value) return true
        return transact("カードのIDを変更", () => {
            if (value) state.catalog[definitionId].sourceCardKey = value
            else delete state.catalog[definitionId].sourceCardKey
        })
    } catch (error) {notify(error.message, true); return false}
}

function addTextCard() {
    const name = byId("newCardName").value.trim()
    const text = byId("newCardText").value
    if (!name) {notify("カード名を入力してください。", true); return false}
    try {
        const sourceCardKey = parseCardId(byId("newCardId").value)
        const cardType = parseCardDetail("cardType", byId("newCardType").value || "other")
        const details = {cardType, cost: cardType === "channel" ? null : parseCardDetail("cost", byId("newCardCost").value), className: parseCardDetail("className", byId("newCardClass").value), power: parseCardDetail("power", byId("newCardPower").value), colors: cardType === "liver" ? readCardDetailColors() : [], tags: cardType === "liver" ? parseCardDetail("tags", byId("newCardTags").value) : []}
        const added = transact("文字カードを登録", () => {state.catalog[uid("definition")] = {name, text, asset: "", ...details, ...(sourceCardKey ? {sourceCardKey} : {})}})
        if (added) showCatalog()
        return added
    } catch (error) {notify(error.message, true); return false}
}
// DECK_EDITOR_UI_START
let deckDrafts = null
let channelDrafts = null
let deckQuantityTouched = {}
const deckEditorUi = {query: "", cardType: "", className: "", sort: "registered", page: 0, currentPage: 0, editingQuantity: ""}
const deckEditorPageSize = 20
const deckEditorSorts = {registered: "登録順", name: "カード名", cost: "レベルが低い順", power: "パワーが高い順", type: "カード種別"}
const deckEditorCollator = new Intl.Collator("ja", {numeric: true, sensitivity: "base"})

function isDeckChannelDefinition(id) {
    return Object.hasOwn(state.catalog, id) && !isColorDefinition(state.catalog[id]) && textCardType(state.catalog[id]) === "channel"
}
function isDeckMainDefinition(id) {
    return Object.hasOwn(state.catalog, id) && !isColorDefinition(state.catalog[id]) && textCardType(state.catalog[id]) !== "channel"
}
function normalizeDeckDraft(player) {
    deckDrafts[player] = (deckDrafts[player] || []).filter(entry => entry.count > 0 && isDeckMainDefinition(entry.definition))
    if (!isDeckChannelDefinition(channelDrafts[player])) channelDrafts[player] = ""
}
function beginDeckEditor(player = state.active) {
    deckDrafts = copy(state.decklists)
    channelDrafts = copy(state.channelDefinitions)
    deckQuantityTouched = Object.fromEntries(playerIds.map(id => [id, new Set()]))
    for (const id of playerIds) normalizeDeckDraft(id)
    Object.assign(deckEditorUi, {query: "", cardType: "", className: "", sort: "registered", page: 0, currentPage: 0, editingQuantity: ""})
    editorPlayer = visiblePlayerIds().includes(player) ? player : "p1"
    showDeckEditor()
}
function setDeckDraftQuantity(definition, rawValue, keepZero = false) {
    if (!deckDrafts) return false
    if (!isDeckMainDefinition(definition)) {
        deckDrafts[editorPlayer] = (deckDrafts[editorPlayer] || []).filter(entry => entry.definition !== definition)
        return false
    }
    const count = integer(rawValue, 0, 500, 0)
    const list = deckDrafts[editorPlayer] || (deckDrafts[editorPlayer] = [])
    const index = list.findIndex(item => item.definition === definition)
    if (index >= 0) {
        if (count > 0 || keepZero) list[index].count = count
        else list.splice(index, 1)
    } else if (count > 0) list.push({definition, count})
    return true
}
function collectDeckDraft() {
    if (!deckDrafts) return
    // Input events own edited quantities. Old copies of an input must not overwrite them.
    for (const element of document.querySelectorAll("[data-deck-quantity]")) {
        if (!deckQuantityTouched[editorPlayer]?.has(element.dataset.deckQuantity)) setDeckDraftQuantity(element.dataset.deckQuantity, element.value)
    }
    const select = byId("channelDefinition")
    if (select) {
        const value = select.value
        channelDrafts[editorPlayer] = isDeckChannelDefinition(value) ? value : ""
    }
    deckEditorUi.editingQuantity = ""
    normalizeDeckDraft(editorPlayer)
}
function deckTotal(renderCurrent = true) {
    const entries = (deckDrafts?.[editorPlayer] || []).filter(entry => entry.count > 0 && isDeckMainDefinition(entry.definition))
    const total = entries.reduce((sum, item) => sum + item.count, 0)
    const counter = byId("deckTotal")
    if (counter) counter.textContent = `山札 ${total}枚`
    const summary = byId("deckCurrentSummary")
    if (summary) summary.textContent = `山札 ${total}枚 · ${entries.length}種類`
    if (renderCurrent) renderDeckCurrentContents()
    return total
}
function deckCurrentContentsHTML() {
    const channel = channelDrafts?.[editorPlayer] || ""
    const entries = (deckDrafts?.[editorPlayer] || []).filter(entry => isDeckMainDefinition(entry.definition)
        && (entry.count > 0 || entry.definition === deckEditorUi.editingQuantity))
    const total = entries.reduce((sum, entry) => sum + entry.count, 0)
    const kinds = entries.filter(entry => entry.count > 0).length
    const pages = Math.max(1, Math.ceil(entries.length / deckEditorPageSize))
    deckEditorUi.currentPage = integer(deckEditorUi.currentPage, 0, pages - 1, 0)
    const start = deckEditorUi.currentPage * deckEditorPageSize
    const rows = entries.slice(start, start + deckEditorPageSize).map(entry => {
        const item = state.catalog[entry.definition]
        return `<div class="deck-current-row">${miniature(item)}<div class="deck-current-info"><strong>${escapeHTML(item.name)}</strong>${item.sourceCardKey ? `<small>${escapeHTML(item.sourceCardKey)}</small>` : ""}</div><label class="deck-current-quantity">枚数<input data-current-deck-quantity="${entry.definition}" aria-label="${escapeHTML(item.name)}${item.sourceCardKey ? ` (${escapeHTML(item.sourceCardKey)})` : ""}の枚数（いまのデッキ）" type="number" min="0" max="500" step="1" value="${entry.count}"></label></div>`
    }).join("")
    const channelItem = isDeckChannelDefinition(channel) ? state.catalog[channel] : null
    return `<div id="deckCurrentSummary" class="deck-current-summary" role="status">山札 ${total}枚 · ${kinds}種類</div><div class="deck-current-channel"><span>チャンネル</span><strong>${channelItem ? escapeHTML(channelItem.name) : "未設定"}</strong>${channelItem?.sourceCardKey ? `<small>${escapeHTML(channelItem.sourceCardKey)}</small>` : ""}</div><div class="deck-current-list">${rows || `<p class="muted">まだカードが入っていません。カード一覧で枚数を指定してください。</p>`}</div>${pages > 1 ? `<div class="deck-current-pagination"><button class="small" data-action="deck-current-prev" ${deckEditorUi.currentPage === 0 ? "disabled" : ""}>前へ</button><span>${deckEditorUi.currentPage + 1} / ${pages}</span><button class="small" data-action="deck-current-next" ${deckEditorUi.currentPage + 1 >= pages ? "disabled" : ""}>次へ</button></div>` : ""}`
}
function renderDeckCurrentContents() {
    const contents = byId("deckCurrentContents")
    if (!contents) return
    const scrollTop = document.querySelector(".deck-current-list")?.scrollTop || 0
    contents.innerHTML = deckCurrentContentsHTML()
    const list = document.querySelector(".deck-current-list")
    if (list) list.scrollTop = scrollTop
}
function syncDeckQuantityInputs(definition, value, source) {
    const inputs = [...document.querySelectorAll("[data-deck-quantity]"), ...document.querySelectorAll("[data-current-deck-quantity]")]
    for (const input of inputs) {
        if (input !== source && (input.dataset.deckQuantity || input.dataset.currentDeckQuantity) === definition) input.value = String(value)
    }
}
function updateDeckQuantityInput(element, commit = false) {
    const definition = element.dataset.deckQuantity || element.dataset.currentDeckQuantity
    if (!definition) return false
    const current = !!element.dataset.currentDeckQuantity
    deckEditorUi.editingQuantity = current && !commit ? definition : ""
    setDeckDraftQuantity(definition, element.value, !commit)
    if (!deckQuantityTouched[editorPlayer]) deckQuantityTouched[editorPlayer] = new Set()
    deckQuantityTouched[editorPlayer].add(definition)
    const count = (deckDrafts?.[editorPlayer] || []).find(entry => entry.definition === definition)?.count || 0
    if (commit) element.value = String(count)
    syncDeckQuantityInputs(definition, element.value, element)
    // A current-deck input keeps its DOM while typing and while using its spinner.
    deckTotal(!current || (commit && count === 0))
    return true
}
function filteredDeckEntries() {
    const normalize = value => String(value).normalize("NFKC").toLocaleLowerCase("ja")
    const terms = normalize(deckEditorUi.query).trim().split(/\s+/).filter(Boolean)
    const entries = Object.entries(state.catalog).filter(([id, item]) => isDeckMainDefinition(id)
        && (!deckEditorUi.cardType || textCardType(item) === deckEditorUi.cardType)
        && (!deckEditorUi.className || item.className === deckEditorUi.className)
        && terms.every(term => normalize(`${item.name} ${item.sourceCardKey || ""} ${liverCardTags(item).map(tag => `#${tag}`).join(" ")} ${item.text || ""}`).includes(term)))
    if (deckEditorUi.sort === "registered") return entries
    const types = Object.keys(textCardTypes)
    const valueForSort = item => deckEditorUi.sort === "cost" ? (textCardType(item) === "channel" ? null : item.cost)
        : deckEditorUi.sort === "power" ? (textCardType(item) === "liver" ? item.power : null) : types.indexOf(textCardType(item))
    return entries.sort((left, right) => {
        if (deckEditorUi.sort === "name") return deckEditorCollator.compare(left[1].name, right[1].name)
        const a = valueForSort(left[1]), b = valueForSort(right[1])
        const aMissing = a === undefined || a === null, bMissing = b === undefined || b === null
        if (aMissing || bMissing) return aMissing === bMissing ? 0 : aMissing ? 1 : -1
        return deckEditorUi.sort === "power" ? b - a : a - b
    })
}
function deckEditorResultsHTML() {
    const entries = filteredDeckEntries()
    const pages = Math.max(1, Math.ceil(entries.length / deckEditorPageSize))
    deckEditorUi.page = integer(deckEditorUi.page, 0, pages - 1, 0)
    const start = deckEditorUi.page * deckEditorPageSize
    const quantities = Object.fromEntries((deckDrafts?.[editorPlayer] || []).map(item => [item.definition, item.count]))
    const rows = entries.slice(start, start + deckEditorPageSize).map(([id, item]) => {
        const effect = Array.from(String(item.text || "").replace(/\s+/g, " ").trim()).slice(0, 75).join("")
        return `<div class="deck-editor-row">${miniature(item)}<div class="deck-editor-card-info"><strong>${escapeHTML(item.name)}</strong><div class="deck-editor-card-summary">${escapeHTML(libraryCardSummary(item))}</div>${effect ? `<div class="deck-editor-effect">${escapeHTML(effect)}</div>` : ""}</div><label>枚数<input data-deck-quantity="${id}" aria-label="${escapeHTML(item.name)}${item.sourceCardKey ? ` (${escapeHTML(item.sourceCardKey)})` : ""}の枚数" type="number" min="0" max="500" step="1" value="${quantities[id] || 0}"></label></div>`
    }).join("")
    const summary = entries.length ? `${entries.length}種類中 ${start + 1}〜${Math.min(start + deckEditorPageSize, entries.length)}種類` : "0種類"
    return `<div class="deck-editor-summary">${summary}</div><div class="deck-editor-list">${rows || `<p class="muted">${Object.values(state.catalog).some(item => !isColorDefinition(item)) ? "条件に一致するカードがありません。" : "カード登録から画像または文字カードを追加してください。"}</p>`}</div><div class="deck-editor-pagination"><button class="small" data-action="deck-editor-prev" ${deckEditorUi.page === 0 ? "disabled" : ""}>前へ</button><span>${deckEditorUi.page + 1} / ${pages}</span><button class="small" data-action="deck-editor-next" ${deckEditorUi.page + 1 >= pages ? "disabled" : ""}>次へ</button></div>`
}
function renderDeckEditorResults() {
    const results = byId("deckEditorResults")
    if (results) results.innerHTML = deckEditorResultsHTML()
    deckTotal()
}
function showDeckEditor() {
    if (!deckDrafts || !channelDrafts) return beginDeckEditor()
    if (!state.showOpponent) editorPlayer = "p1"
    const channel = channelDrafts[editorPlayer] || ""
    const entries = Object.entries(state.catalog).filter(([, item]) => !isColorDefinition(item))
    const channelOptions = entries.filter(([id]) => isDeckChannelDefinition(id)).map(([id, item]) => `<option value="${id}" ${id === channel ? "selected" : ""}>${escapeHTML(item.name)}${item.sourceCardKey ? ` (${escapeHTML(item.sourceCardKey)})` : ""}</option>`).join("")
    const typeOptions = Object.entries(textCardTypes).filter(([value]) => value !== "channel").map(([value, label]) => `<option value="${value}" ${deckEditorUi.cardType === value ? "selected" : ""}>${label}</option>`).join("")
    const classOptions = textCardClasses.map(value => `<option value="${value}" ${deckEditorUi.className === value ? "selected" : ""}>${value}</option>`).join("")
    const sortOptions = Object.entries(deckEditorSorts).map(([value, label]) => `<option value="${value}" ${deckEditorUi.sort === value ? "selected" : ""}>${label}</option>`).join("")
    openModal("デッキ編集", `<div class="modal-toolbar"><label>編集する側<select id="editorPlayer">${visiblePlayerIds().map(player => `<option value="${player}" ${player === editorPlayer ? "selected" : ""}>${escapeHTML(state.players[player].name)}</option>`).join("")}</select></label><strong id="deckTotal"></strong>${state.showOpponent ? `<button data-action="deck-copy-other" class="small">この入力をもう片側へコピー</button>` : ""}</div><div class="control-group"><label for="channelDefinition">チャンネル（1枚）</label><select id="channelDefinition" class="full"><option value="" ${channel ? "" : "selected"}>未設定</option>${channelOptions}</select><p class="help-note">選んだカードをチャンネルに1枚置きます。このカードは山札の枚数に含めません。</p></div><div class="deck-editor-workspace"><div class="deck-editor-browser"><h3>カードを探す</h3><div class="deck-editor-filters"><label>カード名・ID・タグ・効果<input id="deckEditorSearch" type="search" value="${escapeHTML(deckEditorUi.query)}" placeholder="名前やIDで検索" autocomplete="off"></label><label>カード種別<select id="deckEditorType"><option value="">すべて</option>${typeOptions}</select></label><label>クラス<select id="deckEditorClass"><option value="">すべて</option>${classOptions}</select></label><label>並び順<select id="deckEditorSort">${sortOptions}</select></label></div><div id="deckEditorResults">${deckEditorResultsHTML()}</div></div><aside class="deck-editor-current" aria-label="いまのデッキ"><h3>いまのデッキ</h3><div id="deckCurrentContents">${deckCurrentContentsHTML()}</div></aside></div><p class="help-note">公式のデッキ枚数・同名枚数・構築条件は判定しません。デッキは各500枚までです。</p><div class="modal-footer"><button data-action="catalog">カード登録</button><button data-action="deck-export">画像出力</button><button data-action="apply-deck" class="primary">デッキを保存</button>${state.showOpponent ? `<button data-action="apply-both-decks">両方のデッキを保存</button>` : ""}</div><p class="help-note">反映した側の通常カードを置き換え、山札とチャンネルを作り直します。カラーカードは現在の位置に保持します。シャッフルと初期手札は基本操作から指定してください。</p>`, "deck-editor")
    deckTotal()
}
function handleDeckEditorInput(element) {
    if (element.id === "deckEditorSearch") {
        collectDeckDraft()
        deckEditorUi.query = element.value
        deckEditorUi.page = 0
        renderDeckEditorResults()
        return true
    }
    return updateDeckQuantityInput(element)
}
function handleDeckEditorChange(element) {
    if (updateDeckQuantityInput(element, true)) return true
    if (element.id === "editorPlayer") {
        collectDeckDraft()
        editorPlayer = visiblePlayerIds().includes(element.value) ? element.value : "p1"
        deckEditorUi.page = 0
        deckEditorUi.currentPage = 0
        showDeckEditor()
        return true
    }
    if (element.id === "channelDefinition") {
        collectDeckDraft()
        deckEditorUi.page = 0
        renderDeckEditorResults()
        return true
    }
    const field = {deckEditorType: "cardType", deckEditorClass: "className", deckEditorSort: "sort"}[element.id]
    if (field) {
        collectDeckDraft()
        const allowed = field === "cardType" ? ["", ...Object.keys(textCardTypes).filter(type => type !== "channel")] : field === "className" ? ["", ...textCardClasses] : Object.keys(deckEditorSorts)
        deckEditorUi[field] = allowed.includes(element.value) ? element.value : field === "sort" ? "registered" : ""
        deckEditorUi.page = 0
        renderDeckEditorResults()
        return true
    }
    return false
}
function handleDeckEditorAction(action, button) {
    if (action === "deck-export") {showDeckExportPreview(button); return true}
    if (action === "deck-export-save") {saveDeckExportImage(); return true}
    if (action === "deck-export-back") {clearDeckExportPreview(); showDeckEditor(); return true}
    if (["deck-current-prev", "deck-current-next"].includes(action)) {
        if (button?.disabled) return true
        collectDeckDraft()
        deckEditorUi.currentPage += action === "deck-current-next" ? 1 : -1
        renderDeckCurrentContents()
        const list = document.querySelector(".deck-current-list")
        if (list) list.scrollTop = 0
        return true
    }
    if (!["deck-editor-prev", "deck-editor-next"].includes(action)) return false
    if (button?.disabled) return true
    collectDeckDraft()
    deckEditorUi.page += action === "deck-editor-next" ? 1 : -1
    renderDeckEditorResults()
    return true
}

// DECK_EDITOR_UI_END
// DECK_IMAGE_EXPORT_START
let deckExportPreviewURL = ""
let deckExportFileName = ""
const deckExportColors = {R: "#ed2946", O: "#ff8918", Y: "#f3c900", G: "#12ae50", B: "#149bea", I: "#3f49dd", V: "#9c35dc"}

function deckExportBox(context, x, y, width, height, radius, fill, stroke = "") {
    const r = Math.min(radius, width / 2, height / 2)
    context.beginPath()
    context.moveTo(x + r, y)
    context.arcTo(x + width, y, x + width, y + height, r)
    context.arcTo(x + width, y + height, x, y + height, r)
    context.arcTo(x, y + height, x, y, r)
    context.arcTo(x, y, x + width, y, r)
    context.closePath()
    context.fillStyle = fill
    context.fill()
    if (stroke) {
        context.strokeStyle = stroke
        context.lineWidth = 2
        context.stroke()
    }
}

function deckExportLines(context, value, width, maxLines = 2) {
    const characters = Array.from(String(value ?? "").replace(/\s+/g, " ").trim() || "未設定")
    const lines = []
    let line = ""
    for (const character of characters) {
        if (line && context.measureText(line + character).width > width) {
            lines.push(line)
            line = character
        } else line += character
    }
    if (line) lines.push(line)
    if (lines.length <= maxLines) return lines
    const visible = lines.slice(0, maxLines)
    let last = visible[maxLines - 1]
    while (last && context.measureText(last + "…").width > width) last = Array.from(last).slice(0, -1).join("")
    visible[maxLines - 1] = last + "…"
    return visible
}

function deckExportText(context, value, x, y, width, lineHeight, maxLines = 2) {
    const lines = deckExportLines(context, value, width, maxLines)
    for (const [index, line] of lines.entries()) context.fillText(line, x, y + index * lineHeight)
    return lines.length
}

function deckExportBadge(context, x, y, count) {
    context.save()
    context.font = "800 23px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    const label = `×${count}`
    const width = Math.max(53, Math.ceil(context.measureText(label).width) + 22)
    deckExportBox(context, x, y, width, 36, 18, "#213e55")
    context.fillStyle = "#fff"
    context.textAlign = "center"
    context.fillText(label, x + width / 2, y + 26)
    context.restore()
}

function deckExportColorBands(context, item, x, y, height) {
    const colors = liverCardColors(item)
    if (!colors.length) return
    const segment = (height - 61) / 7
    for (const color of colors) {
        const bandY = y + 38 + nijiColors.indexOf(color) * segment
        context.fillStyle = deckExportColors[color]
        context.fillRect(x + 7, bandY, 10, Math.max(7, segment - 3))
        context.strokeStyle = "#354c62"
        context.lineWidth = 1
        context.strokeRect(x + 7, bandY, 10, Math.max(7, segment - 3))
    }
}

function deckExportTypeSideLabel(context, item, x, y, height) {
    const label = cardTypeSideLabel(item)
    if (!label) return
    const top = y + 43, labelHeight = Math.max(0, height - 59)
    context.save()
    deckExportBox(context, x + 7, top, 12, labelHeight, 2, "#e3edf5")
    context.translate(x + 13, top + labelHeight / 2)
    context.rotate(Math.PI / 2)
    context.fillStyle = "#31536e"
    context.font = "750 11px system-ui, sans-serif"
    context.textAlign = "center"
    context.textBaseline = "middle"
    context.fillText(label, 0, 0, labelHeight)
    context.restore()
}
function deckExportTextCard(context, item, x, y, width, height) {
    const channel = textCardType(item) === "channel"
    const hasSideMark = liverCardColors(item).length > 0 || Boolean(cardTypeSideLabel(item))
    context.save()
    const gradient = context.createLinearGradient(x, y, x + width, y + height)
    gradient.addColorStop(0, "#fafdff")
    gradient.addColorStop(1, channel ? "#e9eef8" : "#e7f1f8")
    deckExportBox(context, x, y, width, height, 8, gradient, "#bacbd9")
    deckExportColorBands(context, item, x, y, height)
    deckExportTypeSideLabel(context, item, x, y, height)
    if (!channel && item.cost !== undefined && item.cost !== null) {
        deckExportBox(context, x + 8, y + 8, 36, 30, 6, "#e1eef8", "#b7ccdc")
        context.fillStyle = "#325b79"
        context.font = "800 19px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        context.textAlign = "center"
        context.fillText(String(item.cost), x + 26, y + 30, 31)
    }
    if (item.className) {
        deckExportBox(context, x + width - 42, y + 8, 34, 30, 6, "#f1ebf9", "#d3c6e0")
        context.fillStyle = "#695982"
        context.font = "800 19px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        context.textAlign = "center"
        context.fillText(Array.from(item.className)[0], x + width - 25, y + 30)
    }
    context.fillStyle = "#b8cfde"
    context.font = channel ? "700 63px system-ui, sans-serif" : "700 49px system-ui, sans-serif"
    context.textAlign = "center"
    context.fillText("◇", x + width / 2, y + height * .49)
    context.fillStyle = "#31536e"
    context.font = channel ? "750 29px system-ui, 'Yu Gothic', Meiryo, sans-serif" : "750 20px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    const lines = deckExportLines(context, item.name, width - (hasSideMark ? 46 : 30), channel ? 2 : 3)
    const lineHeight = channel ? 36 : 25
    const first = y + height * .59 - (lines.length - 1) * lineHeight / 2
    for (const [index, line] of lines.entries()) context.fillText(line, x + width / 2 + (hasSideMark ? 6 : 0), first + index * lineHeight)
    if (textCardType(item) === "liver" && item.power !== undefined && item.power !== null) {
        context.font = "800 17px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        const label = String(item.power)
        const badgeWidth = Math.max(39, Math.ceil(context.measureText(label).width) + 16)
        deckExportBox(context, x + width - badgeWidth - 7, y + height - 30, badgeWidth, 24, 5, "#fff4db", "#dfcda7")
        context.fillStyle = "#856a36"
        context.fillText(label, x + width - badgeWidth / 2 - 7, y + height - 12)
    }
    context.restore()
}

function deckExportLoadImage(assetId) {
    return new Promise(resolve => {
        const source = cachedCardImageURL(assetId)
        if (!source) return resolve(null)
        const image = new Image()
        image.onload = () => resolve(image)
        image.onerror = () => resolve(null)
        image.src = source
    })
}

async function deckExportCard(context, item, x, y, width, height, count) {
    const image = item.asset ? await deckExportLoadImage(item.asset) : null
    if (image?.naturalWidth && image?.naturalHeight) {
        deckExportBox(context, x, y, width, height, 8, "#fff", "#bacbd9")
        const scale = Math.min((width - 4) / image.naturalWidth, (height - 4) / image.naturalHeight)
        const imageWidth = image.naturalWidth * scale, imageHeight = image.naturalHeight * scale
        context.drawImage(image, x + (width - imageWidth) / 2, y + (height - imageHeight) / 2, imageWidth, imageHeight)
        deckExportColorBands(context, item, x, y, height)
        deckExportTypeSideLabel(context, item, x, y, height)
    } else deckExportTextCard(context, item, x, y, width, height)
    deckExportBadge(context, x + (liverCardColors(item).length || cardTypeSideLabel(item) ? 24 : 8), y + height - 44, count)
}

async function buildDeckExportCanvas(player, entries, channel, titleOverride = "") {
    const width = 1800, margin = 54, visualWidth = 1162, listX = 1246, listWidth = 500
    const gridX = 79, gridY = 509, columns = 7, cardWidth = 140, cardHeight = 196, columnStep = 162, rowStep = 226
    const listY = 490, dense = entries.length > 120
    const measure = document.createElement("canvas").getContext("2d")
    const rows = entries.map(({item}) => {
        measure.font = `${dense ? 15 : 18}px system-ui, 'Yu Gothic', Meiryo, sans-serif`
        const nameLines = deckExportLines(measure, item.name, 316, dense ? 1 : 2)
        measure.font = `${dense ? 12 : 14}px system-ui, 'Yu Gothic', Meiryo, sans-serif`
        const id = item.sourceCardKey ? `ID: ${item.sourceCardKey}` : "ID未設定"
        const idLines = dense ? [id] : deckExportLines(measure, id, 316, 4)
        return {nameLines, idLines, height: dense ? 50 : Math.max(60, nameLines.length * 23 + idLines.length * 18 + 16)}
    })
    const listHeight = rows.reduce((sum, row) => sum + row.height, 0)
    const gridRows = Math.ceil(entries.length / columns)
    const height = Math.max(1020, gridY + Math.max(1, gridRows) * rowStep + 70, listY + listHeight + 80)
    const canvas = document.createElement("canvas")
    canvas.width = width
    canvas.height = height
    const context = canvas.getContext("2d")
    if (!context) throw new Error("画像の描画を開始できませんでした。")
    context.fillStyle = "#edf3f8"
    context.fillRect(0, 0, width, height)
    context.fillStyle = "#23435d"
    context.fillRect(0, 0, width, 129)
    context.fillStyle = "#fff"
    context.font = "800 46px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    context.fillText(titleOverride || "デッキリスト", margin, 67, 1100)
    context.font = "600 24px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    context.fillStyle = "#c8deed"
    context.fillText("NijinanaSimulator", margin + 2, 105)
    context.textAlign = "right"
    context.font = "700 27px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    context.fillStyle = "#fff"
    if (!titleOverride) context.fillText(state.players[player].name, width - margin, 61, 560)
    context.font = "500 20px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    context.fillStyle = "#c8deed"
    context.fillText(`山札 ${entries.reduce((sum, entry) => sum + entry.count, 0)}枚 · ${entries.length}種類`, width - margin, 101)
    context.textAlign = "left"
    for (const [index, color] of nijiColors.entries()) {
        context.fillStyle = deckExportColors[color]
        context.fillRect(index * width / 7, 124, Math.ceil(width / 7), 5)
    }
    deckExportBox(context, margin, 158, visualWidth, 247, 16, "#fff", "#d5e2eb")
    deckExportBox(context, margin, 423, visualWidth, height - 483, 16, "#fff", "#d5e2eb")
    deckExportBox(context, listX, 158, listWidth, height - 218, 16, "#fff", "#d5e2eb")
    context.fillStyle = "#53738b"
    context.font = "800 19px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    context.fillText("CHANNEL / チャンネル", margin + 30, 196)
    context.fillText("DECK / 山札", margin + 30, 464)
    context.fillText("CARD LIST / カード一覧", listX + 27, 200)
    context.font = "700 20px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    context.fillStyle = "#355773"
    context.fillText("チャンネル", listX + 27, 249)
    context.fillText(`山札　${entries.length}種類`, listX + 27, 451)
    context.strokeStyle = "#dce6ee"
    context.lineWidth = 2
    context.beginPath()
    context.moveTo(listX + 27, 416)
    context.lineTo(listX + listWidth - 27, 416)
    context.stroke()
    if (channel) {
        await deckExportCard(context, channel, margin + 31, 213, 278, 176, 1)
        context.fillStyle = "#31536e"
        context.font = "800 31px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        deckExportText(context, channel.name, margin + 348, 274, 755, 41, 2)
        if (channel.sourceCardKey) {
            context.fillStyle = "#8096a6"
            context.font = "500 18px system-ui, 'Yu Gothic', Meiryo, sans-serif"
            context.fillText(channel.sourceCardKey, margin + 350, 374, 740)
        }
        context.fillStyle = "#31536e"
        context.font = "700 22px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        deckExportText(context, channel.name, listX + 28, 307, 350, 28, 2)
        context.fillStyle = "#788fa2"
        context.font = "500 14px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        deckExportText(context, channel.sourceCardKey ? `ID: ${channel.sourceCardKey}` : "ID未設定", listX + 28, 362, 410, 17, 3)
        deckExportBadge(context, listX + listWidth - 92, 279, 1)
    } else {
        deckExportBox(context, margin + 31, 213, 278, 176, 8, "#f3f7fa", "#d5e2eb")
        context.fillStyle = "#899dad"
        context.font = "700 26px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        context.fillText("チャンネル未設定", margin + 350, 316)
        context.font = "600 20px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        context.fillText("未設定", listX + 28, 308)
    }
    if (!entries.length) {
        context.fillStyle = "#899dad"
        context.font = "600 24px system-ui, 'Yu Gothic', Meiryo, sans-serif"
        context.fillText("山札にカードがありません", margin + 30, 563)
    }
    for (const [index, entry] of entries.entries()) {
        const x = gridX + index % columns * columnStep
        const y = gridY + Math.floor(index / columns) * rowStep
        await deckExportCard(context, entry.item, x, y, cardWidth, cardHeight, entry.count)
        context.fillStyle = "#8196a7"
        context.font = "600 15px system-ui, sans-serif"
        context.textAlign = "right"
        context.fillText(String(index + 1).padStart(2, "0"), x + cardWidth - 2, y + cardHeight + 19)
        context.textAlign = "left"
    }
    let rowY = listY
    for (const [index, entry] of entries.entries()) {
        const row = rows[index]
        if (index % 2 === 0) deckExportBox(context, listX + 18, rowY, listWidth - 36, row.height, 6, "#f4f8fb")
        context.fillStyle = "#8a9faf"
        context.font = "700 15px system-ui, sans-serif"
        context.fillText(String(index + 1).padStart(2, "0"), listX + 29, rowY + (row.height + 11) / 2)
        context.fillStyle = "#344f65"
        context.font = `650 ${dense ? 15 : 18}px system-ui, 'Yu Gothic', Meiryo, sans-serif`
        const nameY = rowY + (dense ? 20 : 26)
        for (const [lineIndex, line] of row.nameLines.entries()) context.fillText(line, listX + 69, nameY + lineIndex * (dense ? 18 : 23))
        context.fillStyle = "#7b91a2"
        context.font = `500 ${dense ? 12 : 14}px system-ui, 'Yu Gothic', Meiryo, sans-serif`
        const idY = rowY + (dense ? 39 : 26 + row.nameLines.length * 23 - 4)
        for (const [lineIndex, line] of row.idLines.entries()) context.fillText(line, listX + 69, idY + lineIndex * (dense ? 15 : 18), 316)
        context.font = "800 19px system-ui, sans-serif"
        context.textAlign = "right"
        context.fillStyle = "#315d7d"
        context.fillText(`×${entry.count}`, listX + listWidth - 28, rowY + (row.height + 13) / 2)
        context.textAlign = "left"
        rowY += row.height
    }
    context.fillStyle = "#7c92a4"
    context.font = "500 17px system-ui, 'Yu Gothic', Meiryo, sans-serif"
    context.fillText("NijinanaSimulator", margin, height - 27)
    return canvas
}

function clearDeckExportPreview() {
    if (deckExportPreviewURL) URL.revokeObjectURL(deckExportPreviewURL)
    deckExportPreviewURL = ""
    deckExportFileName = ""
}

async function showDeckExportPreview(button) {
    if (modalKind !== "deck-editor" || button.disabled) return
    collectDeckDraft()
    const player = editorPlayer
    const entries = (deckDrafts[player] || []).filter(entry => entry.count > 0 && isDeckMainDefinition(entry.definition))
        .map(entry => ({item: state.catalog[entry.definition], count: entry.count}))
    const channel = isDeckChannelDefinition(channelDrafts[player]) ? state.catalog[channelDrafts[player]] : null
    if (!channel && !entries.length) return notify("チャンネルか山札のカードを設定してください。", true)
    if (entries.reduce((sum, entry) => sum + entry.count, 0) > 500) return notify("デッキは500枚までです。", true)
    const original = button.textContent
    button.disabled = true
    button.textContent = "画像を作成中…"
    try {
        const canvas = await buildDeckExportCanvas(player, entries, channel)
        const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/png"))
        if (!blob) throw new Error("PNG画像を作成できませんでした。")
        if (modalKind !== "deck-editor" || !byId("modal").open || !button.isConnected) return
        clearDeckExportPreview()
        deckExportPreviewURL = URL.createObjectURL(blob)
        deckExportFileName = `NijinanaSimulator_deck_${player}_${new Date().toISOString().slice(0, 19).replaceAll(":", "-")}.png`
        openModal("デッキ画像の確認", `<p class="help-note">編集中のデッキを画像にしました。内容を確認してPNGで保存できます。</p><div class="deck-export-preview"><img src="${deckExportPreviewURL}" alt="${escapeHTML(state.players[player].name)}のデッキ画像"></div><div class="modal-footer"><button data-action="deck-export-back">デッキ編集に戻る</button><button data-action="deck-export-save" class="primary">PNG画像を保存</button></div>`, "deck-export")
    } catch (error) {
        console.error("Deck image export failed", error)
        notify(`デッキ画像を作成できませんでした。 ${error.message}`, true)
    } finally {
        if (button.isConnected) {
            button.disabled = false
            button.textContent = original
        }
    }
}

function saveDeckExportImage() {
    if (!deckExportPreviewURL) return notify("デッキ画像がありません。", true)
    const anchor = document.createElement("a")
    anchor.href = deckExportPreviewURL
    anchor.download = deckExportFileName
    anchor.click()
    notify("PNG画像を保存します。ブラウザの保存先を確認してください。")
}
// DECK_IMAGE_EXPORT_END
function applyDeck(both = false) {
    collectDeckDraft()
    const targets = both ? visiblePlayerIds() : [state.showOpponent ? editorPlayer : "p1"]
    for (const player of targets) {
        const count = deckDrafts[player].reduce((sum, item) => sum + item.count, 0)
        if (count > 500) return notify("デッキは各500枚までです。", true)
    }
    if (!transact("デッキ編集の内容から山札とチャンネルを設定", () => {
        for (const player of targets) createDeckRaw(player, deckDrafts[player], channelDrafts[player])
    })) return
    notify("デッキを保存し、山札とチャンネルを更新しました。")
}
function showSettings() {
    const playerSettings = visiblePlayerIds().map(player => `<section><label>プレイヤー名<input id="settingsName_${player}" maxlength="60" value="${escapeHTML(state.players[player].name)}"></label>${state.players[player].meters.map((meter, index) => `<label>カウンター ${index + 1} の名前<input id="settingsMeter_${player}_${index}" maxlength="50" value="${escapeHTML(meter.label)}"></label>`).join("")}</section>`).join("")
    const rows = state.zones.map(zone => {
        const fixedMode = ["deck", "hand"].includes(zone.id)
        const protectedZone = defaultZones().some(item => item.id === zone.id)
        return `<div class="settings-row"><input data-zone-label="${zone.id}" maxlength="80" value="${escapeHTML(zone.label)}" aria-label="ゾーン名"><select data-zone-mode="${zone.id}" ${fixedMode ? "disabled" : ""}><option value="row" ${zone.mode === "row" ? "selected" : ""}>横並び</option><option value="pile" ${zone.mode === "pile" ? "selected" : ""}>山・重ね置き</option>${zone.mode === "hand" ? `<option value="hand" selected>手札</option>` : ""}</select><select data-zone-hidden="${zone.id}" ${zone.id === "hand" ? "disabled" : ""}><option value="false" ${zone.hidden ? "" : "selected"}>表向き</option><option value="true" ${zone.hidden ? "selected" : ""}>裏向き</option></select><button class="small danger" data-action="delete-zone" data-zone="${zone.id}" ${protectedZone ? "disabled" : ""}>削除</button></div>`
    }).join("")
    openModal("ゾーンと表示の設定", `<div class="settings-top">${playerSettings}</div><h3>ゾーン名 / 並べ方 / 移動時の表裏</h3><p class="help-note">参考画像に合わせた基本配置です。ゾーン構成は自分と相手に共通です。相手を非表示にしても、相手のカードやデッキは保持します。「表向き・裏向き」は、そのゾーンに移動したときの既定値です。</p>${rows}<div class="control-row" style="margin-top:16px"><input id="newZoneName" maxlength="80" class="grow" placeholder="新しいゾーンの名前"><select id="newZoneMode"><option value="row">横並び</option><option value="pile">山・重ね置き</option></select><button data-action="add-zone">追加</button></div><p class="help-note">基本エリアは削除できません。山札と手札の並べ方は固定です。追加ゾーンを削除すると、中のカードをログに移します。配置の自由な座標指定やカード同士の付属関係は未対応です。</p><div class="modal-footer"><button class="danger" data-action="clear-board">盤面の通常カードを除く</button><button data-action="apply-settings" class="primary">設定を反映</button></div>`, "settings")
}
function applySettingsRaw() {
    for (const player of visiblePlayerIds()) {
        state.players[player].name = byId(`settingsName_${player}`).value.trim() || state.players[player].name
        for (let index = 0; index < 2; index += 1) state.players[player].meters[index].label = byId(`settingsMeter_${player}_${index}`).value.trim() || `カウンター ${index + 1}`
    }
    for (const zone of state.zones) {
        const label = document.querySelector(`[data-zone-label="${zone.id}"]`)
        const mode = document.querySelector(`[data-zone-mode="${zone.id}"]`)
        const hidden = document.querySelector(`[data-zone-hidden="${zone.id}"]`)
        zone.label = label.value.trim() || zone.label
        if (!["deck", "hand"].includes(zone.id)) zone.mode = mode.value
        zone.hidden = zone.id === "hand" ? false : hidden.value === "true"
    }
}
function addZone() {
    if (state.zones.length >= 40) return notify("ゾーンは40個までです。", true)
    const label = byId("newZoneName").value.trim()
    const mode = byId("newZoneMode").value
    if (!label) return notify("ゾーン名を入力してください。", true)
    transact("新しいゾーンを追加", () => {
        applySettingsRaw()
        const id = uid("zone")
        state.zones.push({id, label, mode, hidden: false})
        for (const player of playerIds) state.locations[player][id] = []
    })
    showSettings()
}
function deleteZone(id) {
    if (defaultZones().some(zone => zone.id === id)) return
    if (!confirm("追加ゾーンを削除し、中のカードをそれぞれのログに移しますか？")) return
    transact("ゾーンを削除", () => {
        applySettingsRaw()
        for (const player of playerIds) {
            moveRaw([...state.locations[player][id]], player, "discard")
            delete state.locations[player][id]
        }
        state.zones = state.zones.filter(zone => zone.id !== id)
    })
    showSettings()
}
function showHelp() {
    openModal("使い方と試作版の範囲", `<div class="help-content"><h3>最初に</h3><p>本アプリは、参考画像の配置を再現した非公式の手動シミュレータです。左にカラーライン7枠、中央にフィールド・セット・チャンネル・ステージ、右に山札とログ、下に7色のにじエリアを用意しています。相手の全フィールドは向かい合う配置で、上ににじエリア、下にフィールド、右にカラーライン、左に山札とログを表示します。初期手札の枚数は任意設定です。カードの効果、合法手、勝敗をプレイヤーが判断してください。</p><h3>カードとデッキを用意する</h3><p>「カード登録」の「Excelを取り込む」で.xlsxファイルを選ぶと、カード一覧を表示します。内容を確認し、使うカードを選んで登録してください。カード名・種別・クラス・色・レベル・パワー・タグ・効果を読み取ります。「効果を見る」で全文を確認できます。同じIDの登録済みカードには、未入力の効果・色・タグを追加できます。設定済みの効果・色と他の項目は保持します。色はライバーカードだけに反映します。カードごとの画像の追加や、文字カードの手入力もできます。文字カードにはIDを入力・編集できます。文字カードの種別はライバー・イベント・ステージ・チャンネル、クラスは太陽・月・彗星・星から選べます。左上にレベル、右上にクラスの頭文字、ライバーカードの右下にパワーを表示します。ライバーは7色から複数の色を設定でき、レベルの下から左端に沿って鮮やかな色帯を縦に表示します。画像カードも登録画面から種別と色を設定できます。チャンネルカードにレベルはありません。未設定の項目は表示しません。登録済みカードは12種類ずつ表示し、検索・種別の絞り込みと「編集」から変更できます。「デッキ管理」の専用ページで名前付きのデッキを複数保存できます。左にカードの詳細、中央にチャンネルとデッキ、右に追加用カード一覧を表示します。保存は盤面を変更せず、編集画面は開いたままです。盤面の「デッキを選ぶ」から保存デッキを山札とチャンネルへ反映してください。カード登録は両ページで共有します。「画像出力」では、カード画像・名前・ID・枚数を1枚のPNGにまとめます。「TXT出力」でデッキ名・チャンネル・カードのID・名前・枚数・並び順をテキストに保存できます。「TXT読込」で確認後、新しいデッキとして編集できます。使用するカードは先に登録してください。チャンネルの選択肢にはチャンネルカードだけを表示し、山札用のカード一覧には含めません。名前・ID検索、種別・クラスの絞り込み、並べ替えができ、ページ移動後も入力した枚数を保持します。中央のデッキ欄には編集中のカードと枚数を表示し、枚数を直接変更できます。チャンネルに指定したカードは山札に含めず、手札などから移すこともできません。山札を作り、操作側を選んで「枚をランダムに引き直す」を実行します。画像は長辺1200px以下に変換します。Excelは「全カード」シートを優先し、ID・カード名・種類・クラス・レベル・Powerと、任意のColors・効果列を読みます。デッキ一覧画像の自動切り出しと文字認識は入っていません。</p><h3>盤面を操作する</h3><p>カードをクリックして選択し、ゾーンへドラッグするか、左の「選択」タブから移動します。<kbd>Shift</kbd>＋クリックで複数選択。フィールド・手札などのカードや山札・ログの一覧内のカードは、ダブルクリックで内容を拡大確認します。盤面の山札・ログはダブルクリックすると一覧を開きます。山は先頭が一番上です。カードを同じゾーン内でドラッグすると、離した位置に並べ替えます。同じ位置に戻したときは履歴やリプレイの操作数には加えません。山札の一覧ではドラッグか位置番号の入力で順番を変更できます。一覧や上から確認でカードの内容を開いたときは、画面上部の戻るボタンで元の画面に戻れます。横向きはカード内の回転表示と「横」表示で確認できます。セットへの移動は初期設定で裏向きになります。セット・ステージは各1枚までです。2枚目を置くと、元のカードを表向き・縦向きでログへ送ります。カラーラインは上から順番に並びます。基本操作の「カラーカードの初期化」で、両側のカラーカードを裏向きで戻せます。「カラーカードのシャッフル」は、両者のカラーライン上にあるカラーカードの順番を、それぞれランダムに並べ替えます。表裏やにじエリアのカードはそのままで、「元に戻す」で取り消せます。以前の保存データにある旧ゾーンは削除し、中の通常カードをログへ移します。</p><h3>山札を調べる</h3><p>「上から確認」は、指定枚数を見せるだけで山札の順序を変えません。確認したカードを山札の下や別のゾーンへ移すと、そのカードを確認一覧から外します。表示枚数は減り、未確認のカードを追加で見せません。「山札の一覧」はPCでは50枚ずつ、スマートフォンでは24枚ずつ表示し、ページを移っても選択を保持します。一覧のカードをクリックして複数選択し、移動先と「上・先頭へ／下・末尾へ」を選んで「移動」を押します。移動先には「フィールド（横向き）」も選べます。ダブルクリックでめくって確認した画面でも同じ操作ができます。盤面の山札を選ぶと、枚数を指定してドロー・上から確認・上からログへ移動できます。複数枚は選んだ順に並びます。</p><h3>手動のまま残す部分</h3><p>コスト支払い、発動条件、効果の適用、処理順序、構築条件、通常カードの初期配置、ターン開始時処理、勝敗は自動判定しません。7色のカラーカードの初期配置と、デッキ編集で指定したチャンネル1枚を用意します。フェイズのボタンも表示だけを変更します。引き直しは確認表示を挟まず実行し、チャンネルとカラーカードを残して通常カードを山札に集めます。「元に戻す」で取り消せます。「フィールドをすべて縦向きにする」はフィールド内だけに適用します。「次へ」はターン表示を進め、フェイズをスタートに戻します。操作側は切り替えません。ライバーを選ぶとパワーを変更でき、同名の別カードには影響しません。「元のパワーに戻す」で登録値に戻せます。変更したパワー・カード上のカウンター・一時メモは通常の移動では残り、引き直しと初期盤面への復帰でリセットします。手札や山へ移すと縦向きになります。必要に応じてプレイヤーが状態を修正してください。</p><h3>一人回しと相手フィールド</h3><p>初期状態では自分の盤面と、相手のカラーライン・にじエリアを表示します。両側に赤・橙・黄・緑・青・藍・紫を各1枚、裏向きで用意します。カラーカードを選び、同じボタンで「表にする／裏にする」を切り替えられます。カラーライン上のカラーカードは、ダブルクリックでもその1枚の表裏を切り替えられます。「同じ色のにじエリアへ」で移動できます。「次へ」は自分のターン表示を進め、ドローなどの自動処理は行いません。左上の「相手の全フィールドを表示」を押すと、1台のPCで両方を操作できます。相手の全フィールドを表示しても「次へ」で操作側は変わりません。操作側は左側のプレイヤーボタンで切り替えます。相手をカラーだけの表示にすると、自分のターン操作に戻ります。相手のカラーライン・にじエリアは引き続き操作でき、他の相手カード・デッキ・カウンターも保持します。表示設定もJSON保存とブラウザ内保存に含めます。</p><p>相手を表示中は通常、操作側の手札だけを見せます。「両方の手札を表示する」で両方を見られます。一覧や拡大から相手の非公開カードも確認できるため、対戦相手から情報を保護する機能ではありません。通信対戦やCPU対戦は入っていません。</p><h3>初期盤面に戻す</h3><p>基本操作の「初期盤面に戻す」で、両側の山札とチャンネルを登録済みのデッキ設定から作り直し、7色をそれぞれのカラーラインに裏向きで戻します。手札・フィールド・セット・ステージ・ログ・にじエリアを空にし、カードのメモとカウンター、プレイヤーのカウンターもリセットします。ターン1のスタートフェイズに戻ります。カード登録、デッキ設定、相手の表示設定は残ります。手札はドローや引き直しで用意してください。「元に戻す」で直前の盤面に戻せます。</p><h3>リプレイを記録する</h3><p>盤面上部の「記録開始」を押して操作し、「記録終了」で完了します。最初の盤面と、その後のカード移動・表裏・縦横・カウンター・ターンなどを記録します。「再生」では一時停止、前後の操作への移動、再生位置を調整できます。「1ステップ」で操作ごとの時間を0.5〜5秒から選べます（初期値1.0秒）。実際の操作時間に関係なく、一定の間隔で進みます。再生中にフェイズが変わると、移動先のフェイズ名を自動でカットイン表示します。通常操作中や記録中には表示しません。山札の「上から確認」「すべて見る」や拡大確認も記録し、再生時にはその時に見たカードを、盤面に重なるカットインで表示します。カードの上に順番、下に名前を表示します。移動後も未確認のカードを追加で見せません。以前のリプレイも再生できますが、記録していなかった確認場面は表示できません。「操作に戻る」で再生前の盤面に戻ります。「リプレイ保存」でカード画像を含むJSONを保存し、「リプレイ読込」で再生できます。通常のJSON保存と自動保存にもリプレイを含めます。1つの記録は200操作・100MBまでです。山札の確認と確認終了も操作数に含みます。画面動画・音声の録画ではなく、盤面操作の再現です。記録し直す前に、残すリプレイを保存してください。</p><h3>テキストカットイン</h3><p>記録中の「テキストカットイン」で、好きな文章を1ステップとして挟めます。記録後も再生位置を選び、「テキストを挿入」でその直後に追加できます。最初の盤面を選ぶと最初の操作の前に挿入します。文章は300文字まで、改行も可能です。テキストの場面では文章の編集と削除ができます。表示時間は「1ステップ」の設定に従い、最後のステップでは停止して表示を残します。カットインの「閉じる」は表示だけを隠し、記録の内容は残します。編集結果も自動保存・JSON保存に含めます。</p><h3>保存と巻き戻し</h3><p>ブラウザ内に自動保存を試みますが、ファイルの移動、別のブラウザ、保存領域の削除や制限によって復元できないことがあります。「JSON保存」でカード画像・デッキ・盤面・メモをファイルに保存してください。「読込」はそのファイルを復元します。履歴欄の「履歴を初期化」で、盤面を保ったまま履歴だけを消せます。元に戻す・やり直すは現在の起動中の直近80操作までです。操作ログは保存しますが、JSON読込後の履歴再生や巻き戻しには対応していません。</p><section class="mobile-only" style="flex-direction:column"><h3>スマートフォンの操作</h3><p>画面下の「基本操作」「選択」「履歴」で操作欄を開き、「盤面」で戻ります。カードをタップしてから「選択」を開くと、移動や表裏の変更、内容の確認ができます。「複数選択」をオンにするとカードをまとめて選べます。山札・ログの「一覧」からも、1枚を選んで「内容を確認」を開けます。カードが多いゾーンは横にスワイプできます。</p></section><h3>キー操作</h3><p><kbd>D</kbd> ドロー、<kbd>R</kbd> 向きを変更、<kbd>F</kbd> 表裏反転、<kbd>Ctrl</kbd>＋<kbd>Z</kbd> 元に戻す、<kbd>Ctrl</kbd>＋<kbd>X</kbd> やり直す、<kbd>Ctrl</kbd>＋<kbd>S</kbd> JSON保存。入力中やダイアログを開いている間は、盤面のキー操作を止めます。</p><h3>オフライン動作</h3><p>HTMLと共通のCSS・JavaScriptを同じ配置で用意して使います。複数ページの保存共有には同じサイトのURLから開いてください。追加ライブラリ、アカウント、外部APIは不要です。公式のカード画像やイラストは同梱していません。このアプリのデータ保存・画像変換は端末内で行い、アプリからの外部通信を許可しない設定を入れています。</p></div>`, "help")
}
function checkState(candidate, images, deep = false) {
    const object = value => value !== null && typeof value === "object" && !Array.isArray(value)
    const safeId = value => typeof value === "string" && /^[A-Za-z0-9_]{1,90}$/.test(value) && !["__proto__", "constructor", "prototype"].includes(value)
    const text = (value, maximum) => typeof value === "string" && value.length <= maximum
    const number = (value, minimum, maximum) => Number.isInteger(value) && value >= minimum && value <= maximum
    const require = (condition, message) => {if (!condition) throw new Error(message)}
    require(object(candidate) && candidate.version === 1, "Unsupported state format.")
    require(candidate.layoutVersion === undefined || candidate.layoutVersion === 2, "Unsupported layout format.")
    require(candidate.colorSetupVersion === undefined || candidate.colorSetupVersion === 1, "Unsupported color setup format.")
    require(candidate.rulesVersion === undefined || candidate.rulesVersion === 1, "Unsupported rules format.")
    require(candidate.rulesVersion !== 1 || candidate.layoutVersion === 2, "Current rules require the playmat layout.")
    require(candidate.colorSetupVersion !== 1 || candidate.layoutVersion === 2, "Color setup requires the playmat layout.")
    require(candidate.phase === undefined || number(candidate.phase, 0, 4), "Invalid phase.")
    require(candidate.showOpponent === undefined || typeof candidate.showOpponent === "boolean", "Invalid opponent display setting.")
    require(playerIds.includes(candidate.active) && number(candidate.turn, 1, 999999), "Invalid player or turn.")
    require(Array.isArray(candidate.zones) && candidate.zones.length >= 3 && candidate.zones.length <= (candidate.colorSetupVersion === 1 ? 53 : 40), "Invalid zone list.")
    const zoneIds = new Set()
    for (const zone of candidate.zones) {
        require(object(zone) && safeId(zone.id) && !zoneIds.has(zone.id), "Invalid zone ID.")
        require(text(zone.label, 80) && zone.label.length > 0 && ["row", "pile", "hand"].includes(zone.mode) && typeof zone.hidden === "boolean", "Invalid zone settings.")
        require(zone.mode !== "hand" || zone.id === "hand", "Only the hand zone may use hand mode.")
        zoneIds.add(zone.id)
    }
    require(["deck", "hand", "discard"].every(id => zoneIds.has(id)), "Required zones are missing.")
    if (candidate.layoutVersion === 2) require(defaultZones().every(zone => zoneIds.has(zone.id)), "Playmat zones are missing.")
    require(candidate.zones.find(zone => zone.id === "deck").mode === "pile" && candidate.zones.find(zone => zone.id === "hand").mode === "hand", "Invalid deck or hand mode.")
    require(object(images) && Object.keys(images).length <= 1000, "Invalid image collection.")
    if (deep) {
        let total = 0
        for (const [id, value] of Object.entries(images)) {
            require(safeId(id) && typeof value === "string" && value.length <= 4000000, "Invalid image asset.")
            require(/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/]+=*$/.test(value), "Only embedded raster images are supported.")
            total += value.length
        }
        require(total <= 90000000, "Image collection exceeds the file limit.")
    }
    require(object(candidate.catalog) && Object.keys(candidate.catalog).length <= 507, "Invalid card catalog.")
    require(Object.values(candidate.catalog).filter(item => !isColorDefinition(item)).length <= 500, "Card catalog exceeds 500 regular definitions.")
    const definedColors = new Set()
    for (const [id, item] of Object.entries(candidate.catalog)) {
        require(safeId(id) && object(item) && text(item.name, 200) && item.name.length > 0 && text(item.text, 10000) && typeof item.asset === "string", "Invalid card definition.")
        require(item.asset === "" || (safeId(item.asset) && Object.hasOwn(images, item.asset)), "A card image is missing.")
        require(item.kind === undefined || item.kind === "color", "Invalid card kind.")
        require(item.cost === undefined || item.cost === null || number(item.cost, 0, 999), "Invalid card level.")
        require(item.className === undefined || text(item.className, 40), "Invalid card class.")
        require(item.cardType === undefined || (typeof item.cardType === "string" && Object.hasOwn(textCardTypes, item.cardType)), "Invalid card type.")
        require(item.isLiver === undefined || typeof item.isLiver === "boolean", "Invalid liver card setting.")
        require(item.power === undefined || item.power === null || number(item.power, 0, 999999), "Invalid card power.")
        require(item.colors === undefined || (Array.isArray(item.colors) && item.colors.length <= 7 && item.colors.every(color => nijiColors.includes(color)) && new Set(item.colors).size === item.colors.length), "Invalid liver card colors.")
        require(item.tags === undefined || validCardTags(item.tags), "Invalid liver card tags.")
        require(item.sourceCardKey === undefined || (typeof item.sourceCardKey === "string" && /^[A-Za-z0-9_-]{1,100}$/.test(item.sourceCardKey) && !["__proto__", "constructor", "prototype"].includes(item.sourceCardKey)), "Invalid source card key.")
        if (isColorDefinition(item)) {
            require(["cost", "className", "cardType", "isLiver", "power", "colors", "tags", "sourceCardKey"].every(key => item[key] === undefined), "Color cards cannot define text card details.")
            require(nijiColors.includes(item.color) && !definedColors.has(item.color), "Invalid or duplicate color definition.")
            definedColors.add(item.color)
        } else require(item.color === undefined, "Only color cards may define a color.")
    }
    if (candidate.channelDefinitions !== undefined) {
        require(object(candidate.channelDefinitions) && Object.keys(candidate.channelDefinitions).length === 2, "Invalid channel configuration.")
        for (const player of playerIds) {
            const id = candidate.channelDefinitions[player]
            require(id === "" || (safeId(id) && Object.hasOwn(candidate.catalog, id) && !isColorDefinition(candidate.catalog[id])), "Invalid channel definition.")
        }
    }
    require(object(candidate.cards) && Object.keys(candidate.cards).length <= 1016, "Invalid card instances.")
    const ownedColors = new Set()
    let regularCardCount = 0
    for (const [id, card] of Object.entries(candidate.cards)) {
        require(safeId(id) && object(card) && safeId(card.definition) && Object.hasOwn(candidate.catalog, card.definition), "Invalid card reference.")
        require(typeof card.faceDown === "boolean" && typeof card.rotated === "boolean" && number(card.counter, -99999, 99999) && text(card.note, 2000), "Invalid card state.")
        require(card.powerOverride === undefined || card.powerOverride === null || number(card.powerOverride, 0, 999999), "Invalid card power override.")
        if (isColorDefinition(candidate.catalog[card.definition])) {
            require(card.powerOverride === undefined, "Color cards cannot override power.")
            const key = `${card.colorOwner}_${candidate.catalog[card.definition].color}`
            require(playerIds.includes(card.colorOwner) && !ownedColors.has(key), "Invalid or duplicate color card owner.")
            ownedColors.add(key)
        } else {
            require(card.colorOwner === undefined, "Only color cards may define a color owner.")
            regularCardCount += 1
        }
    }
    require(regularCardCount <= 1002, "Card instances exceed 1000 regular cards and two channels.")
    if (candidate.colorSetupVersion === 1) require(ownedColors.size === 14, "Each player needs all seven color cards.")
    require(object(candidate.players) && object(candidate.locations) && object(candidate.decklists), "Missing player data.")
    const located = new Set()
    for (const player of playerIds) {
        const data = candidate.players[player]
        require(object(data) && text(data.name, 60) && data.name.length > 0 && Array.isArray(data.meters) && data.meters.length === 2, "Invalid player settings.")
        for (const meter of data.meters) require(object(meter) && text(meter.label, 50) && number(meter.value, -99999, 99999), "Invalid player counter.")
        require(object(candidate.locations[player]) && Object.keys(candidate.locations[player]).length === zoneIds.size, "Invalid zone contents.")
        for (const zone of candidate.zones) {
            const ids = candidate.locations[player][zone.id]
            require(Array.isArray(ids) && ids.length <= 1016, "Invalid zone cards.")
            require(candidate.rulesVersion !== 1 || !["set", "stage", "channel"].includes(zone.id) || ids.length <= 1, "Set, stage, and channel each hold at most one card.")
            for (const id of ids) {
                require(safeId(id) && Object.hasOwn(candidate.cards, id) && !located.has(id), "A card is duplicated or missing.")
                located.add(id)
            }
        }
        require(Array.isArray(candidate.decklists[player]) && candidate.decklists[player].length <= 500, "Invalid deck definition.")
        const definitions = new Set()
        let count = 0
        for (const item of candidate.decklists[player]) {
            require(object(item) && safeId(item.definition) && Object.hasOwn(candidate.catalog, item.definition) && !definitions.has(item.definition) && number(item.count, 1, 500), "Invalid deck entry.")
            require(!isColorDefinition(candidate.catalog[item.definition]), "Color cards cannot enter a deck list.")
            count += item.count
            definitions.add(item.definition)
        }
        require(count <= 500, "Deck exceeds 500 cards.")
    }
    require(located.size === Object.keys(candidate.cards).length, "A card has no zone.")
    require(Array.isArray(candidate.log) && candidate.log.length <= 200, "Invalid action log.")
    for (const entry of candidate.log) require(object(entry) && text(entry.time, 40) && Number.isFinite(Date.parse(entry.time)) && text(entry.label, 300), "Invalid log entry.")
    return true
}
function buildRecord(replayRecord = replay.data) {
    const live = replay.mode === "replay" ? replay.manual : {state, assets}
    const includedAssets = {...(replayRecord?.assets || {})}
    for (const item of Object.values(live.state.catalog)) if (item.asset) includedAssets[item.asset] = live.assets[item.asset]
    const record = {format: "nijinana-manual-table", formatVersion: 1, savedAt: new Date().toISOString(), state: copy(live.state), assets: includedAssets}
    if (replayRecord) {
        const {assets: recordedAssets, ...recorded} = replayRecord
        record.replay = copy(recorded)
    }
    record.savedDecks = copy(savedDecks)
    return record
}
function validateRecord(record) {
    if (!record || record.format !== "nijinana-manual-table" || record.formatVersion !== 1) throw new Error("This is not a supported playtable file.")
    if (new Blob([JSON.stringify(record)]).size > 100000000) throw new Error("盤面とリプレイを含む保存データは100MBまでです。")
    checkState(record.state, record.assets, true)
    if (record.replay !== undefined) validateReplayRecord({...record.replay, assets: record.assets})
    if (record.savedDecks !== undefined) validateSavedDecks(record.savedDecks, record.state.catalog)
}
function saveJSON() {
    try {
        const record = buildRecord()
        validateRecord(record)
        const blob = new Blob([JSON.stringify(record)], {type: "application/json"})
        const anchor = document.createElement("a")
        const url = URL.createObjectURL(blob)
        anchor.href = url
        anchor.download = `nijinana_${new Date().toISOString().slice(0, 19).replaceAll(":", "-")}.json`
        anchor.click()
        setTimeout(() => URL.revokeObjectURL(url), 15000)
        notify("JSONファイルを書き出しました。ブラウザの保存先を確認してください。")
    } catch (error) {
        console.error("Export failed", error)
        notify(`保存できませんでした。${error.message}`, true)
    }
}
async function readStateFile(file) {
    if (!file) return
    if (!storageReady && !(storageFailed && appPage === "board")) return notify("保存領域の準備が終わってから読み込んでください。", true)
    if (replay.recording || replay.mode !== "manual") return notify("記録・再生を終了してから盤面を読み込んでください。")
    const replayEpoch = replay.epoch
    let previousRecord = null
    try {
        if (file.size > 100000000) throw new Error("File exceeds 100 MB.")
        const record = JSON.parse(await file.text())
        validateRecord(record)
        if (replayEpoch !== replay.epoch || replay.recording || replay.mode !== "manual") throw new Error("操作状態が変わったため、読み込みを中止しました。")
        if (!confirm(appPage === "deck" ? "カード登録・保存デッキをファイルから復元します。現在の盤面と盤面で使用中の登録カードは保持します。続けますか？" : "カード登録・保存デッキ・盤面をファイルの内容に置き換えます。続けますか？")) return
        if (appPage === "deck" && !(await confirmDeckManagerLeave())) return
        if (storageReady) await flushStorageSave()
        previousRecord = buildRecord()
        const restoredState = upgradeLayout(copy(record.state))
        checkState(restoredState, record.assets, true)
        state = restoredState
        assets = copy(record.assets)
        restoreReplayFromBoard(record)
        const importedDecks = record.savedDecks ?? playerIds.filter(player => state.decklists[player].length || state.channelDefinitions[player]).map(player => {
            const now = new Date().toISOString()
            return {id: uid("deck"), name: `${state.players[player].name}のデッキ`, list: copy(state.decklists[player]).filter(entry => isDeckMainDefinition(entry.definition)), channelDefinition: isDeckChannelDefinition(state.channelDefinitions[player]) ? state.channelDefinitions[player] : "", createdAt: now, updatedAt: now}
        })
        if (storageReady) await replaceNamedDecks(importedDecks)
        else {validateSavedDecks(importedDecks, state.catalog); savedDecks = copy(importedDecks)}
        previousRecord = null
        undoStack = []
        redoStack = []
        ui.selection.clear()
        ui.tab = "operation"
        ui.movePlayer = state.active
        ui.touched = true
        resetDeckManagerAfterImport()
        closeModal(false)
        render()
        notify("カード登録・保存デッキを復元しました。")
    } catch (error) {
        if (previousRecord) {
            state = copy(previousRecord.state)
            assets = copy(previousRecord.assets)
            restoreReplayFromBoard(previousRecord)
            render()
        }
        console.error("Import failed", error)
        notify(`読み込めませんでした。${error.message}`, true)
    }
}
function imageData(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader()
        reader.onerror = () => reject(new Error("Image file could not be read."))
        reader.onload = () => {
            const image = new Image()
            image.onerror = () => reject(new Error("Image could not be decoded."))
            image.onload = () => {
                if (!image.naturalWidth || !image.naturalHeight || image.naturalWidth * image.naturalHeight > 80000000) return reject(new Error("Invalid image dimensions."))
                const ratio = Math.min(1, 1200 / Math.max(image.naturalWidth, image.naturalHeight))
                const canvas = document.createElement("canvas")
                canvas.width = Math.max(1, Math.round(image.naturalWidth * ratio))
                canvas.height = Math.max(1, Math.round(image.naturalHeight * ratio))
                const context = canvas.getContext("2d")
                context.fillStyle = "#ffffff"
                context.fillRect(0, 0, canvas.width, canvas.height)
                context.drawImage(image, 0, 0, canvas.width, canvas.height)
                resolve(canvas.toDataURL("image/jpeg", 0.92))
            }
            image.src = reader.result
        }
        reader.readAsDataURL(file)
    })
}
async function importImages(files) {
    if (replay.mode !== "manual") return
    const replayEpoch = replay.epoch
    if (!files.length) return
    const valid = [...files]
    if (valid.length > 80 || valid.length + Object.values(state.catalog).filter(item => !isColorDefinition(item)).length > 500) return notify("画像は一度に80枚、登録合計500種類までです。", true)
    if (valid.some(file => !["image/png", "image/jpeg", "image/webp"].includes(file.type) || file.size > 10000000)) return notify("PNG・JPEG・WebPの各10MB以下の画像を選んでください。", true)
    notify("端末内で画像を変換しています。")
    try {
        const prepared = []
        for (const file of valid) prepared.push({name: file.name.replace(/\.[^.]+$/, "").slice(0, 200) || "カード", data: await imageData(file)})
        if (replayEpoch !== replay.epoch || replay.mode !== "manual") throw new Error("操作状態が変わったため、画像の登録を中止しました。")
        const usedAssets = new Set(Object.values(state.catalog).map(item => item.asset).filter(Boolean))
        const total = [...usedAssets].reduce((sum, id) => sum + assets[id].length, 0) + prepared.reduce((sum, item) => sum + item.data.length, 0)
        if (total > 80000000) throw new Error("Image collection exceeds the safe export limit.")
        const success = transact(`${prepared.length}種類のカード画像を登録`, () => {
            for (const item of prepared) {
                const assetId = uid("image")
                assets[assetId] = item.data
                state.catalog[uid("definition")] = {name: item.name, text: "", asset: assetId}
            }
        })
        if (!success) return
        if (modalKind === "catalog") showCatalog()
        notify(`${prepared.length}種類の画像を登録しました。デッキ編集で枚数を指定してください。`)
    } catch (error) {
        console.error("Image import failed", error)
        notify(`画像を読み込めませんでした。${error.message}`, true)
    }
}
function operateSelection(kind) {
    const ids = selectedIds().filter(id => kind !== "rotate" || where(id)?.zone !== "hand")
    if (!ids.length) return notify("カードを選択してください。")
    const label = kind === "rotate" ? "選択したカードの向きを変更" : "選択したカードの表裏を反転"
    transact(label, () => {
        for (const id of ids) {
            if (kind === "rotate") state.cards[id].rotated = !state.cards[id].rotated
            else state.cards[id].faceDown = !state.cards[id].faceDown
        }
    })
}
function operateColors(action) {
    const ids = selectedIds().filter(id => colorForCard(id))
    if (!ids.length) return notify("カラーカードを選択してください。")
    const label = action === "color-reveal" ? "カラーカードを表にする" : action === "color-to-niji" ? "カラーカードを同じ色のにじエリアへ移動" : ids.every(id => where(id).zone === "colorline") ? "カラーカードを裏にする" : "カラーカードを裏向きでカラーラインへ戻す"
    transact(label, () => {
        for (const id of ids) {
            const player = state.cards[id].colorOwner || where(id).player
            const staysOnLine = action === "color-to-line" && where(id).zone === "colorline"
            if (action !== "color-reveal" && !staysOnLine) moveRaw([id], player, action === "color-to-niji" ? `niji_${colorForCard(id).toLowerCase()}` : "colorline")
            state.cards[id].faceDown = action === "color-to-line"
        }
    })
}
async function goToDeckManager() {
    try {
        if (storageReady) await flushStorageSave()
        await waitLocalDeckTransfer()
        window.location.href = appDeckManagerURL()
    } catch (error) {notify(`保存できなかったため、ページの移動を中止しました。${error.message}`, true)}
}
function handleAction(button) {
    if (appPage === "deck" && ((deckManagerPendingLeave && !button.dataset.action.startsWith("manager-leave-") && button.dataset.action !== "close-modal") || (deckManager.saving && !button.dataset.action.startsWith("manager-leave-")))) return
    const action = button.dataset.action
    if (!storageReady && !storageFailed && !["help", "close-modal"].includes(action)) {notify("保存データを読み込み中です。少しお待ちください。"); return}
    if (handleDeckManagerAction(action, button)) return
    if (action === "mobile-close") {setMobilePanel(false); return}
    if (action === "replay-record") {
        if (replay.data) openModal("新しい記録を開始", `<p>現在のリプレイを新しい記録に置き換えます。残す場合は先に保存してください。</p><div class="modal-footer"><button data-action="replay-save">リプレイ保存</button><button data-action="close-modal">キャンセル</button><button data-action="replay-record-confirm" class="primary">新しく記録する</button></div>`, "replay-confirm")
        else startReplayRecording()
        return
    }
    if (action === "replay-record-confirm") {startReplayRecording(); return}
    if (action === "replay-stop-recording") {stopReplayRecording(); return}
    if (action === "replay-enter") {if (enterReplay()) toggleReplayPlayback(); return}
    if (action === "replay-play-pause") {toggleReplayPlayback(); return}
    if (action === "replay-prev") {seekReplay(replay.index - 1); return}
    if (action === "replay-next") {seekReplay(replay.index + 1); return}
    if (action === "replay-exit") {leaveReplay(); return}
    if (action === "replay-save") {saveReplay(); return}
    if (action === "replay-load") {byId("replayPicker").click(); return}
    if (action === "replay-add-cutin") {openReplayCutInEditor(); return}
    if (action === "replay-edit-cutin") {openReplayCutInEditor(true); return}
    if (action === "replay-save-cutin") {saveReplayCutInFromEditor(); return}
    if (action === "replay-delete-cutin") {deleteReplayTextCutIn(); return}
    if (action === "replay-dismiss-cutin") {dismissReplayCutIn(); return}
    if (handleReplayInspectionAction(action, button)) return
    if (replay.mode === "replay" && !["save", "help", "close-modal"].includes(action)) return
    if (action === "mobile-panel") {setMobilePanel(true, button.dataset.tab); return}
    if (action === "mobile-multiple") {
        if (!mobileViewport?.matches) return
        mobileUi.multiple = !mobileUi.multiple
        renderSideContent()
        if (mobileUi.multiple) setMobilePanel(false)
        return
    }
    if (handleCatalogAction(action, button) || handleDeckEditorAction(action, button) || handleInspectorAction(action, button)) return
    if (action === "catalog") showCatalog()
    else if (action === "card-library") showCardLibrary()
    else if (action === "pick-card-library") byId("cardLibraryPicker").click()
    else if (action === "library-register") registerLibraryCards()
    else if (action === "library-effect") toggleLibraryEffect(button.dataset.libraryKey)
    else if (action === "library-clear") {libraryUi.selected.clear(); renderLibraryResults()}
    else if (action === "library-page") {libraryUi.page += integer(button.dataset.delta, -1, 1, 0); renderLibraryResults()}
    else if (action === "library-select-page") {
        const registered = registeredLibraryDefinitions()
        for (const card of filteredLibraryCards().slice(libraryUi.page * libraryPageSize, (libraryUi.page + 1) * libraryPageSize)) {
            if (isLibraryCardEligible(card, registered)) libraryUi.selected.add(card.key)
        }
        renderLibraryResults()
    }
    else if (action === "deck-editor") showSavedDeckSelector(button.dataset.player || state.active)
    else if (action === "settings") showSettings()
    else if (action === "help") showHelp()
    else if (action === "close-modal") closeModal()
    else if (action === "preview-return") returnToInspector()
    else if (action === "save") saveJSON()
    else if (action === "load") byId("statePicker").click()
    else if (action === "pick-images") byId("imagePicker").click()
    else if (action === "add-text-card") addTextCard()
    else if (action === "demo") demo()
    else if (["color-reveal", "color-to-niji", "color-to-line"].includes(action)) operateColors(action)
    else if (action === "reset-colors") {
        transact("両側のカラーカードを初期化", () => {
            for (const player of playerIds) resetColorCards(player)
        })
    }
    else if (action === "draw") draw()
    else if (action === "shuffle") shuffleDeck()
    else if (action === "shuffle-colors") shuffleColorCards()
    else if (action === "ready") readyField()
    else if (action === "reset-board") resetBoard()
    else if (action === "redeal") redeal()
    else if (action === "undo") undo()
    else if (action === "redo") redo()
    else if (action === "clear-history") clearHistory()
    else if (action === "peek") inspectZone(state.active, "deck", ui.peekCount)
    else if (action === "inspect") inspectZone(button.dataset.player, button.dataset.zone)
    else if (action === "tab") {ui.tab = button.dataset.tab
        render()
    }
    else if (action === "active") {
        if (!visiblePlayerIds().includes(button.dataset.player)) return
        state.active = button.dataset.player
        ui.movePlayer = state.active
        ui.selection.clear()
        ui.touched = true
        captureReplayFrame("操作するプレイヤーを変更")
        render()
        scheduleSave()
    }
    else if (action === "toggle-opponent") transact(state.showOpponent ? "相手はカラーだけ表示" : "相手の全フィールドを表示", () => {
        state.showOpponent = !state.showOpponent
        ui.selection.clear()
        normalizePlayerView()
        ui.movePlayer = state.active
    })
    else if (action === "phase") transact(`${phases[integer(button.dataset.phase, 0, 4)]}フェイズを表示`, () => {
        state.phase = integer(button.dataset.phase, 0, 4)
    })
    else if (action === "next-turn") transact("ターン表示を進める", () => {
        state.turn += 1
        state.phase = 0
        ui.selection.clear()
    })
    else if (action === "clear-selection") {ui.selection.clear()
        render()
    }
    else if (action === "rotate" || action === "flip") operateSelection(action)
    else if (action === "set-selected-power") setSelectedPower(byId("selectedPower")?.value)
    else if (action === "reset-selected-power") setSelectedPower(null)
    else if (action === "preview-selected") {
        const id = selectedIds()[0]
        if (id) preview(id)
    }
    else if (action === "move-selected") moveCards(selectedIds(), ui.movePlayer, ui.moveZone, ui.position)
    else if (action === "deck-move") {
        const destination = byId("deckMoveDestination").value
        const position = byId("deckMovePosition").value === "first" ? "first" : "last"
        moveDeckSelection(destination === "field-rotated" ? "field" : destination, destination === "field-rotated", button.dataset.sourceCard, position)
    }
    else if (action === "inspect-reorder" && modalKind === "zone" && inspectContext?.zone === "deck" && inspectContext.limit === null) {
        const player = inspectContext.player
        const deck = state.locations[player].deck
        const moving = selectedIds().filter(id => where(id)?.player === player && where(id)?.zone === "deck")
        if (!moving.length) return notify("山札のカードを選択してください。")
        const value = byId("inspectDeckPosition")?.value?.trim()
        const position = Number(value)
        if (!value || !Number.isInteger(position) || position < 1 || position > deck.length) return notify("山札内の位置を入力してください。")
        const finalIndex = Math.min(position - 1, deck.length - moving.length)
        if (reorderZone(moving, player, "deck", finalIndex, true)) {
            inspectContext.page = Math.floor(finalIndex / currentInspectorPageSize())
            renderZoneInspector()
        }
    }
    else if (["selected-deck-draw", "selected-deck-peek", "selected-deck-discard"].includes(action)) {
        const operation = action.slice("selected-deck-".length)
        const player = button.dataset.player
        operateSelectedDeck(operation, player, byId(`selectedDeck_${player}_${operation}`)?.value)
    }
    else if (action === "card-counter") transact("選択したカードのカウンターを変更", () => {
        for (const id of selectedIds()) state.cards[id].counter = integer(state.cards[id].counter + Number(button.dataset.delta), -99999, 99999)
    })
    else if (action === "reset-counter") transact("選択したカードのカウンターを0に変更", () => {
        for (const id of selectedIds()) state.cards[id].counter = 0
    })
    else if (action === "save-note") {
        const id = selectedIds()[0]
        const note = byId("cardNote").value
        if (id) transact("カードの一時メモを変更", () => {state.cards[id].note = note})
    }
    else if (action === "meter") transact("プレイヤーのカウンターを変更", () => {
        const meter = state.players[button.dataset.player].meters[Number(button.dataset.index)]
        meter.value = integer(meter.value + Number(button.dataset.delta), -99999, 99999)
    })
    else if (action === "inspect-select-all") {
        for (const id of visibleInspectorIds()) ui.selection.add(id)
        render()
    }
    else if (action === "inspect-show-all") {inspectContext.limit = null
        inspectContext.snapshot = null
        inspectContext.page = 0
        renderZoneInspector()
        captureReplayFrame(`${state.players[inspectContext.player].name}の山札をすべて確認`)
    }
    else if (["inspect-move", "inspect-top", "inspect-bottom"].includes(action)) {
        const target = action === "inspect-move" ? byId("inspectDestination").value : "deck"
        const position = action === "inspect-top" ? "first" : "last"
        const moving = selectedIds()
        moveCards(moving, inspectContext.player, target, position)
        if (inspectContext.snapshot) inspectContext.snapshot = inspectContext.snapshot.filter(id => !moving.includes(id))
        ui.selection.clear()
        render()
    }
    else if (action === "deck-copy-other") {
        if (!state.showOpponent) return
        collectDeckDraft()
        const other = editorPlayer === "p1" ? "p2" : "p1"
        deckDrafts[other] = copy(deckDrafts[editorPlayer])
        channelDrafts[other] = channelDrafts[editorPlayer]
        notify("もう片側の入力にコピーしました。「両方のデッキを保存」で山札を作成できます。")
    }
    else if (action === "apply-deck") applyDeck(false)
    else if (action === "apply-both-decks") applyDeck(true)
    else if (action === "add-zone") addZone()
    else if (action === "delete-zone") deleteZone(button.dataset.zone)
    else if (action === "apply-settings") {
        transact("ゾーン名と表示設定を変更", applySettingsRaw)
        closeModal()
    }
    else if (action === "clear-board") {
        const label = state.showOpponent ? "両方の盤面" : "自分の盤面"
        if (!confirm(`${label}から通常カードを取り除きます。チャンネル・カラーカード・カード登録・デッキの入力内容は残します。実行しますか？`)) return
        transact(`${label}をクリア`, () => {
            for (const player of visiblePlayerIds()) {
                for (const zone of state.zones) {
                    if (zone.id === "channel") continue
                    for (const id of state.locations[player][zone.id]) if (!colorForCard(id)) delete state.cards[id]
                    state.locations[player][zone.id] = state.locations[player][zone.id].filter(id => colorForCard(id))
                }
                for (const meter of state.players[player].meters) meter.value = 0
            }
        })
        closeModal()
    }
    else if (action === "delete-definition") {
        const id = button.dataset.definition
        const used = Object.values(state.cards).some(card => card.definition === id) || playerIds.some(player => state.channelDefinitions[player] === id || state.decklists[player].some(item => item.definition === id)) || savedDeckUsesDefinition(id) || deckManagerUsesDefinition(id)
        if (!used && !isColorDefinition(state.catalog[id])) {
            transact("未使用のカード登録を削除", () => {delete state.catalog[id]})
            showCatalog()
        }
    }
}
document.addEventListener("click", event => {
    const navigation = event.target.closest("a[href]")
    if (appPage === "board" && navigation && (/^deck\/?(?:index\.html)?$/.test(navigation.getAttribute("href")) || (navigation.href === appDeckManagerURL()))) {
        event.preventDefault()
        goToDeckManager()
        return
    }
    const button = event.target.closest("[data-action]")
    if (button) {
        if (button.disabled) return
        try {handleAction(button)} catch (error) {
            console.error("Action handler failed", error)
            notify(`操作できませんでした。${error.message}`, true)
        }
        return
    }
    if (replay.mode === "replay") return
    const card = event.target.closest("[data-card-id]")
    if (!card) return
    const id = card.dataset.cardId
    if (!Object.hasOwn(state.cards, id)) return
    const multiple = event.shiftKey || event.ctrlKey || event.metaKey || modalKind === "zone" || (mobileViewport?.matches && mobileUi.multiple)
    if (!multiple) ui.selection.clear()
    if (multiple && ui.selection.has(id)) ui.selection.delete(id)
    else ui.selection.add(id)
    const place = where(id)
    ui.movePlayer = place.player
    ui.tab = "selection"
    refreshSelectionUI()
})
document.addEventListener("dblclick", event => {
    if (replay.mode === "replay") return
    const pile = event.target.closest(".zone-deck, .zone-discard")
    if (pile?.dataset.dropPlayer) {
        inspectZone(pile.dataset.dropPlayer, pile.dataset.dropZone)
        return
    }
    const card = event.target.closest("[data-card-id]")
    if (!card) return
    const id = card.dataset.cardId
    if (!Object.hasOwn(state.cards, id)) return
    const place = where(id)
    if (place?.zone === "deck" && modalKind !== "zone") return
    if (place?.zone === "colorline" && colorForCard(id)) {
        transact("カラーカードの表裏を反転", () => {state.cards[id].faceDown = !state.cards[id].faceDown})
    } else preview(id)
})
document.addEventListener("input", event => {
    if (replay.mode === "replay") return
    const element = event.target
    if (handleDeckManagerInput(element) || handleCatalogInput(element) || handleDeckEditorInput(element)) return
    if (element.id === "drawCount") ui.drawCount = integer(element.value, 1, 500)
    else if (element.id === "peekCount") ui.peekCount = integer(element.value, 1, 500)
    else if (["drawCount", "peekCount", "discardCount"].includes(element.dataset.deckCount)) ui[element.dataset.deckCount] = integer(element.value, 1, 500)
    else if (element.id === "initialCount") ui.initialCount = integer(element.value, 0, 500, 0)
    else if (element.id === "librarySearch") {libraryUi.query = element.value; libraryUi.page = 0; renderLibraryResults()}
})
document.addEventListener("change", event => {
    const element = event.target
    if (element.id === "replaySeek") {seekReplay(element.value); return}
    if (element.id === "replayStepSeconds") {setReplayStepSeconds(element.value); return}
    if (replay.mode === "replay") return
    if (handleDeckManagerChange(element) || handleCatalogChange(element) || handleDeckEditorChange(element)) return
    if (element.id === "libraryType" || element.id === "libraryClass") {
        libraryUi[element.id === "libraryType" ? "cardType" : "className"] = element.value
        libraryUi.page = 0
        renderLibraryResults()
        return
    }
    if (element.dataset.libraryKey) {selectLibraryCard(element.dataset.libraryKey, element.checked); return}
    if (element.id === "bothHands") {ui.bothHands = element.checked
        captureReplayFrame("手札の表示を変更")
        if (replay.recording) scheduleSave()
        renderBoard()
    }
    else if (element.id === "movePlayer") {
        ui.movePlayer = playerIds.includes(element.value) ? element.value : "p1"
        normalizePlayerView()
        renderSideContent()
    }
    else if (element.id === "moveZone" || element.id === "inspectDestination") ui.moveZone = element.value
    else if (element.id === "deckMoveDestination") ui.deckMoveZone = element.value
    else if (element.id === "movePosition" || element.id === "deckMovePosition") ui.position = element.value
    else if (element.id === "newCardType") {
        syncCardDetailFields({cardType: element.value})
    }
    else if (element.dataset.catalogColor) {
        const id = element.dataset.definition
        updateCatalogCardDetail(id, "colors", readCardDetailColors(id))
        syncCardDetailFields(state.catalog[id], id)
    }
    else if (element.dataset.catalogDetail) {
        const field = element.dataset.catalogDetail
        const id = element.dataset.definition
        updateCatalogCardDetail(id, field, element.value)
        const item = state.catalog[id]
        if (item) {
            element.value = field === "cardType" ? textCardType(item) : field === "tags" ? (item.tags || []).join(", ") : (item[field] ?? "")
            if (field === "cardType") syncCardDetailFields(item, id)
        }
    }
    else if (element.dataset.catalogName) {
        const name = element.value.trim()
        const id = element.dataset.catalogName
        if (!name) {element.value = state.catalog[id].name
            return notify("カード名は空にできません。", true)
        }
        transact("登録カードの名前を変更", () => {state.catalog[id].name = name})
    }
    else if (element.dataset.catalogId) {
        const id = element.dataset.catalogId
        const previous = state.catalog[id]?.sourceCardKey || ""
        const scrollTop = document.querySelector(".catalog-list")?.scrollTop || 0
        if (!updateCatalogCardId(id, element.value)) element.value = previous
        else if ((state.catalog[id]?.sourceCardKey || "") !== previous) {
            renderCatalogResults()
            const list = document.querySelector(".catalog-list")
            if (list) list.scrollTop = scrollTop
        }
    }
    else if (element.dataset.catalogText) {
        const text = element.value
        transact("登録カードの説明を変更", () => {state.catalog[element.dataset.catalogText].text = text})
    }
})
byId("imagePicker").addEventListener("change", async event => {
    const files = [...event.target.files]
    event.target.value = ""
    await importImages(files)
})
byId("statePicker").addEventListener("change", async event => {
    const file = event.target.files[0]
    event.target.value = ""
    await readStateFile(file)
})
byId("cardLibraryPicker").addEventListener("change", async event => {
    const file = event.target.files[0]
    event.target.value = ""
    await readCardLibraryFile(file)
})
byId("replayPicker").addEventListener("change", async event => {
    const file = event.target.files[0]
    event.target.value = ""
    await readReplayFile(file)
})
byId("modal").addEventListener("close", () => {
    if (byId("modal").open) return
    if (appPage === "deck" && deckManagerPendingLeave) {finishDeckManagerLeave(false); return}
    clearDeckExportPreview()
    modalKind = ""
    inspectContext = null
    previewCardId = null
    previewReturnContext = null
    replay.cutInEditor = null
    if (appPage === "deck") renderDeckPage(); else captureReplayFrame("山札の確認を終了")
})
document.addEventListener("dragstart", event => {
    if (replay.mode === "replay") {event.preventDefault(); return}
    const element = event.target.closest("[data-card-id]")
    if (!element) return
    const id = element.dataset.cardId
    if (!ui.selection.has(id)) ui.selection = new Set([id])
    event.dataTransfer.setData("text/plain", JSON.stringify(selectedIds()))
    event.dataTransfer.effectAllowed = "move"
    ui.dragging = true
})
let activeDropZone = null
let dropIndicator = null
function clearDropFeedback() {
    activeDropZone?.classList.remove("drag-over")
    activeDropZone = null
    dropIndicator?.remove()
    dropIndicator = null
}
function boardDropPosition(zone, event) {
    const player = zone.dataset.dropPlayer
    const zoneId = zone.dataset.dropZone
    const current = state.locations[player]?.[zoneId]
    const cardsBox = zone.querySelector(".zone-cards")
    if (!current || !cardsBox) return null
    if (zoneById(zoneId)?.mode === "pile") return {boundary: current.length, reference: null}
    const currentIds = new Set(current)
    const elements = [...cardsBox.querySelectorAll("[data-card-id]")]
        .filter(element => currentIds.has(element.dataset.cardId))
    if (!elements.length) return {boundary: 0, reference: null}
    const direction = getComputedStyle(cardsBox).flexDirection
    const vertical = direction.startsWith("column")
    const reverse = direction.endsWith("reverse")
    const coordinate = vertical ? event.clientY : event.clientX
    let boundary = elements.length
    for (let index = 0; index < elements.length; index += 1) {
        const rect = elements[index].getBoundingClientRect()
        const middle = vertical ? (rect.top + rect.bottom) / 2 : (rect.left + rect.right) / 2
        if (reverse ? coordinate > middle : coordinate < middle) {boundary = index; break}
    }
    return {boundary, reference: elements[Math.min(boundary, elements.length - 1)], before: boundary < elements.length, vertical, reverse}
}
function inspectorDropPosition(grid, event) {
    const elements = [...grid.querySelectorAll(".inspect-item")]
    const start = inspectContext.page * currentInspectorPageSize()
    if (!elements.length) return {boundary: start, reference: null}
    const rows = []
    elements.forEach((element, index) => {
        const rect = element.getBoundingClientRect()
        let row = rows[rows.length - 1]
        if (!row || Math.abs(rect.top - row.top) > 6) {
            row = {top: rect.top, bottom: rect.bottom, start: index, end: index, elements: []}
            rows.push(row)
        }
        row.bottom = Math.max(row.bottom, rect.bottom)
        row.end = index + 1
        row.elements.push({element, rect})
    })
    let row = rows[rows.length - 1]
    for (let index = 0; index < rows.length - 1; index += 1) {
        if (event.clientY < (rows[index].bottom + rows[index + 1].top) / 2) {row = rows[index]; break}
    }
    let local = row.end
    for (let index = 0; index < row.elements.length; index += 1) {
        const rect = row.elements[index].rect
        if (event.clientX < (rect.left + rect.right) / 2) {local = row.start + index; break}
    }
    const reference = local < row.end ? elements[local] : elements[row.end - 1]
    return {boundary: start + local, reference, before: local < row.end, vertical: false, reverse: false}
}
function dropPosition(zone, event) {
    return zone.dataset.inspectorDrop === "true" ? inspectorDropPosition(zone, event) : boardDropPosition(zone, event)
}
function showDropIndicator(position, zone) {
    if (!position?.reference) {dropIndicator?.remove(); dropIndicator = null; return}
    if (!dropIndicator) {
        dropIndicator = document.createElement("div")
        dropIndicator.className = "drop-insertion-marker"
    }
    const parent = zone.dataset.inspectorDrop === "true" ? byId("modal") : document.body
    if (dropIndicator.parentElement !== parent) parent.appendChild(dropIndicator)
    const rect = position.reference.getBoundingClientRect()
    const before = position.before !== position.reverse
    if (position.vertical) {
        dropIndicator.style.left = `${rect.left}px`
        dropIndicator.style.top = `${(before ? rect.top : rect.bottom) - 2}px`
        dropIndicator.style.width = `${rect.width}px`
        dropIndicator.style.height = "4px"
    } else {
        dropIndicator.style.left = `${(before ? rect.left : rect.right) - 2}px`
        dropIndicator.style.top = `${rect.top}px`
        dropIndicator.style.width = "4px"
        dropIndicator.style.height = `${rect.height}px`
    }
}
function acceptCardDrop(event) {
    const zone = event.target.closest("[data-drop-zone]")
    if (!zone || !ui.dragging) return
    event.preventDefault()
    event.dataTransfer.dropEffect = "move"
    if (activeDropZone && activeDropZone !== zone) activeDropZone.classList.remove("drag-over")
    activeDropZone = zone
    zone.classList.add("drag-over")
    showDropIndicator(dropPosition(zone, event), zone)
}
document.addEventListener("dragenter", acceptCardDrop)
document.addEventListener("dragover", acceptCardDrop)
document.addEventListener("dragleave", event => {
    const zone = event.target.closest("[data-drop-zone]")
    if (zone && zone === activeDropZone && !zone.contains(event.relatedTarget)) clearDropFeedback()
})
document.addEventListener("dragend", () => {
    ui.dragging = false
    clearDropFeedback()
})
document.addEventListener("drop", event => {
    if (replay.mode === "replay") return
    const zone = event.target.closest("[data-drop-zone]")
    if (!zone || !ui.dragging) return
    event.preventDefault()
    const position = dropPosition(zone, event)
    ui.dragging = false
    clearDropFeedback()
    try {
        const ids = JSON.parse(event.dataTransfer.getData("text/plain"))
        if (!Array.isArray(ids) || !ids.every(id => typeof id === "string" && Object.hasOwn(state.cards, id))) return
        ui.selection = new Set(ids)
        ui.movePlayer = zone.dataset.dropPlayer
        ui.tab = "selection"
        if (!dropCardsAt([...new Set(ids)], zone.dataset.dropPlayer, zone.dataset.dropZone, position.boundary, zone.dataset.inspectorDrop === "true")) refreshSelectionUI()
    } catch (error) {
        console.error("Drag operation failed", error)
        notify("カードを移動できませんでした。", true)
    }
})
document.addEventListener("scroll", () => {if (replay.mode === "replay") positionReplayCutIn()}, true)
globalThis.addEventListener?.("resize", () => {if (replay.mode === "replay") positionReplayCutIn()})
document.addEventListener("keydown", event => {
    if (appPage === "deck") {
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s" && !byId("modal").open) {event.preventDefault(); saveCurrentNamedDeck()}
        return
    }
    if (!storageReady) return
    if (event.key === "Escape" && mobileUi.panelOpen && !byId("modal").open) {event.preventDefault(); setMobilePanel(false); return}
    const editing = event.target.matches("input, textarea, select, [contenteditable=true]")
    if (editing || byId("modal").open) return
    const key = event.key.toLowerCase()
    const modifier = event.ctrlKey || event.metaKey
    if (replay.mode === "replay") {
        if (key === " " && !event.target.matches("button")) {event.preventDefault(); toggleReplayPlayback()}
        else if (key === "arrowleft") {event.preventDefault(); seekReplay(replay.index - 1)}
        else if (key === "arrowright") {event.preventDefault(); seekReplay(replay.index + 1)}
        return
    }
    if (modifier && key === "s") {event.preventDefault()
        saveJSON()
    }
    else if (modifier && key === "z") {event.preventDefault()
        if (event.shiftKey) redo()
        else undo()
    }
    else if (modifier && key === "x") {event.preventDefault()
        redo()
    }
    else if (!modifier && !event.altKey && key === "d") draw()
    else if (!modifier && !event.altKey && key === "r") operateSelection("rotate")
    else if (!modifier && !event.altKey && key === "f") operateSelection("flip")
    else if (key === "escape") {ui.selection.clear()
        render()
    }
})
document.addEventListener("DOMContentLoaded", async () => {
    if (appDeckRedirect) {location.replace(appDeckManagerURL()); return}
    if (localDeckBridge) {await serveLocalDeckTransfer(); return}
    render()
    const ready = await initializeStorage()
    if (localDeckRoot && ready) await initializeLocalDeckTransfer()
    else {localDeckTransferPending = false; render()}
})
