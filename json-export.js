/* Export default cards as references; keep complete records inside browser storage. */
const JSON_EXPORT_MAX_BYTES = 100000000
const JSON_EXPORT_MAX_CATALOG = 507

function jsonRecordObject(value) {return value !== null && typeof value === "object" && !Array.isArray(value)}
function jsonRecordSafeId(value) {
    return typeof value === "string" && /^[A-Za-z0-9_]{1,90}$/.test(value) && !["__proto__", "constructor", "prototype"].includes(value)
}
function jsonRecordSourceId(value) {
    return typeof value === "string" && /^[A-Za-z0-9_-]{1,100}$/.test(value) && !["__proto__", "constructor", "prototype"].includes(value)
}
function jsonRecordRequire(condition, message) {if (!condition) throw new Error(message)}
function jsonRecordBytes(value) {return new Blob([JSON.stringify(value)]).size}
function jsonRecordCheckEnvelope(record) {
    jsonRecordRequire(jsonRecordObject(record) && ["nijinana-manual-table", "nijinana-replay"].includes(record.format) && [1, 2].includes(record.formatVersion), "対応していないJSON形式です。")
    jsonRecordRequire(jsonRecordBytes(record) <= JSON_EXPORT_MAX_BYTES, "JSONファイルは100MBまでです。")
    if (record.format === "nijinana-manual-table") {
        jsonRecordRequire(jsonRecordObject(record.state), "盤面の形式が正しくありません。")
        if (record.replay !== undefined) {
            jsonRecordRequire(jsonRecordObject(record.replay) && record.replay.format === "nijinana-replay" && record.replay.formatVersion === record.formatVersion, "盤面内のリプレイ形式が正しくありません。")
        }
    }
    const replayRecord = record.format === "nijinana-replay" ? record : record.replay
    if (replayRecord) {
        jsonRecordRequire(Array.isArray(replayRecord.frames) && replayRecord.frames.length >= 1 && replayRecord.frames.length <= REPLAY_MAX_OPERATIONS + 1, "リプレイの操作数が不正です。")
        for (const frame of replayRecord.frames) jsonRecordRequire(jsonRecordObject(frame) && jsonRecordObject(frame.state), "リプレイの盤面形式が正しくありません。")
    }
}
function jsonRecordStates(record) {
    const states = record.format === "nijinana-manual-table" ? [record.state] : []
    const replayRecord = record.format === "nijinana-replay" ? record : record.replay
    if (replayRecord) for (const frame of replayRecord.frames) states.push(frame.state)
    return states
}
function jsonRecordValidate(record) {
    if (record.format === "nijinana-manual-table") validateRecord(record)
    else validateReplayRecord(record)
}
function jsonRecordBundleMap() {
    const definitions = bundledCardDefinitions()
    jsonRecordRequire(Array.isArray(definitions) && definitions.length > 0, "標準カードを読み込んでからJSONを書き出してください。")
    const result = new Map()
    for (const item of definitions) {
        jsonRecordRequire(jsonRecordObject(item) && jsonRecordSourceId(item.sourceCardKey) && !result.has(item.sourceCardKey), "標準カードのIDが正しくありません。")
        result.set(item.sourceCardKey, item)
    }
    return result
}
function jsonRecordSavedDeckReferences(record) {
    const used = new Set()
    for (const deck of record.savedDecks || []) {
        for (const entry of deck.list) used.add(entry.definition)
        if (deck.channelDefinition) used.add(deck.channelDefinition)
    }
    return used
}
function jsonRecordCompactState(candidate, bundled, additionalReferences = new Set()) {
    const used = storageDefinitionReferences(candidate)
    for (const id of additionalReferences) used.add(id)
    const catalog = [], references = []
    for (const [id, item] of Object.entries(candidate.catalog)) {
        const color = isColorDefinition(item)
        const standard = !color && bundled.has(item.sourceCardKey)
        if (!color && !standard) catalog.push([id, item])
        else if (used.has(id)) references.push([id, color ? {color: item.color} : {sourceCardKey: item.sourceCardKey}])
    }
    candidate.catalog = Object.fromEntries(catalog)
    candidate.defaultCardReferences = Object.fromEntries(references)
}
function compactJsonRecord(fullRecord) {
    jsonRecordCheckEnvelope(fullRecord)
    jsonRecordRequire(fullRecord.formatVersion === 1, "書き出すJSONは完全な盤面データである必要があります。")
    for (const candidate of jsonRecordStates(fullRecord)) jsonRecordRequire(candidate.defaultCardReferences === undefined, "盤面に標準カードの参照情報が重複しています。")
    jsonRecordValidate(fullRecord)
    const bundled = jsonRecordBundleMap(), record = copy(fullRecord)
    const states = jsonRecordStates(record)
    for (const candidate of states) jsonRecordCompactState(candidate, bundled, candidate === record.state ? jsonRecordSavedDeckReferences(record) : undefined)
    const usedAssets = new Set()
    for (const candidate of states) for (const item of Object.values(candidate.catalog)) if (item.asset) usedAssets.add(item.asset)
    record.assets = Object.fromEntries(Object.entries(record.assets).filter(([id]) => usedAssets.has(id)))
    record.formatVersion = 2
    if (record.replay) {
        record.replay.formatVersion = 2
        delete record.replay.assets
    }
    jsonRecordCheckEnvelope(record)
    return record
}
function jsonRecordExpandState(candidate, getBundle, getColors) {
    jsonRecordRequire(jsonRecordObject(candidate.catalog) && jsonRecordObject(candidate.defaultCardReferences), "標準カードの参照情報が正しくありません。")
    const entries = Object.entries(candidate.defaultCardReferences)
    jsonRecordRequire(entries.length <= JSON_EXPORT_MAX_CATALOG && Object.keys(candidate.catalog).length + entries.length <= JSON_EXPORT_MAX_CATALOG, "カード登録の上限を超えています。")
    for (const [id, reference] of entries) {
        jsonRecordRequire(jsonRecordSafeId(id) && !Object.hasOwn(candidate.catalog, id) && jsonRecordObject(reference) && Object.keys(reference).length === 1, "標準カードの参照情報が重複しているか、不正です。")
        let definition
        if (Object.hasOwn(reference, "sourceCardKey")) {
            jsonRecordRequire(jsonRecordSourceId(reference.sourceCardKey), "標準カードのIDが正しくありません。")
            definition = getBundle().get(reference.sourceCardKey)
            jsonRecordRequire(definition !== undefined, "標準カード「" + reference.sourceCardKey + "」がdefault-cards.csvに見つかりません。CSVを確認してください。")
        } else {
            jsonRecordRequire(Object.hasOwn(reference, "color") && nijiColors.includes(reference.color), "カラーカードの色が正しくありません。")
            definition = getColors().get(reference.color)
            jsonRecordRequire(definition !== undefined, "カラーカードを復元できませんでした。")
        }
        candidate.catalog[id] = copy(definition)
    }
    delete candidate.defaultCardReferences
}
function expandJsonRecord(inputRecord) {
    jsonRecordCheckEnvelope(inputRecord)
    const record = copy(inputRecord), states = jsonRecordStates(record)
    if (record.formatVersion === 1) {
        for (const candidate of states) jsonRecordRequire(candidate.defaultCardReferences === undefined, "標準カードの参照情報とJSONのバージョンが一致しません。")
        jsonRecordValidate(record)
        return record
    }
    let bundled = null, colors = null, expandedBytes = jsonRecordBytes(record)
    const getBundle = () => {
        if (bundled === null) bundled = jsonRecordBundleMap()
        return bundled
    }
    const getColors = () => {
        if (colors === null) colors = new Map(Object.values(blankState().catalog).filter(isColorDefinition).map(item => [item.color, item]))
        return colors
    }
    for (const candidate of states) {
        const previousBytes = jsonRecordBytes(candidate)
        jsonRecordExpandState(candidate, getBundle, getColors)
        expandedBytes += jsonRecordBytes(candidate) - previousBytes
        jsonRecordRequire(expandedBytes <= JSON_EXPORT_MAX_BYTES, "標準カードの復元後のJSONが100MBを超えます。")
    }
    record.formatVersion = 1
    if (record.replay) {
        record.replay.formatVersion = 1
        delete record.replay.assets
    }
    jsonRecordValidate(record)
    return record
}
