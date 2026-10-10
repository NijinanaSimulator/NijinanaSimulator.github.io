/* Shared cards and named decks use separate stores so the deck page never saves the playmat. */
let savedDecks = []
let storageReady = false
let storageFailed = false
let storageCatalogBaseline = {}
let storageSaveQueue = Promise.resolve()
let storageInitialization = null
let storageRefreshTimer = null
let storageChannel = null
const storageMigrationKey = "shared_decks_migrated_v2"

function storageStatus(message) {
    const element = byId("saveStatus")
    if (element) element.textContent = message
}
function storageLiveData() {return replay.mode === "replay" && replay.manual ? replay.manual : {state, assets}}
function regularStorageCatalog(catalog) {return Object.fromEntries(Object.entries(catalog).filter(([, item]) => !isColorDefinition(item)))}
function storageEqual(left, right) {return JSON.stringify(left) === JSON.stringify(right)}
function storageCatalogDelta(catalog, baseline = storageCatalogBaseline) {
    const regular = regularStorageCatalog(catalog), changed = {}, removed = []
    for (const [id, item] of Object.entries(regular)) if (!storageEqual(item, baseline[id])) changed[id] = copy(item)
    for (const id of Object.keys(baseline)) if (!Object.hasOwn(regular, id)) removed.push(id)
    return {changed, removed}
}
function storageDefinitionReferences(candidate) {
    const used = new Set(Object.values(candidate.cards || {}).map(card => card.definition))
    for (const player of playerIds) {
        for (const entry of candidate.decklists?.[player] || []) used.add(entry.definition)
        if (candidate.channelDefinitions?.[player]) used.add(candidate.channelDefinitions[player])
    }
    return used
}
function savedDeckUsesDefinition(id) {
    return savedDecks.some(deck => deck.channelDefinition === id || deck.list.some(entry => entry.definition === id))
}
function uniqueSavedDeckName(name, decks, excludedId = "") {
    const requested = String(name || "").trim()
    const used = new Set(decks.filter(deck => deck.id !== excludedId).map(deck => deck.name.trim()))
    if (!requested || requested.length > 120 || !used.has(requested)) return requested
    for (let index = 2; ; index += 1) {
        const suffix = ` (${index})`
        const candidate = `${requested.slice(0, 120 - suffix.length).trimEnd()}${suffix}`
        if (!used.has(candidate)) return candidate
    }
}
function validateSavedDecks(decks, catalog) {
    const object = value => value !== null && typeof value === "object" && !Array.isArray(value)
    const safeId = value => typeof value === "string" && /^[A-Za-z0-9_]{1,90}$/.test(value) && !["__proto__", "constructor", "prototype"].includes(value)
    const require = (condition, message) => {if (!condition) throw new Error(message)}
    require(Array.isArray(decks) && decks.length <= 100, "保存できるデッキは100個までです。")
    require(object(catalog), "カード登録の形式が正しくありません。")
    const ids = new Set()
    for (const deck of decks) {
        require(object(deck) && safeId(deck.id) && !ids.has(deck.id), "デッキIDの形式が正しくありません。")
        ids.add(deck.id)
        require(typeof deck.name === "string" && deck.name.trim().length > 0 && deck.name.length <= 120, "デッキ名は120文字以内で入力してください。")
        for (const field of ["createdAt", "updatedAt"]) require(typeof deck[field] === "string" && deck[field].length <= 40 && Number.isFinite(Date.parse(deck[field])), "デッキの保存日時が正しくありません。")
        require(deck.channelDefinition === "" || (safeId(deck.channelDefinition) && Object.hasOwn(catalog, deck.channelDefinition) && !isColorDefinition(catalog[deck.channelDefinition]) && textCardType(catalog[deck.channelDefinition]) === "channel"), "チャンネルカードを確認してください。")
        require(Array.isArray(deck.list) && deck.list.length <= 500, "デッキのカード一覧が正しくありません。")
        const definitions = new Set()
        let total = 0
        for (const entry of deck.list) {
            require(object(entry) && safeId(entry.definition) && Object.hasOwn(catalog, entry.definition) && !definitions.has(entry.definition), "デッキに未登録または重複したカードがあります。")
            require(!isColorDefinition(catalog[entry.definition]) && textCardType(catalog[entry.definition]) !== "channel", "カラー・チャンネルカードは山札に入れられません。")
            require(Number.isInteger(entry.count) && entry.count >= 1 && entry.count <= 500, "カードの枚数は1〜500枚で指定してください。")
            definitions.add(entry.definition)
            total += entry.count
        }
        require(total <= 500, "デッキは500枚までです。")
    }
    return true
}
function requireDeckStorage() {
    if (!storageReady || !database) throw new Error(storageFailed ? "ブラウザ内の保存を利用できません。保存を有効にして画面を読み込み直してください。" : "保存先を準備しています。少し待ってからもう一度操作してください。")
}
function storageReadRequest(request) {
    return new Promise((resolve, reject) => {
        request.onsuccess = () => resolve(request.result)
        request.onerror = () => reject(request.error || new Error("ブラウザ内の保存データを読み込めませんでした。"))
    })
}
function storageTransactionDone(transaction) {
    return new Promise((resolve, reject) => {
        transaction.oncomplete = () => resolve()
        transaction.onerror = () => reject(transaction.error || new Error("ブラウザ内への保存に失敗しました。"))
        transaction.onabort = transaction.onerror
    })
}
async function readSharedStorage() {
    const transaction = database.transaction(["catalog", "assets", "decks"], "readonly")
    const done = storageTransactionDone(transaction), catalogStore = transaction.objectStore("catalog"), assetStore = transaction.objectStore("assets")
    const [requests] = await Promise.all([Promise.all([
        storageReadRequest(catalogStore.getAllKeys()), storageReadRequest(catalogStore.getAll()),
        storageReadRequest(assetStore.getAllKeys()), storageReadRequest(assetStore.getAll()),
        storageReadRequest(transaction.objectStore("decks").getAll())
    ]), done])
    return {catalog: Object.fromEntries(requests[0].map((key, index) => [key, requests[1][index]])),
        assets: Object.fromEntries(requests[2].map((key, index) => [key, requests[3][index]])), decks: requests[4]}
}
function applySharedStorage(snapshot, preserveChanges = true) {
    const live = storageLiveData()
    const delta = preserveChanges ? storageCatalogDelta(live.state.catalog) : {changed: {}, removed: []}
    const used = storageDefinitionReferences(live.state)
    const nextCatalog = Object.fromEntries(Object.entries(live.state.catalog).filter(([, item]) => isColorDefinition(item)))
    Object.assign(nextCatalog, copy(snapshot.catalog), delta.changed)
    for (const id of delta.removed) if (!snapshot.decks.some(deck => deck.channelDefinition === id || deck.list.some(entry => entry.definition === id)) && !used.has(id)) delete nextCatalog[id]
    // A board can still refer to an older definition; keep that copy until the card leaves the board.
    for (const id of used) if (!Object.hasOwn(nextCatalog, id) && live.state.catalog[id]) nextCatalog[id] = copy(live.state.catalog[id])
    const nextAssets = {...snapshot.assets, ...live.assets}
    for (const [id, item] of Object.entries(snapshot.catalog)) if (item.asset && !Object.hasOwn(delta.changed, id) && snapshot.assets[item.asset]) nextAssets[item.asset] = snapshot.assets[item.asset]
    checkState({...live.state, catalog: nextCatalog}, nextAssets)
    validateSavedDecks(snapshot.decks, nextCatalog)
    live.state.catalog = nextCatalog
    if (replay.mode === "replay" && replay.manual) replay.manual.assets = nextAssets
    else assets = nextAssets
    storageCatalogBaseline = copy(snapshot.catalog)
    savedDecks = copy(snapshot.decks).sort((left, right) => Date.parse(right.updatedAt) - Date.parse(left.updatedAt) || left.name.localeCompare(right.name, "ja"))
}
function renderSharedStorage() {
    render()
    if (modalKind === "deck-picker" && typeof renderSavedDeckSelector === "function") renderSavedDeckSelector()
}
function broadcastDeckStorage() {try {storageChannel?.postMessage({type: "changed"})} catch {}}
function queueStorageOperation(operation) {
    const pending = storageSaveQueue.then(operation, operation)
    storageSaveQueue = pending.catch(() => {})
    return pending
}
async function persistStorageSnapshot(replacementDecks = null) {
    requireDeckStorage()
    const live = storageLiveData(), catalog = copy(regularStorageCatalog(live.state.catalog)), imageData = {...live.assets}
    const delta = storageCatalogDelta(live.state.catalog), boardRecord = appPage === "board" ? buildRecord() : null
    const transaction = database.transaction(["boards", "catalog", "assets", "decks"], "readwrite")
    const done = storageTransactionDone(transaction), catalogStore = transaction.objectStore("catalog"), decksStore = transaction.objectStore("decks"), assetStore = transaction.objectStore("assets")
    let result = null, operationError = null
    const latest = {keys: null, values: null, decks: null, board: null, assetKeys: null, assetValues: null}
    const process = () => {
        if (Object.values(latest).some(value => value === null)) return
        try {
            const remoteCatalog = Object.fromEntries(latest.keys.map((key, index) => [key, latest.values[index]]))
            const remoteAssets = Object.fromEntries(latest.assetKeys.map((key, index) => [key, latest.assetValues[index]]))
            const nextDecks = replacementDecks === null ? latest.decks : copy(replacementDecks)
            const mergedCatalog = replacementDecks === null ? {...remoteCatalog, ...delta.changed} : {...catalog}
            if (replacementDecks !== null && appPage === "deck" && latest.board?.state) {
                for (const id of storageDefinitionReferences(latest.board.state)) {
                    const boardItem = latest.board.state.catalog?.[id]
                    if (!Object.hasOwn(mergedCatalog, id) && boardItem && !isColorDefinition(boardItem)) mergedCatalog[id] = copy(boardItem)
                }
            }
            if (replacementDecks === null) {
                const boardUsed = appPage === "board" ? storageDefinitionReferences(live.state) : storageDefinitionReferences(latest.board?.state || {})
                for (const id of delta.removed) {
                    if (nextDecks.some(deck => deck.channelDefinition === id || deck.list.some(entry => entry.definition === id)) || boardUsed.has(id)) throw new Error("保存済みのデッキや盤面で使っているカードは削除できません。")
                    delete mergedCatalog[id]
                }
            }
            validateSavedDecks(nextDecks, mergedCatalog)
            if (Object.keys(mergedCatalog).length > 500) throw new Error("カード登録は500種類までです。別の画面で追加したカードも含めて確認してください。")
            if (replacementDecks !== null) {
                catalogStore.clear()
                decksStore.clear()
                for (const [id, item] of Object.entries(mergedCatalog)) catalogStore.put(item, id)
                for (const deck of nextDecks) decksStore.put(deck, deck.id)
            } else {
                for (const [id, item] of Object.entries(delta.changed)) catalogStore.put(item, id)
                for (const id of delta.removed) catalogStore.delete(id)
            }
            const sharedAssetIds = new Set(Object.values(mergedCatalog).map(item => item.asset).filter(Boolean))
            const nextAssets = Object.fromEntries(Object.entries(remoteAssets).filter(([id]) => sharedAssetIds.has(id)))
            for (const id of Object.keys(remoteAssets)) if (!sharedAssetIds.has(id)) assetStore.delete(id)
            for (const [id, item] of Object.entries(mergedCatalog)) if (item.asset) {
                const localDefinition = replacementDecks !== null ? Object.hasOwn(catalog, id) : Object.hasOwn(delta.changed, id)
                const data = localDefinition ? imageData[item.asset] || remoteAssets[item.asset] || latest.board?.assets?.[item.asset]
                    : remoteAssets[item.asset] || latest.board?.assets?.[item.asset] || imageData[item.asset]
                if (!data) throw new Error("カード画像が見つかりません。カード登録を確認してください。")
                if (data !== remoteAssets[item.asset]) assetStore.put(data, item.asset)
                nextAssets[item.asset] = data
            }
            if (boardRecord) {
                const boardCatalog = {...boardRecord.state.catalog, ...mergedCatalog}
                for (const id of delta.removed) delete boardCatalog[id]
                boardRecord.state.catalog = boardCatalog
                for (const item of Object.values(boardCatalog)) if (item.asset && nextAssets[item.asset]) boardRecord.assets[item.asset] = nextAssets[item.asset]
                boardRecord.savedDecks = copy(nextDecks)
                validateRecord(boardRecord)
                transaction.objectStore("boards").put(boardRecord, "current")
            }
            result = {catalog: mergedCatalog, assets: nextAssets, decks: nextDecks}
        } catch (error) {operationError = error; transaction.abort()}
    }
    const read = (key, request) => {
        request.onsuccess = () => {latest[key] = request.result === undefined ? false : request.result; process()}
        request.onerror = () => {operationError = request.error; transaction.abort()}
    }
    read("keys", catalogStore.getAllKeys())
    read("values", catalogStore.getAll())
    read("decks", decksStore.getAll())
    read("board", transaction.objectStore("boards").get("current"))
    read("assetKeys", assetStore.getAllKeys())
    read("assetValues", assetStore.getAll())
    try {await done} catch (error) {throw operationError || error}
    applySharedStorage(result)
    broadcastDeckStorage()
    return result
}
function scheduleSave(allowReplay = false) {
    if (replay.mode === "replay" && !allowReplay) return
    clearTimeout(saveTimer)
    saveVersion += 1
    if (!storageReady || !database) {storageStatus(storageFailed ? "自動保存は利用できません。JSON保存を使用してください。" : "ブラウザ内の保存先を準備中 / バックアップはJSON保存"); return}
    const version = saveVersion
    storageStatus("このブラウザ内に保存中")
    saveTimer = setTimeout(() => {
        saveTimer = null
        queueStorageOperation(() => persistStorageSnapshot()).then(() => {
            if (version === saveVersion) storageStatus("このブラウザ内に保存済み / バックアップはJSON保存")
        }).catch(error => {
            console.error("Local save failed", error)
            storageStatus("自動保存に失敗しました。JSON保存を使用してください。")
            notify(error.message || "ブラウザ内への保存に失敗しました。", true)
        })
    }, 450)
}
async function flushStorageSave() {
    requireDeckStorage()
    clearTimeout(saveTimer)
    saveTimer = null
    const version = ++saveVersion
    storageStatus("このブラウザ内に保存中")
    try {
        const result = await queueStorageOperation(() => persistStorageSnapshot())
        if (version === saveVersion) storageStatus("このブラウザ内に保存済み / バックアップはJSON保存")
        return result
    } catch (error) {storageStatus("保存に失敗しました。JSON保存を使用してください。"); throw error}
}
async function refreshDeckStorage() {
    requireDeckStorage()
    return queueStorageOperation(async () => {
        applySharedStorage(await readSharedStorage())
        renderSharedStorage()
        return savedDecks
    })
}
async function saveNamedDeck(deck) {
    requireDeckStorage()
    await flushStorageSave()
    return queueStorageOperation(async () => {
        const next = copy(deck), now = new Date().toISOString()
        next.name = String(next.name || "").trim()
        if (!next.createdAt) next.createdAt = now
        next.updatedAt = now
        const transaction = database.transaction(["decks", "catalog"], "readwrite"), done = storageTransactionDone(transaction)
        let operationError = null, decks = null, keys = null, values = null
        const process = () => {
            if (!decks || !keys || !values) return
            try {
                const existing = decks.find(item => item.id === next.id)
                if (existing) next.createdAt = existing.createdAt
                next.name = uniqueSavedDeckName(next.name, decks, next.id)
                const all = decks.filter(item => item.id !== next.id).concat(next)
                validateSavedDecks(all, Object.fromEntries(keys.map((key, index) => [key, values[index]])))
                transaction.objectStore("decks").put(next, next.id)
            } catch (error) {operationError = error; transaction.abort()}
        }
        const deckRead = transaction.objectStore("decks").getAll(), keyRead = transaction.objectStore("catalog").getAllKeys(), valueRead = transaction.objectStore("catalog").getAll()
        deckRead.onsuccess = () => {decks = deckRead.result; process()}
        keyRead.onsuccess = () => {keys = keyRead.result; process()}
        valueRead.onsuccess = () => {values = valueRead.result; process()}
        try {await done} catch (error) {throw operationError || error}
        applySharedStorage(await readSharedStorage())
        broadcastDeckStorage()
        return copy(next)
    })
}
async function deleteNamedDeck(id) {
    requireDeckStorage()
    return queueStorageOperation(async () => {
        const transaction = database.transaction("decks", "readwrite"), done = storageTransactionDone(transaction)
        transaction.objectStore("decks").delete(id)
        await done
        applySharedStorage(await readSharedStorage())
        broadcastDeckStorage()
        return true
    })
}
async function replaceNamedDecks(decks) {
    requireDeckStorage()
    validateSavedDecks(decks, storageLiveData().state.catalog)
    clearTimeout(saveTimer)
    saveTimer = null
    return queueStorageOperation(() => persistStorageSnapshot(copy(decks)))
}
async function importLocalDeckStorage(snapshot, migrationKey) {
    requireDeckStorage()
    const object = value => value !== null && typeof value === "object" && !Array.isArray(value)
    const safeId = value => typeof value === "string" && /^[A-Za-z0-9_]{1,90}$/.test(value) && !["__proto__", "constructor", "prototype"].includes(value)
    if (!safeId(migrationKey) || ["current", storageMigrationKey].includes(migrationKey)) throw new Error("移行の保存先が正しくありません。")
    if (!object(snapshot) || !object(snapshot.catalog) || !object(snapshot.assets) || !Array.isArray(snapshot.decks)) throw new Error("旧デッキの保存データを確認できませんでした。")
    if (new Blob([JSON.stringify(snapshot)]).size > 100000000) throw new Error("移行するカードとデッキは100MBまでです。")
    const incoming = {catalog: copy(snapshot.catalog), assets: {...snapshot.assets}, decks: copy(snapshot.decks)}
    if (Object.values(incoming.catalog).some(isColorDefinition)) throw new Error("移行するカード登録にカラーカードを含めることはできません。")
    const validationState = blankState()
    Object.assign(validationState.catalog, incoming.catalog)
    checkState(validationState, incoming.assets, true)
    validateSavedDecks(incoming.decks, incoming.catalog)
    if (saveTimer !== null) await flushStorageSave()
    else await storageSaveQueue
    return queueStorageOperation(async () => {
        const transaction = database.transaction(["boards", "catalog", "assets", "decks"], "readwrite"), done = storageTransactionDone(transaction)
        const latest = {keys: null, values: null, assetKeys: null, assetValues: null, decks: null, board: null, marker: null}
        let operationError = null, result = null, counts = null
        const canonical = value => JSON.stringify(value, (_key, item) => object(item) ? Object.fromEntries(Object.keys(item).sort().map(key => [key, item[key]])) : item)
        const deckSignature = deck => canonical({name: deck.name.trim(), list: deck.list, channelDefinition: deck.channelDefinition})
        const process = () => {
            if (Object.values(latest).some(value => value === null)) return
            try {
                const catalog = Object.fromEntries(latest.keys.map((key, index) => [key, latest.values[index]]))
                const images = Object.fromEntries(latest.assetKeys.map((key, index) => [key, latest.assetValues[index]]))
                const decks = copy(latest.decks)
                if (latest.marker) {
                    result = {catalog, assets: images, decks}
                    counts = {cards: 0, assets: 0, decks: 0, skippedDecks: incoming.decks.length, alreadyImported: true}
                    return
                }
                // Retain definitions referenced by the existing playmat, without writing its record.
                if (latest.board?.state) for (const id of storageDefinitionReferences(latest.board.state)) {
                    const item = latest.board.state.catalog?.[id]
                    if (item && !isColorDefinition(item) && !Object.hasOwn(catalog, id)) catalog[id] = copy(item)
                    const definition = catalog[id]
                    if (definition?.asset && !Object.hasOwn(images, definition.asset) && latest.board.assets?.[definition.asset]) images[definition.asset] = latest.board.assets[definition.asset]
                }
                const definitionIds = new Set([...Object.keys(catalog), ...Object.keys(incoming.catalog), ...Object.keys(validationState.catalog).filter(id => isColorDefinition(validationState.catalog[id]))])
                const imageIds = new Set([...Object.keys(images), ...Object.keys(incoming.assets)])
                const deckIds = new Set([...decks.map(deck => deck.id), ...incoming.decks.map(deck => deck.id)])
                const freshId = (prefix, ids) => {let id; do {id = uid(prefix)} while (ids.has(id)); ids.add(id); return id}
                const assetMap = new Map(), definitionMap = new Map(), imageByData = new Map(Object.entries(images).map(([id, data]) => [data, id]))
                const definitionByValue = new Map(Object.entries(catalog).map(([id, item]) => [canonical(item), id]))
                const used = new Set([...storageDefinitionReferences(latest.board?.state || {}), ...storageDefinitionReferences(state), ...storageDefinitionReferences(storageLiveData().state)])
                for (const deck of decks) {if (deck.channelDefinition) used.add(deck.channelDefinition); for (const entry of deck.list) used.add(entry.definition)}
                let bundled = []
                try {if (typeof bundledCardDefinitions === "function") bundled = bundledCardDefinitions()} catch {}
                const bundledByKey = new Map(bundled.map(item => [item.sourceCardKey, canonical(item)]))
                const untouchedBuiltins = new Map(Object.entries(catalog).filter(([id, item]) => id.startsWith("builtin_") && !used.has(id) && !(typeof deckManagerUsesDefinition === "function" && deckManagerUsesDefinition(id)) && bundledByKey.get(item.sourceCardKey) === canonical(item)).map(([id, item]) => [item.sourceCardKey, id]))
                const updatedDefinitions = new Set()
                counts = {cards: 0, assets: 0, decks: 0, skippedDecks: 0, alreadyImported: false}
                for (const [sourceId, original] of Object.entries(incoming.catalog)) {
                    const item = copy(original)
                    if (item.asset) {
                        const sourceAsset = item.asset, data = incoming.assets[sourceAsset]
                        if (!assetMap.has(sourceAsset)) {
                            let id = sourceAsset
                            if (!Object.hasOwn(images, id) || images[id] !== data) id = imageByData.get(data) || (Object.hasOwn(images, id) ? freshId("asset", imageIds) : id)
                            if (!Object.hasOwn(images, id)) {images[id] = data; imageByData.set(data, id); counts.assets += 1}
                            assetMap.set(sourceAsset, id)
                        }
                        item.asset = assetMap.get(sourceAsset)
                    }
                    const signature = canonical(item)
                    let id = sourceId
                    const builtin = untouchedBuiltins.get(item.sourceCardKey)
                    if (builtin) {
                        id = builtin
                        const previous = canonical(catalog[id])
                        if (previous !== signature) {
                            if (definitionByValue.get(previous) === id) definitionByValue.delete(previous)
                            catalog[id] = item
                            updatedDefinitions.add(id)
                        }
                        definitionByValue.set(signature, id)
                        untouchedBuiltins.delete(item.sourceCardKey)
                    } else if (!Object.hasOwn(catalog, id) || canonical(catalog[id]) !== signature) id = definitionByValue.get(signature) || (Object.hasOwn(catalog, id) ? freshId("definition", definitionIds) : id)
                    if (!Object.hasOwn(catalog, id)) {catalog[id] = item; definitionByValue.set(signature, id); counts.cards += 1}
                    definitionMap.set(sourceId, id)
                }
                const existingDecks = new Set(decks.map(deckSignature))
                for (const original of incoming.decks) {
                    const deck = copy(original), list = [], entries = new Map()
                    for (const entry of deck.list) {
                        const definition = definitionMap.get(entry.definition)
                        if (entries.has(definition)) entries.get(definition).count += entry.count
                        else {const mapped = {definition, count: entry.count}; entries.set(definition, mapped); list.push(mapped)}
                    }
                    deck.list = list
                    deck.channelDefinition = deck.channelDefinition ? definitionMap.get(deck.channelDefinition) : ""
                    const signature = deckSignature(deck)
                    if (existingDecks.has(signature)) {counts.skippedDecks += 1; continue}
                    if (decks.some(existing => existing.id === deck.id)) deck.id = freshId("deck", deckIds)
                    deck.name = uniqueSavedDeckName(deck.name, decks)
                    decks.push(deck)
                    existingDecks.add(signature)
                    counts.decks += 1
                }
                const mergedState = blankState()
                Object.assign(mergedState.catalog, catalog)
                checkState(mergedState, images, true)
                const live = storageLiveData()
                const liveColors = Object.fromEntries(Object.entries(live.state.catalog).filter(([, item]) => isColorDefinition(item)))
                checkState({...live.state, catalog: {...liveColors, ...catalog}}, {...live.assets, ...images}, true)
                validateSavedDecks(decks, catalog)
                const catalogStore = transaction.objectStore("catalog"), assetStore = transaction.objectStore("assets"), deckStore = transaction.objectStore("decks")
                const existingDefinitions = new Set(latest.keys), existingImages = new Set(latest.assetKeys), existingIds = new Set(latest.decks.map(deck => deck.id))
                for (const [id, item] of Object.entries(catalog)) if (!existingDefinitions.has(id) || updatedDefinitions.has(id)) catalogStore.put(item, id)
                for (const [id, data] of Object.entries(images)) if (!existingImages.has(id)) assetStore.put(data, id)
                for (const deck of decks) if (!existingIds.has(deck.id)) deckStore.put(deck, deck.id)
                transaction.objectStore("boards").put({importedAt: new Date().toISOString(), ...counts}, migrationKey)
                result = {catalog, assets: images, decks}
            } catch (error) {operationError = error; transaction.abort()}
        }
        const read = (key, request) => {
            request.onsuccess = () => {latest[key] = request.result === undefined ? false : request.result; process()}
            request.onerror = () => {operationError = request.error; transaction.abort()}
        }
        read("keys", transaction.objectStore("catalog").getAllKeys())
        read("values", transaction.objectStore("catalog").getAll())
        read("assetKeys", transaction.objectStore("assets").getAllKeys())
        read("assetValues", transaction.objectStore("assets").getAll())
        read("decks", transaction.objectStore("decks").getAll())
        read("board", transaction.objectStore("boards").get("current"))
        read("marker", transaction.objectStore("boards").get(migrationKey))
        try {await done} catch (error) {throw operationError || error}
        applySharedStorage(result)
        renderSharedStorage()
        broadcastDeckStorage()
        return counts
    })
}
async function migrateDeckStorage() {
    const transaction = database.transaction(["boards", "catalog", "assets", "decks"], "readwrite"), done = storageTransactionDone(transaction)
    let operationError = null
    const marker = transaction.objectStore("boards").get(storageMigrationKey)
    marker.onsuccess = () => {
        if (marker.result) return
        const previous = transaction.objectStore("boards").get("current")
        previous.onsuccess = () => {
            try {
                const record = previous.result
                if (record) {
                    validateRecord(record)
                    const candidate = upgradeLayout(copy(record.state))
                    checkState(candidate, record.assets, true)
                    const catalog = regularStorageCatalog(candidate.catalog)
                    for (const [id, item] of Object.entries(catalog)) transaction.objectStore("catalog").put(item, id)
                    for (const item of Object.values(catalog)) if (item.asset) transaction.objectStore("assets").put(record.assets[item.asset], item.asset)
                    let decks = record.savedDecks
                    if (decks === undefined) {
                        const now = typeof record.savedAt === "string" && record.savedAt.length <= 40 && Number.isFinite(Date.parse(record.savedAt)) ? record.savedAt : new Date().toISOString()
                        decks = playerIds.map(player => {
                            const channel = candidate.channelDefinitions[player]
                            return {id: `migrated_${player}`, name: `${candidate.players[player].name} のデッキ`.slice(0, 120),
                                list: copy(candidate.decklists[player]).filter(entry => textCardType(catalog[entry.definition]) !== "channel"),
                                channelDefinition: channel && textCardType(catalog[channel]) === "channel" ? channel : "", createdAt: now, updatedAt: now}
                        }).filter(deck => deck.list.length || deck.channelDefinition)
                    }
                    validateSavedDecks(decks, catalog)
                    for (const deck of decks) transaction.objectStore("decks").put(deck, deck.id)
                }
                transaction.objectStore("boards").put(true, storageMigrationKey)
            } catch (error) {operationError = error; transaction.abort()}
        }
    }
    try {await done} catch (error) {throw operationError || error}
}
function connectDeckStorageUpdates() {
    const refresh = () => {
        if (!storageReady) return
        clearTimeout(storageRefreshTimer)
        storageRefreshTimer = setTimeout(() => {refreshDeckStorage().catch(error => console.error("Shared data refresh failed", error))}, 100)
    }
    try {
        if (typeof BroadcastChannel === "function") {
            storageChannel = new BroadcastChannel("nijinana_shared_decks_v2")
            storageChannel.onmessage = event => {if (event.data?.type === "changed") refresh()}
        }
    } catch {}
    // A page restored from the browser's back/forward cache still holds the old in-memory decks.
    window.addEventListener("pageshow", refresh)
    window.addEventListener("focus", refresh)
    document.addEventListener("visibilitychange", () => {if (!document.hidden) refresh()})
}
async function initializeBundledCardStorage() {
    if (typeof bundledDefinitionsCache !== "undefined" && bundledDefinitionsCache === null) return {added: 0, updated: 0, limitReached: false, unavailable: true}
    if (typeof localDeckLegacyEntry !== "undefined" && localDeckLegacyEntry) return {added: 0, updated: 0, limitReached: false}
    if (typeof bundledCardDefinitions !== "function") return {added: 0, updated: 0, limitReached: false}
    if (!database) throw new Error("標準カードの保存先を開けませんでした。")
    const definitions = copy(bundledCardDefinitions())
    if (!Array.isArray(definitions) || definitions.length > 500) throw new Error("標準カードの形式を確認してください。")
    const sourceIds = new Set(), validationState = blankState()
    for (const [index, item] of definitions.entries()) {
        if (!item || typeof item !== "object" || Array.isArray(item) || typeof item.sourceCardKey !== "string" || !/^[A-Za-z0-9_-]{1,100}$/.test(item.sourceCardKey) || ["__proto__", "constructor", "prototype"].includes(item.sourceCardKey) || sourceIds.has(item.sourceCardKey) || item.asset !== "" || isColorDefinition(item)) throw new Error("標準カードのID・画像・種別を確認してください。")
        sourceIds.add(item.sourceCardKey)
        validationState.catalog[`bundled_validation_${index}`] = item
    }
    checkState(validationState, {}, true)
    const transaction = database.transaction(["catalog", "assets", "decks"], "readwrite"), done = storageTransactionDone(transaction)
    const latest = {keys: null, values: null, assetKeys: null, assetValues: null, decks: null}
    let operationError = null, result = null
    const process = () => {
        if (Object.values(latest).some(value => value === null)) return
        try {
            const catalog = Object.fromEntries(latest.keys.map((key, index) => [key, latest.values[index]]))
            const updated = backfillBundledCardTags(catalog, definitions)
            const registered = new Set(Object.values(catalog).map(item => item.sourceCardKey).filter(Boolean))
            const missing = definitions.filter(item => !registered.has(item.sourceCardKey))
            const limitReached = Object.values(catalog).filter(item => !isColorDefinition(item)).length + missing.length > 500
            const additions = []
            for (const item of limitReached ? [] : missing) {
                const base = `builtin_${item.sourceCardKey.replaceAll("-", "_")}`
                let id = base.slice(0, 90), index = 2
                while (Object.hasOwn(catalog, id)) {const suffix = `_${index++}`; id = `${base.slice(0, 90 - suffix.length)}${suffix}`}
                catalog[id] = item
                additions.push({id, item})
            }
            const candidate = blankState()
            Object.assign(candidate.catalog, catalog)
            const images = Object.fromEntries(latest.assetKeys.map((key, index) => [key, latest.assetValues[index]]))
            checkState(candidate, images, true)
            validateSavedDecks(latest.decks, catalog)
            const catalogStore = transaction.objectStore("catalog")
            for (const {id, item} of [...updated, ...additions]) catalogStore.put(item, id)
            result = {added: additions.length, updated: updated.length, limitReached, ...(limitReached ? {missing: missing.length} : {})}
        } catch (error) {operationError = error; transaction.abort()}
    }
    const read = (key, request) => {
        request.onsuccess = () => {latest[key] = request.result; process()}
        request.onerror = () => {operationError = request.error; transaction.abort()}
    }
    read("keys", transaction.objectStore("catalog").getAllKeys())
    read("values", transaction.objectStore("catalog").getAll())
    read("assetKeys", transaction.objectStore("assets").getAllKeys())
    read("assetValues", transaction.objectStore("assets").getAll())
    read("decks", transaction.objectStore("decks").getAll())
    try {await done} catch (error) {throw operationError || error}
    return result
}
async function initializeBundledCardsAfterTransfer() {
    if (!storageReady || !database) return {added: 0, updated: 0, failed: true}
    return queueStorageOperation(async () => {
        try {
            const result = await initializeBundledCardStorage()
            if (result.limitReached) notify("カード登録が500種類を超えるため、標準カードを追加できませんでした。登録を整理してから開き直してください。", true)
            if (result.added || result.updated) {
                applySharedStorage(await readSharedStorage())
                renderSharedStorage()
                broadcastDeckStorage()
            }
            return result
        } catch (error) {
            console.error("Bundled card registration failed", error)
            notify(`標準カードを登録できませんでした。${error.message || String(error)}`, true)
            return {added: 0, updated: 0, failed: true}
        }
    })
}
function initializeStorage() {
    if (storageInitialization) return storageInitialization
    storageFailed = false
    storageInitialization = new Promise((resolve, reject) => {
        let request
        try {request = indexedDB.open("nijinana_manual_table_v1", 2)} catch (error) {reject(error); return}
        const timeout = setTimeout(() => {if (!storageReady) storageStatus("保存先を確認できません。JSON保存を使用してください。")}, 4000)
        request.onupgradeneeded = () => {
            for (const name of ["boards", "catalog", "assets", "decks"]) if (!request.result.objectStoreNames.contains(name)) request.result.createObjectStore(name)
        }
        request.onblocked = () => storageStatus("以前の画面を閉じてから読み込み直してください。保存先の更新を待っています。")
        request.onerror = () => {clearTimeout(timeout); reject(request.error || new Error("保存先を開けませんでした。"))}
        request.onsuccess = async () => {
            database = request.result
            database.onversionchange = () => {
                storageReady = false
                storageFailed = true
                database.close()
                database = null
                storageStatus("保存先を更新しました。この画面を読み込み直してください。")
                renderSharedStorage()
            }
            try {
                await migrateDeckStorage()
                if (!(typeof localDeckTransferPending !== "undefined" && localDeckTransferPending)) {
                    try {
                        const bundled = await initializeBundledCardStorage()
                        if (bundled.limitReached) notify("カード登録が500種類を超えるため、標準カードを追加できませんでした。登録を整理してから開き直してください。", true)
                    } catch (error) {
                        console.error("Bundled card registration failed", error)
                        notify(`標準カードを登録できませんでした。${error.message || String(error)}`, true)
                    }
                }
                const transaction = database.transaction("boards", "readonly"), done = storageTransactionDone(transaction)
                const [record] = await Promise.all([storageReadRequest(transaction.objectStore("boards").get("current")), done])
                let restored = false
                if (record && !ui.touched) {
                    validateRecord(record)
                    const restoredState = upgradeLayout(copy(record.state))
                    checkState(restoredState, record.assets, true)
                    state = restoredState
                    assets = copy(record.assets)
                    normalizePlayerView()
                    restoreReplayFromBoard(record)
                    ui.movePlayer = state.active
                    restored = true
                }
                const shared = await readSharedStorage()
                storageCatalogBaseline = restored ? copy(regularStorageCatalog(state.catalog)) : {}
                applySharedStorage(shared)
                storageReady = true
                connectDeckStorageUpdates()
                renderSharedStorage()
                clearTimeout(timeout)
                storageStatus(restored ? "このブラウザ内のカード・デッキを復元" : "ブラウザ内保存が利用可能 / バックアップはJSON保存")
                if (ui.touched) scheduleSave()
                resolve(true)
            } catch (error) {clearTimeout(timeout); reject(error)}
        }
    }).catch(error => {
        storageReady = false
        storageFailed = true
        console.error("Storage initialization failed", error)
        storageStatus("自動保存を利用できません。JSON保存を使用してください。")
        if (database) {database.close(); database = null}
        try {renderSharedStorage()} catch (renderError) {console.error("Storage error display failed", renderError)}
        return false
    })
    return storageInitialization
}
