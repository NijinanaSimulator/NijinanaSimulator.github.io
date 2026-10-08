// Both screens use one entry document; local files also keep a bridge to their previous save area.
const localDeckTransferMarker = 'local_file_deck_import_v1'
const localDeckBridgePrefix = '#local-deck-transfer='
const localDeckFile = location.protocol === 'file:'
const appDeckEntry = document.body.dataset.page === 'deck'
const appRootEntry = !appDeckEntry
const appDeckRedirect = !localDeckFile && appDeckEntry
const localDeckLegacyEntry = localDeckFile && appDeckEntry
const localDeckBridge = localDeckLegacyEntry && location.hash.startsWith(localDeckBridgePrefix)
const localDeckRoot = localDeckFile && !localDeckLegacyEntry
let localDeckTransferPending = localDeckRoot
let localDeckTransferPromise = null
let localDeckHashNavigation = false

function appRootURL(hash = '') {
    const relative = localDeckFile ? (appDeckEntry ? '../index.html' : 'index.html') : (appDeckEntry ? '../' : './')
    const url = new URL(relative, location.href)
    url.hash = hash
    return url.href
}
function appBoardURL() {return appRootURL()}
function appDeckManagerURL() {return appRootURL('deck')}

if (appRootEntry && location.hash === '#deck') {
    document.body.dataset.page = 'deck'
    document.title = 'デッキ管理 | NijinanaSimulator'
    document.querySelector('.header-actions').innerHTML = `<a class="app-link" href="${appRootURL()}">シミュレーターへ</a><button data-action="catalog">カード登録</button><button data-action="save">JSON保存</button><button data-action="load">読込</button>`
    document.querySelector('.subbar > span').textContent = 'デッキ管理'
    for (const selector of ['.layout', '.mobile-panel-backdrop', '#mobileNav', '#replayCutIn']) document.querySelector(selector)?.remove()
    const root = document.createElement('main')
    root.id = 'deckPage'; root.className = 'deck-page'; root.setAttribute('aria-label', 'デッキ管理')
    document.body.insertBefore(root, document.getElementById('modal'))
}
if (localDeckLegacyEntry && !localDeckBridge) {
    document.querySelector('.header-actions').innerHTML = `<a class="app-link" href="${appRootURL('deck')}">デッキ管理へ</a><button data-action="save">旧データをJSON保存</button>`
}

function renderLocalDeckEntry() {
    if (localDeckBridge || appDeckRedirect) return true
    if (localDeckLegacyEntry) {
        const root = deckManagerRoot()
        const decks = storageReady ? savedDecks.map(deck => `<li>${escapeHTML(deck.name)}（${deckManagerTotal(deck)}枚）</li>`).join('') : ''
        root.innerHTML = `<section class="dm-local-entry"><h1>デッキ管理の場所が変わりました</h1><p>シミュレーターと同じ保存先を使うため、デッキ管理をまとめました。下のボタンから開くと、ここに保存したデッキとカードを引き継ぎます。</p>${decks ? `<ul>${decks}</ul>` : ''}<div class="control-row"><a class="dm-link-button primary" href="${appRootURL('deck')}">デッキ管理を開く</a><button data-action="save" ${storageReady ? '' : 'disabled'}>旧データをJSON保存</button></div><p class="muted">元の保存データは残ります。</p></section>`
        return true
    }
    if (localDeckRoot && localDeckTransferPending && appPage === 'deck') {
        deckManagerRoot().innerHTML = '<div class="dm-loading" role="status">保存したカードとデッキを引き継いでいます…</div>'
        return true
    }
    return false
}
function renderLocalDeckTransferNotice(error) {
    document.getElementById('localDeckTransferNotice')?.remove()
    if (!error) return
    const panel = document.createElement('section')
    panel.id = 'localDeckTransferNotice'; panel.className = 'dm-local-transfer-notice'; panel.setAttribute('role', 'status')
    panel.innerHTML = `<p>以前のデッキ管理からデータを引き継げませんでした。${escapeHTML(error.message || String(error))}</p><div class="control-row"><button data-action="local-deck-transfer-retry">もう一度引き継ぐ</button><a class="dm-link-button" href="deck/index.html">旧データを開く</a></div>`
    const anchor = document.querySelector('.layout') || document.getElementById('deckPage') || document.getElementById('modal')
    document.body.insertBefore(panel, anchor)
}
async function serveLocalDeckTransfer() {
    const nonce = location.hash.slice(localDeckBridgePrefix.length)
    if (!localDeckBridge || window.parent === window || !/^[A-Za-z0-9_]{1,90}$/.test(nonce)) return
    try {
        if (!(await initializeStorage())) throw new Error('旧データの保存先を開けませんでした。')
        const snapshot = await readSharedStorage()
        window.parent.postMessage({type: 'nijinana-local-deck-transfer', nonce, snapshot}, '*')
    } catch (error) {
        window.parent.postMessage({type: 'nijinana-local-deck-transfer', nonce, error: error.message || String(error)}, '*')
    }
}
function requestLocalDeckSnapshot() {
    return new Promise((resolve, reject) => {
        const nonce = uid('transfer'), frame = document.createElement('iframe')
        frame.hidden = true; frame.title = '保存デッキの引き継ぎ'
        const target = new URL('deck/index.html', location.href)
        target.hash = `${localDeckBridgePrefix.slice(1)}${nonce}`
        const cleanup = () => {clearTimeout(timer); window.removeEventListener('message', receive); frame.remove()}
        const receive = event => {
            if (event.source !== frame.contentWindow || event.data?.type !== 'nijinana-local-deck-transfer' || event.data.nonce !== nonce) return
            const data = event.data; cleanup()
            if (data.error) reject(new Error(String(data.error)))
            else resolve(data.snapshot)
        }
        const timer = setTimeout(() => {cleanup(); reject(new Error('旧データの読み込みが完了しませんでした。旧データ画面のJSON保存も利用できます。'))}, 12000)
        window.addEventListener('message', receive)
        frame.addEventListener('error', () => {cleanup(); reject(new Error('旧データの画面を読み込めませんでした。'))}, {once: true})
        frame.src = target.href; document.body.appendChild(frame)
    })
}
function initializeLocalDeckTransfer() {
    if (!localDeckRoot || !storageReady) return Promise.resolve(false)
    if (localDeckTransferPromise) return localDeckTransferPromise
    localDeckTransferPending = true; renderLocalDeckTransferNotice(null); render()
    localDeckTransferPromise = (async () => {
        const transaction = database.transaction('boards', 'readonly'), done = storageTransactionDone(transaction)
        const [alreadyImported] = await Promise.all([storageReadRequest(transaction.objectStore('boards').get(localDeckTransferMarker)), done])
        if (alreadyImported) return true
        const snapshot = await requestLocalDeckSnapshot()
        if (!snapshot.decks?.length && !Object.keys(snapshot.catalog || {}).length) return true
        const result = await importLocalDeckStorage(snapshot, localDeckTransferMarker)
        if (result.decks) notify(`${result.decks}個の保存デッキを引き継ぎました。`)
        return true
    })().catch(error => {
        console.error('Local deck transfer failed', error)
        renderLocalDeckTransferNotice(error)
        return false
    }).finally(() => {localDeckTransferPending = false; localDeckTransferPromise = null; renderSharedStorage()})
    return localDeckTransferPromise
}
async function waitLocalDeckTransfer() {if (localDeckRoot && localDeckTransferPending) await (localDeckTransferPromise || initializeLocalDeckTransfer())}
function restoreLocalDeckHash() {
    const url = appRootURL(appPage === 'deck' ? 'deck' : '')
    history.replaceState(null, '', url)
}
window.addEventListener('hashchange', async () => {
    if (!appRootEntry || localDeckHashNavigation || (location.hash === '#deck' ? 'deck' : 'board') === appPage) return
    localDeckHashNavigation = true
    try {
        if (appPage === 'deck' && !(await confirmDeckManagerLeave())) {restoreLocalDeckHash(); return}
        await waitLocalDeckTransfer()
        if (storageReady) await flushStorageSave()
        if (appPage === 'deck') discardDeckManagerDraft()
        location.reload()
    } catch (error) {
        restoreLocalDeckHash(); notify(`移動できませんでした。${error.message}`, true)
    } finally {localDeckHashNavigation = false}
})
