/* デッキ TXT v1: カードの公開 ID / 名前で共有し、内部 ID は出力しません。 */
const deckTextHeader = '# NijinanaSimulator Deck TXT v1'
const deckTextMaxBytes = 1024 * 1024
const deckTextMaxLines = 1000

function deckTextEncodeField(value) {
    return String(value).replace(/\\/g, '\\\\').replace(/\t/g, '\\t').replace(/\r/g, '\\r').replace(/\n/g, '\\n')
}
function deckTextDecodeField(value, line) {
    let result = ''
    for (let index = 0; index < value.length; index++) {
        const character = value[index]
        if (character !== '\\') {result += character; continue}
        const escape = value[++index]
        if (escape === '\\') result += '\\'
        else if (escape === 't') result += '\t'
        else if (escape === 'r') result += '\r'
        else if (escape === 'n') result += '\n'
        else throw new Error(`${line}行目: 文字のエスケープが正しくありません。`)
    }
    if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(result)) throw new Error(`${line}行目: 使用できない制御文字があります。`)
    return result
}
function deckTextRequire(condition, message, line) {
    if (!condition) throw new Error(`${line ? `${line}行目: ` : ''}${message}`)
}
function deckTextCardType(item) {return item.cardType ?? (item.isLiver ? 'liver' : 'other')}
function deckTextValidCardId(value) {
    return value === '' || (/^[A-Za-z0-9_-]{1,100}$/.test(value) && !['__proto__', 'constructor', 'prototype'].includes(value))
}
function deckTextValidateName(name, maximum, label, line) {
    deckTextRequire(typeof name === 'string' && name.trim().length > 0 && name.length <= maximum, `${label}は1〜${maximum}文字で指定してください。`, line)
    deckTextRequire(!/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(name), `${label}に使用できない制御文字があります。`, line)
}
function deckTextCheckSize(text) {
    deckTextRequire(new TextEncoder().encode(text).byteLength <= deckTextMaxBytes, 'TXTファイルは1MB以内にしてください。')
}

function buildDeckText(deck, catalog) {
    deckTextRequire(deck && typeof deck === 'object' && catalog && typeof catalog === 'object' && !Array.isArray(catalog), 'デッキまたはカード登録の形式が正しくありません。')
    deckTextValidateName(deck.name, 120, 'デッキ名')
    deckTextRequire(Array.isArray(deck.list) && deck.list.length <= 500, 'デッキのカード一覧が正しくありません。')
    const row = (definition, channel) => {
        deckTextRequire(typeof definition === 'string' && Object.hasOwn(catalog, definition), 'デッキに未登録のカードがあります。')
        const item = catalog[definition]
        deckTextRequire(item && item.kind !== 'color', 'カラーカードはTXTのデッキに含められません。')
        deckTextRequire((deckTextCardType(item) === 'channel') === channel, channel ? 'チャンネルにはチャンネルカードを指定してください。' : 'チャンネルカードは通常デッキに含められません。')
        deckTextValidateName(item.name, 200, 'カード名')
        const id = item.sourceCardKey ?? ''
        deckTextRequire(typeof id === 'string' && deckTextValidCardId(id), 'カードIDの形式が正しくありません。')
        return [deckTextEncodeField(id), deckTextEncodeField(item.name)]
    }
    const lines = [deckTextHeader, `デッキ名\t${deckTextEncodeField(deck.name)}`, '', '[チャンネル]', 'ID\tカード名']
    deckTextRequire(deck.channelDefinition === '' || typeof deck.channelDefinition === 'string', 'チャンネルカードを確認してください。')
    if (deck.channelDefinition) lines.push(row(deck.channelDefinition, true).join('\t'))
    lines.push('', '[デッキ]', 'ID\t枚数\tカード名')
    let total = 0
    for (const entry of deck.list) {
        deckTextRequire(entry && Number.isInteger(entry.count) && entry.count >= 1 && entry.count <= 500, 'カードの枚数は1〜500枚で指定してください。')
        total += entry.count
        deckTextRequire(total <= 500, 'デッキは500枚までです。')
        const [id, name] = row(entry.definition, false)
        lines.push(`${id}\t${entry.count}\t${name}`)
    }
    const text = `${lines.join('\r\n')}\r\n`
    deckTextCheckSize(text)
    return text
}

function parseDeckText(text, catalog) {
    deckTextRequire(typeof text === 'string', 'TXTファイルを確認してください。')
    deckTextCheckSize(text)
    deckTextRequire(catalog && typeof catalog === 'object' && !Array.isArray(catalog), 'カード登録の形式が正しくありません。')
    const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/)
    deckTextRequire(lines.length <= deckTextMaxLines, `TXTファイルは${deckTextMaxLines}行以内にしてください。`)
    const carriageReturnLine = lines.findIndex(line => line.includes('\r'))
    deckTextRequire(carriageReturnLine < 0, '改行は LF または CRLF で指定し、カード名の改行はエスケープしてください。', carriageReturnLine + 1)
    deckTextRequire(lines[0] === deckTextHeader, 'NijinanaSimulator Deck TXT v1 のファイルを選んでください。', 1)
    const nameFields = (lines[1] ?? '').split('\t')
    deckTextRequire(nameFields.length === 2 && nameFields[0] === 'デッキ名', 'デッキ名の行が正しくありません。', 2)
    const name = deckTextDecodeField(nameFields[1], 2)
    deckTextValidateName(name, 120, 'デッキ名', 2)
    const issues = [], list = [], entries = new Map(), catalogEntries = Object.entries(catalog)
    let index = 2, channelDefinition = '', total = 0
    const skipEmpty = () => {while (lines[index] === '') index++}
    const expect = expected => {
        skipEmpty()
        deckTextRequire(lines[index] === expected, `「${expected.replace(/\t/g, ' / ')}」の行が必要です。`, index + 1)
        index++
    }
    const parseIdentity = (fields, line, channel) => {
        const id = deckTextDecodeField(fields[0], line), cardName = deckTextDecodeField(fields[channel ? 1 : 2], line)
        deckTextRequire(deckTextValidCardId(id), 'カードIDの形式が正しくありません。', line)
        deckTextRequire(cardName.length <= 200, 'カード名は200文字以内で指定してください。', line)
        deckTextRequire(id !== '' || cardName.trim() !== '', 'カードIDまたはカード名を指定してください。', line)
        const matches = catalogEntries.filter(([, item]) => item && (id ? item.sourceCardKey === id : item.name === cardName))
        const validMatches = matches.filter(([, item]) => item.kind !== 'color' && (deckTextCardType(item) === 'channel') === channel)
        if (matches.length && !validMatches.length) {
            if (matches.every(([, item]) => item.kind === 'color')) throw new Error(`${line}行目: カラーカードはTXTのデッキに含められません。`)
            throw new Error(`${line}行目: ${channel ? 'チャンネルにはチャンネルカードを指定してください。' : 'チャンネルカードは通常デッキに含められません。'}`)
        }
        let candidates = validMatches
        if (id && candidates.length > 1 && cardName !== '') candidates = candidates.filter(([, item]) => item.name === cardName)
        if (candidates.length === 1) return candidates[0][0]
        const label = id ? `ID「${id}」${cardName ? `（${cardName}）` : ''}` : `カード名「${cardName}」`
        issues.push({line, message: `${label}${matches.length ? 'の登録が複数あり、カードを特定できません。' : 'のカードが未登録です。カード登録から追加してください。'}`})
        return ''
    }
    expect('[チャンネル]')
    expect('ID\tカード名')
    skipEmpty()
    if (lines[index] !== '[デッキ]') {
        const line = index + 1, fields = (lines[index] ?? '').split('\t')
        deckTextRequire(fields.length === 2, 'チャンネルの行は「ID / カード名」で指定してください。', line)
        channelDefinition = parseIdentity(fields, line, true)
        index++
    }
    expect('[デッキ]')
    expect('ID\t枚数\tカード名')
    for (; index < lines.length; index++) {
        if (lines[index] === '') continue
        const line = index + 1, fields = lines[index].split('\t')
        deckTextRequire(fields.length === 3, 'カードの行は「ID / 枚数 / カード名」で指定してください。', line)
        deckTextRequire(/^[0-9]+$/.test(fields[1]), '枚数は1〜500の整数で指定してください。', line)
        const count = Number(fields[1])
        deckTextRequire(Number.isSafeInteger(count) && count >= 1 && count <= 500, '枚数は1〜500の整数で指定してください。', line)
        total += count
        deckTextRequire(total <= 500, 'デッキは500枚までです。', line)
        const definition = parseIdentity(fields, line, false)
        if (!definition) continue
        if (entries.has(definition)) entries.get(definition).count += count
        else {const entry = {definition, count}; entries.set(definition, entry); list.push(entry)}
    }
    return {name, list, channelDefinition, issues}
}
