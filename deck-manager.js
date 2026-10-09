// Named deck editor. Cards and images come from the shared application storage.
const deckManager = {initialized:false,view:'gallery',draft:null,baseline:'',selected:'',query:'',type:'',className:'',colors:new Set(),sort:'registered',sortDirection:'asc',page:0,catalogMode:'main',saving:false,exporting:false,exportEpoch:0,pickerPlayer:'p1',textImporting:false,pickerLoading:false,pickerError:'',pickerEpoch:0}
const deckManagerSorts={registered:'登録順',name:'カード名',id:'ID',cost:'レベル',power:'パワー'}
const deckManagerSortDirections={asc:'昇順',desc:'降順'}
const deckManagerPageSize=18
const deckManagerDraftKey='nijinana_deck_editor_draft_v1'
const deckManagerCollator=new Intl.Collator('ja',{numeric:true,sensitivity:'base'})
const deckManagerColorNames={R:'赤',O:'橙',Y:'黄',G:'緑',B:'青',I:'藍',V:'紫'}
function deckManagerEntries(){return Object.entries(state.catalog).filter(([,item])=>!isColorDefinition(item))}
function deckManagerCount(value){const n=Number(value);return Number.isFinite(n)?Math.max(0,Math.min(500,Math.trunc(n))):0}
function deckManagerDecks(){return Array.isArray(savedDecks)?savedDecks:Object.values(savedDecks||{})}
function deckManagerMain(id){return !!state.catalog[id]&&!isColorDefinition(state.catalog[id])&&textCardType(state.catalog[id])!=='channel'}
function deckManagerChannel(id){return !!state.catalog[id]&&!isColorDefinition(state.catalog[id])&&textCardType(state.catalog[id])==='channel'}
function deckManagerTotal(deck){return (deck?.list||[]).reduce((sum,item)=>sum+deckManagerCount(item.count),0)}
function deckManagerSignature(deck){return JSON.stringify({name:deck?.name||'',channelDefinition:deck?.channelDefinition||'',list:deck?.list||[]})}
function deckManagerDirty(){return !!deckManager.draft&&deckManagerSignature(deckManager.draft)!==deckManager.baseline}
function deckManagerUsesDefinition(id){return !!deckManager.draft&&(deckManager.draft.channelDefinition===id||deckManager.draft.list.some(item=>item.definition===id))}
function resetDeckManagerAfterImport(){deckManagerTextReadEpoch++;deckManagerTextImport=null;deckManager.textImporting=false;deckManager.draft=null;deckManager.baseline='';deckManager.selected='';deckManager.view='gallery';deckManager.initialized=true;try{sessionStorage.removeItem(deckManagerDraftKey)}catch{}}
function deckManagerRemember(){
 if(appPage!=='deck')return
 try{sessionStorage.removeItem(deckManagerDraftKey)}catch{}
 const needed=deckManagerDirty()||deckManager.saving
 if(needed===deckManagerUnloadGuardActive)return
 deckManagerUnloadGuardActive=needed
 if(needed)globalThis.addEventListener?.('beforeunload',deckManagerBeforeUnload)
 else globalThis.removeEventListener?.('beforeunload',deckManagerBeforeUnload)
}
function deckManagerInitialize(){
 if(deckManager.initialized||!storageReady)return
 deckManager.initialized=true
 try{sessionStorage.removeItem(deckManagerDraftKey)}catch{}
}
function deckManagerRoot(){let root=byId('deckPage');if(!root){root=document.createElement('main');root.id='deckPage';document.body.insertBefore(root,byId('modal')||null)}root.classList.add('dm-page');return root}
function deckDefinitionFaceHTML(definition){
 const item=typeof definition==='string'?state.catalog[definition]:definition
 if(!item)return '<span class="dm-card-face card"><span class="dm-empty-symbol">◇</span></span>'
 const picture=item.asset?`<img src="${escapeHTML(cachedCardImageURL(item.asset))}" alt="${escapeHTML(item.name)}" draggable="false" loading="lazy" decoding="async">`:''
 return `<span class="dm-card-face card ${picture?'':'text-card'}">${picture?liverImageFaceHTML(item,picture):textCardFaceHTML(item)}</span>`
}
function deckManagerCardSummary(item){
 const type=textCardType(item),parts=[textCardTypes[type]||'カード']
 if(type!=='channel'&&item.cost!==null&&item.cost!==undefined)parts.push(`レベル ${item.cost}`)
 if(item.className)parts.push(item.className)
 if(type==='liver'&&item.power!==null&&item.power!==undefined)parts.push(`パワー ${item.power}`)
 return parts.join(' · ')
}
function deckManagerCoverHTML(deck){
 const cards=(deck.list||[]).filter(item=>deckManagerMain(item.definition)).slice(0,3).map((entry,i)=>`<span class="dm-cover-card dm-cover-card-${i}">${deckDefinitionFaceHTML(entry.definition)}</span>`).join('')
 const channel=deckManagerChannel(deck.channelDefinition)?`<span class="dm-cover-channel">${deckDefinitionFaceHTML(deck.channelDefinition)}</span>`:''
 return `<span class="dm-cover ${cards?'':'dm-cover-empty'}">${cards||'<span class="dm-empty-symbol">◇</span>'}${channel}</span>`
}
function deckManagerGalleryHTML(){
 const decks=deckManagerDecks().slice().sort((a,b)=>String(b.updatedAt||'').localeCompare(String(a.updatedAt||'')))
 const tiles=decks.map(deck=>{
  const channel=state.catalog[deck.channelDefinition]
  return `<article class="dm-deck-tile"><button class="dm-deck-open" data-action="manager-edit" data-deck-id="${escapeHTML(deck.id)}">${deckManagerCoverHTML(deck)}<strong>${escapeHTML(deck.name||'無題のデッキ')}</strong><span class="dm-deck-meta">${deckManagerTotal(deck)}枚 · ${(deck.list||[]).length}種類</span><span class="dm-deck-channel">${channel?escapeHTML(channel.name):'チャンネル未設定'}</span></button><div class="dm-tile-actions"><button data-action="manager-edit" data-deck-id="${escapeHTML(deck.id)}">編集</button><button data-action="manager-duplicate" data-deck-id="${escapeHTML(deck.id)}">複製</button><button class="dm-delete" data-action="manager-delete" data-deck-id="${escapeHTML(deck.id)}" aria-label="${escapeHTML(deck.name)}を削除">削除</button></div></article>`
 }).join('')
 return `<section class="dm-gallery"><div class="dm-page-title"><div><p class="dm-eyebrow">NijinanaSimulator</p><h2>デッキ</h2><p class="dm-description">保存したデッキを選んで、編集できます。</p></div><div class="dm-gallery-actions"><button data-action="manager-txt-import" ${deckManager.textImporting?'disabled':''}>${deckManager.textImporting?'読み込み中…':'TXT読込'}</button><button data-action="catalog">カード登録</button></div></div><div class="dm-gallery-grid"><button class="dm-new-deck" data-action="manager-new"><span aria-hidden="true">＋</span><strong>新しいデッキ</strong></button>${tiles}</div>${decks.length?'':'<p class="dm-gallery-note">カード登録で用意したカードを使って、デッキを作成できます。</p>'}</section>`
}
function deckManagerDetailHTML(){
 const draft=deckManager.draft,id=deckManager.selected,item=state.catalog[id]
 if(!item)return '<div class="dm-detail-empty"><span aria-hidden="true">◇</span><strong>カードを選択</strong><p>カードの内容をここに表示します。</p></div>'
 const channel=textCardType(item)==='channel',quantity=channel?Number(draft.channelDefinition===id):draft.list.find(entry=>entry.definition===id)?.count||0
 const colors=Array.isArray(item.colors)?item.colors.map(color=>`<span class="dm-color-dot dm-color-${color.toLowerCase()}" title="${deckManagerColorNames[color]||color}" aria-label="${deckManagerColorNames[color]||color}"></span>`).join(''):''
 return `<div class="dm-detail-heading"><strong>${escapeHTML(item.name)}</strong>${item.sourceCardKey?`<span class="dm-detail-id">ID: ${escapeHTML(item.sourceCardKey)}</span>`:''}</div><div class="dm-detail-image ${channel?'dm-detail-channel':''}">${deckDefinitionFaceHTML(item)}</div><p class="dm-detail-summary">${escapeHTML(deckManagerCardSummary(item))}</p>${cardTagsHTML(item,'dm-detail-tags')}${colors?`<div class="dm-detail-colors">${colors}</div>`:''}<div class="dm-detail-actions">${channel?`<button data-action="manager-set-channel" data-definition="${escapeHTML(id)}" class="primary">${quantity?'選択中のチャンネル':'チャンネルに設定'}</button>${quantity?'<button data-action="manager-clear-channel">チャンネルを外す</button>':''}`:`<button data-action="manager-minus" data-definition="${escapeHTML(id)}" ${quantity?'':'disabled'}>−1</button><span>${quantity}枚</span><button data-action="manager-plus" data-definition="${escapeHTML(id)}" class="primary">＋1</button>`}</div><div class="dm-detail-effect"><h3>効果・説明</h3><p>${escapeHTML(item.text||'効果・説明は未登録です。')}</p></div>`
}
function deckManagerDeckContentsHTML(){
 const draft=deckManager.draft,channel=deckManagerChannel(draft.channelDefinition)?state.catalog[draft.channelDefinition]:null,entries=draft.list.filter(entry=>deckManagerMain(entry.definition)&&entry.count>0)
 const tiles=entries.map(entry=>{
  const item=state.catalog[entry.definition],id=escapeHTML(entry.definition)
  return `<article class="dm-included-card ${deckManager.selected===entry.definition?'is-selected':''}" data-manager-entry="${id}"><button class="dm-select-card" draggable="true" data-manager-drag-definition="${id}" data-manager-drag-origin="deck" data-action="manager-select" data-definition="${id}" aria-label="${escapeHTML(item.name)}の内容">${deckDefinitionFaceHTML(item)}<span class="dm-card-name">${escapeHTML(item.name)}</span></button><div class="dm-quantity"><button data-action="manager-minus" data-definition="${id}" aria-label="${escapeHTML(item.name)}を1枚減らす">−</button><input data-manager-quantity="${id}" type="number" min="0" max="500" value="${entry.count}" aria-label="${escapeHTML(item.name)}の枚数"><button data-action="manager-plus" data-definition="${id}" aria-label="${escapeHTML(item.name)}を1枚増やす">＋</button></div></article>`
 }).join('')
 return `<section class="dm-channel-section" data-manager-drop="channel"><div class="dm-section-heading"><h3>チャンネル</h3><button class="small" data-action="manager-channel-mode">${channel?'変更':'選ぶ'}</button></div>${channel?`<div class="dm-channel-selection"><button data-action="manager-select" data-definition="${escapeHTML(draft.channelDefinition)}" draggable="true" data-manager-drag-definition="${escapeHTML(draft.channelDefinition)}" data-manager-drag-origin="channel" class="dm-channel-card">${deckDefinitionFaceHTML(channel)}</button><div><strong>${escapeHTML(channel.name)}</strong>${channel.sourceCardKey?`<small>ID: ${escapeHTML(channel.sourceCardKey)}</small>`:''}<button class="small" data-action="manager-clear-channel">外す</button></div></div>`:'<button class="dm-channel-empty" data-action="manager-channel-mode"><span>＋</span> チャンネルを選択</button>'}</section><section class="dm-main-section" data-manager-drop="main"><div class="dm-section-heading"><h3>デッキ <span>${deckManagerTotal(draft)}枚</span></h3><span>${entries.length}種類</span></div><div class="dm-deck-grid">${tiles||'<div class="dm-deck-empty"><span aria-hidden="true">＋</span><p>右のカード一覧から、カードを追加してください。</p></div>'}</div></section>`
}
function deckManagerFilteredCatalog(){
 const normalize=value=>String(value||'').normalize('NFKC').toLocaleLowerCase('ja'),terms=normalize(deckManager.query).trim().split(/\s+/).filter(Boolean)
 const entries=deckManagerEntries().filter(([,item])=>(deckManager.catalogMode==='channel'?textCardType(item)==='channel':textCardType(item)!=='channel')&&(!deckManager.type||textCardType(item)===deckManager.type)&&(!deckManager.className||item.className===deckManager.className)&&[...deckManager.colors].every(color=>item.colors?.includes(color))&&terms.every(term=>normalize(`${item.name} ${item.sourceCardKey||''} ${item.text||''} ${liverCardTags(item).map(tag=>`#${tag}`).join(' ')}`).includes(term)))
 const direction=deckManager.sortDirection==='desc'?-1:1
 if(deckManager.sort==='registered')return direction===1?entries:entries.reverse()
 return entries.sort((a,b)=>{
  if(deckManager.sort==='name')return direction*deckManagerCollator.compare(a[1].name,b[1].name)
  const field=deckManager.sort==='id'?'sourceCardKey':deckManager.sort==='power'?'power':'cost',left=a[1][field],right=b[1][field]
  const missing=value=>value===null||value===undefined||(field==='sourceCardKey'&&value==='')
  if(missing(left))return missing(right)?0:1
  if(missing(right))return -1
  return direction*(field==='sourceCardKey'?deckManagerCollator.compare(left,right):left-right)
 })
}
function deckManagerCatalogHTML(){
 const entries=deckManagerFilteredCatalog(),pages=Math.max(1,Math.ceil(entries.length/deckManagerPageSize));deckManager.page=Math.max(0,Math.min(pages-1,deckManager.page));const start=deckManager.page*deckManagerPageSize
 const tiles=entries.slice(start,start+deckManagerPageSize).map(([id,item])=>{
  const quantity=deckManager.catalogMode==='channel'?Number(deckManager.draft.channelDefinition===id):deckManager.draft.list.find(entry=>entry.definition===id)?.count||0
  return `<article class="dm-catalog-card ${deckManager.selected===id?'is-selected':''}"><button class="dm-select-card" draggable="true" data-manager-drag-definition="${escapeHTML(id)}" data-action="manager-select" data-definition="${escapeHTML(id)}" aria-label="${escapeHTML(item.name)}の内容">${deckDefinitionFaceHTML(item)}<span class="dm-card-name">${escapeHTML(item.name)}</span>${item.sourceCardKey?`<small class="dm-card-id">${escapeHTML(item.sourceCardKey)}</small>`:''}</button><div class="dm-catalog-card-footer">${deckManager.catalogMode==='channel'?`<button class="small ${quantity?'dm-added':''}" data-action="manager-set-channel" data-definition="${escapeHTML(id)}">${quantity?'選択中':'設定'}</button>`:`<button class="small" data-action="manager-minus" data-definition="${escapeHTML(id)}" aria-label="${escapeHTML(item.name)}を1枚減らす" ${quantity?'':'disabled'}>−</button><span>${quantity}枚</span><button class="small" data-action="manager-plus" data-definition="${escapeHTML(id)}" aria-label="${escapeHTML(item.name)}を1枚追加">＋</button>`}</div></article>`
 }).join('')
 return `<div class="dm-catalog-count">${entries.length}種類${entries.length?` · ${start+1}〜${Math.min(start+deckManagerPageSize,entries.length)}`:''}</div><div class="dm-catalog-grid">${tiles||`<p class="dm-catalog-empty">${deckManagerEntries().length?'条件に一致するカードがありません。':'カード登録からカードを用意してください。'}</p>`}</div><div class="dm-catalog-pagination"><button class="small" data-action="manager-page-prev" ${deckManager.page?'':'disabled'}>前へ</button><span>${deckManager.page+1} / ${pages}</span><button class="small" data-action="manager-page-next" ${deckManager.page+1<pages?'':'disabled'}>次へ</button></div>`
}
function deckManagerFiltersHTML(){
 const types=Object.entries(textCardTypes).filter(([type])=>type!=='channel').map(([value,label])=>`<option value="${value}" ${deckManager.type===value?'selected':''}>${label}</option>`).join('')
 const classes=textCardClasses.map(value=>`<option value="${value}" ${deckManager.className===value?'selected':''}>${value}</option>`).join('')
 const sorts=Object.entries(deckManagerSorts).map(([value,label])=>`<option value="${value}" ${deckManager.sort===value?'selected':''}>${label}</option>`).join('')
 const directions=Object.entries(deckManagerSortDirections).map(([value,label])=>`<option value="${value}" ${deckManager.sortDirection===value?'selected':''}>${label}</option>`).join('')
 const colors=Object.entries(deckManagerColorNames).map(([value,name])=>`<label class="dm-color-filter" title="${name}"><input type="checkbox" data-manager-color="${value}" ${deckManager.colors.has(value)?'checked':''}><span class="dm-color-dot dm-color-${value.toLowerCase()}"><span>${value}</span></span></label>`).join('')
 return `<div class="dm-catalog-heading"><h3>カード一覧</h3></div><div class="dm-catalog-tabs"><button data-action="manager-main-mode" class="${deckManager.catalogMode==='main'?'active':''}">デッキ用</button><button data-action="manager-channel-mode" class="${deckManager.catalogMode==='channel'?'active':''}">チャンネル</button></div><div class="dm-filters"><label class="dm-search"><span class="sr-only">カード名・ID・タグ・効果で検索</span><input id="managerSearch" type="search" value="${escapeHTML(deckManager.query)}" placeholder="カード名・ID・タグ・効果で検索" autocomplete="off"></label><div class="dm-filter-row">${deckManager.catalogMode==='main'?`<label><span>種別</span><select id="managerType"><option value="">すべて</option>${types}</select></label>`:''}<label><span>クラス</span><select id="managerClass"><option value="">すべて</option>${classes}</select></label><label><span>並び順</span><select id="managerSort">${sorts}</select></label><label class="dm-sort-direction"><span>順序</span><select id="managerSortDirection">${directions}</select></label></div>${deckManager.catalogMode==='main'?`<div class="dm-color-filters"><span>色</span>${colors}<button class="ghost small" data-action="manager-clear-filters">解除</button></div>`:''}</div>`
}
function deckManagerEditorHTML(){
 const draft=deckManager.draft
 return `<div class="dm-editor-toolbar"><button class="dm-back" data-action="manager-list">← デッキ一覧</button><label class="dm-name-label"><span class="sr-only">デッキ名</span><input id="managerDeckName" value="${escapeHTML(draft.name)}" maxlength="120" placeholder="デッキ名" autocomplete="off"></label><span id="managerDirtyStatus" class="dm-save-state">${deckManagerDirty()?'未保存の変更':'保存済み'}</span><div class="dm-editor-actions"><button data-action="manager-export" ${deckManager.exporting?'disabled':''}>${deckManager.exporting?'画像を作成中…':'画像出力'}</button><button data-action="manager-txt-export">TXT出力</button><button data-action="manager-txt-import" ${deckManager.textImporting?'disabled':''}>${deckManager.textImporting?'読み込み中…':'TXT読込'}</button><button data-action="manager-save" class="primary" ${deckManager.saving?'disabled':''}>${deckManager.saving?'保存中…':'デッキを保存'}</button></div></div><div class="dm-editor"><aside id="managerDetail" class="dm-detail" aria-label="選択したカード">${deckManagerDetailHTML()}</aside><div id="managerContents" class="dm-contents">${deckManagerDeckContentsHTML()}</div><aside class="dm-browser" data-manager-drop="catalog" aria-label="カード一覧">${deckManagerFiltersHTML()}<div id="managerCatalog">${deckManagerCatalogHTML()}</div></aside></div>`
}
function renderDeckPage(){
 if(appPage!=='deck')return
 if(renderLocalDeckEntry())return
 deckManagerInitialize();const root=deckManagerRoot()
 if(!storageReady){root.innerHTML=typeof storageFailed!=='undefined'&&storageFailed?`<div class="dm-loading" role="status"><p>このブラウザでは保存領域を利用できません。別のブラウザで開くか、ブラウザの保存設定を確認してください。</p><a class="dm-link-button" href="${appBoardURL()}">盤面に戻る</a></div>`:'<div class="dm-loading" role="status">保存したカードとデッキを読み込み中…</div>';return}
 if(deckManager.draft&&!deckManagerDirty()&&!deckManager.saving){const latest=deckManagerDecks().find(deck=>deck.id===deckManager.draft.id);if(latest){deckManager.draft=copy(latest);deckManager.baseline=deckManagerSignature(latest)}else{deckManager.draft=null;deckManager.view='gallery';try{sessionStorage.removeItem(deckManagerDraftKey)}catch{}}}
 const focused=document.activeElement,focusKey=root.contains(focused)?focused.id||focused.dataset?.managerQuantity:'',quantityFocus=!focused?.id&&focused?.dataset?.managerQuantity,selection=typeof focused?.selectionStart==='number'?[focused.selectionStart,focused.selectionEnd]:null
 root.innerHTML=deckManager.view==='editor'&&deckManager.draft?deckManagerEditorHTML():deckManagerGalleryHTML();root.inert=deckManager.saving;scheduleDeckManagerViewport()
 if(focusKey&&!deckManager.saving){const restored=quantityFocus?[...root.querySelectorAll('[data-manager-quantity]')].find(element=>element.dataset.managerQuantity===focusKey):byId(focusKey);if(restored){restored.focus({preventScroll:true});if(selection&&typeof restored.setSelectionRange==='function')try{restored.setSelectionRange(...selection)}catch{}}}
 deckManagerRemember()
}
function deckManagerRefreshParts({catalog=true,contents=true,detail=true}={}){
 if(deckManager.view!=='editor'||!deckManager.draft)return
 const catalogNode=byId('managerCatalog'),catalogScroll=catalogNode?.querySelector('.dm-catalog-grid')?.scrollTop||0,contentsNode=byId('managerContents'),contentsScroll=contentsNode?.scrollTop||0
 if(catalog&&catalogNode){catalogNode.innerHTML=deckManagerCatalogHTML();const grid=catalogNode.querySelector('.dm-catalog-grid');if(grid)grid.scrollTop=catalogScroll}
 if(contents&&contentsNode){contentsNode.innerHTML=deckManagerDeckContentsHTML();contentsNode.scrollTop=contentsScroll}
 if(detail&&byId('managerDetail'))byId('managerDetail').innerHTML=deckManagerDetailHTML()
 const status=byId('managerDirtyStatus');if(status)status.textContent=deckManagerDirty()?'未保存の変更':'保存済み'
 deckManagerRemember();scheduleDeckManagerViewport()
}
async function deckManagerOpen(deck=null,{unsaved=false}={}){
 if(!(await confirmDeckManagerLeave()))return false
 discardDeckManagerDraft()
 const now=new Date().toISOString();deckManager.draft=deck?copy(deck):{id:uid('deck'),name:'新しいデッキ',list:[],channelDefinition:'',createdAt:now,updatedAt:now}
 deckManager.baseline=deck&&!unsaved?deckManagerSignature(deckManager.draft):'';deckManager.selected=deckManager.draft.channelDefinition||deckManager.draft.list[0]?.definition||'';deckManager.view='editor'
 Object.assign(deckManager,{query:'',type:'',className:'',sort:'registered',sortDirection:'asc',page:0,catalogMode:'main'});deckManager.colors.clear();renderDeckPage()
 return true
}
async function deckManagerLeaveToGallery(){
 if(!(await confirmDeckManagerLeave()))return false
 discardDeckManagerDraft();renderDeckPage();return true
}
function deckManagerSetQuantity(id,value,rerender=true){
 if(!deckManager.draft||!deckManagerMain(id))return false
 const count=deckManagerCount(value),list=deckManager.draft.list,old=list.find(entry=>entry.definition===id)
 if((old?.count||0)===count)return true
 if(count+deckManagerTotal(deckManager.draft)-(old?.count||0)>500){notify('デッキは500枚までです。',true);return false}
 if(old&&count)old.count=count;else if(old)list.splice(list.indexOf(old),1);else if(count)list.push({definition:id,count})
 deckManager.selected=id
 if(rerender)deckManagerRefreshParts();else{deckManagerRefreshParts({contents:false});const title=byId('managerContents')?.querySelector('.dm-main-section h3 span');if(title)title.textContent=`${deckManagerTotal(deckManager.draft)}枚`}
 return true
}
async function saveCurrentNamedDeck(){
 if(appPage!=='deck'||!storageReady||!deckManager.draft||deckManager.saving)return false
 const name=(byId('managerDeckName')?.value??deckManager.draft.name).trim()
 if(!name){notify('デッキ名を入力してください。',true);byId('managerDeckName')?.focus();return false}
 if(deckManagerTotal(deckManager.draft)>500){notify('デッキは500枚までです。',true);return false}
 if(deckManager.draft.list.some(item=>!deckManagerMain(item.definition))||(deckManager.draft.channelDefinition&&!deckManagerChannel(deckManager.draft.channelDefinition))){notify('登録が見つからないカード、または種別が変わったカードがあります。カード登録を確認してください。',true);return false}
 deckManager.draft.name=name;const now=new Date().toISOString(),deck={...copy(deckManager.draft),createdAt:deckManager.draft.createdAt||now,updatedAt:now}
 deckManager.saving=true;renderDeckPage()
 try{const result=await saveNamedDeck(deck);if(result===false)return false;const saved=result&&typeof result==='object'?result:deck;if(deckManager.draft?.id===deck.id){deckManager.draft=copy(saved);deckManager.baseline=deckManagerSignature(saved)}notify(`「${saved.name}」を保存しました。`);return true}
 catch(error){console.error('Deck save failed',error);notify(`デッキを保存できませんでした。${error.message}`,true);return false}
 finally{deckManager.saving=false;renderDeckPage()}
}
async function deckManagerExport(button){
 if(!deckManager.draft||deckManager.exporting)return
 const draft=copy(deckManager.draft),entries=draft.list.filter(entry=>entry.count>0&&deckManagerMain(entry.definition)).map(entry=>({...entry,item:state.catalog[entry.definition]})),channel=deckManagerChannel(draft.channelDefinition)?state.catalog[draft.channelDefinition]:null
 if(!entries.length&&!channel)return notify('画像にするカードを追加してください。',true)
 const epoch=++deckManager.exportEpoch;deckManager.exporting=true;if(button){button.disabled=true;button.textContent='画像を作成中…'}
 try{
  const canvas=await buildDeckExportCanvas('p1',entries,channel,draft.name),blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'))
  if(!blob)throw new Error('画像を作成できませんでした。')
  if(epoch!==deckManager.exportEpoch||deckManager.view!=='editor')return
  clearDeckExportPreview();deckExportPreviewURL=URL.createObjectURL(blob);const fileName=draft.name.replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_').slice(0,80)||'deck';deckExportFileName=`NijinanaSimulator_${fileName}.png`
  openModal('デッキ画像の確認',`<div class="deck-export-preview"><img src="${deckExportPreviewURL}" alt="${escapeHTML(draft.name)}のデッキ画像"></div><div class="modal-footer"><button data-action="deck-export-back">デッキ編集に戻る</button><button data-action="deck-export-save" class="primary">PNG画像を保存</button></div>`,'deck-export')
 }catch(error){console.error('Deck export failed',error);notify(`画像を出力できませんでした。${error.message}`,true)}
 finally{if(epoch!==deckManager.exportEpoch)return;deckManager.exporting=false;const currentButton=byId('deckPage')?.querySelector('[data-action="manager-export"]');if(currentButton){currentButton.disabled=false;currentButton.textContent='画像出力'}else if(button?.isConnected){button.disabled=false;button.textContent='画像出力'}}
}
function deckManagerPageURL(){return appDeckManagerURL()}
async function showSavedDeckSelector(player=state.active){
 deckManager.pickerPlayer=playerIds.includes(player)?player:'p1'
 const epoch=++deckManager.pickerEpoch;deckManager.pickerLoading=true;deckManager.pickerError=''
 openModal('デッキを選択','','deck-picker');renderSavedDeckSelector()
 try{
  if(!storageReady)throw new Error('保存領域の準備が終わってから開いてください。')
  await waitLocalDeckTransfer()
  await refreshDeckStorage()
 }catch(error){if(epoch===deckManager.pickerEpoch)deckManager.pickerError=error.message||'読み込みに失敗しました。'}
 finally{if(epoch===deckManager.pickerEpoch){deckManager.pickerLoading=false;if(modalKind==='deck-picker')renderSavedDeckSelector()}}
}
function renderSavedDeckSelector(){
 if(modalKind!=='deck-picker')return
 if(deckManager.pickerLoading||deckManager.pickerError){
  byId('modalContent').innerHTML=`<p class="dm-picker-description" role="status">${deckManager.pickerLoading?'保存したデッキを読み込み中…':`保存したデッキを読み込めませんでした。${escapeHTML(deckManager.pickerError)}`}</p>${deckManager.pickerError?'<button data-action="reload-saved-decks">再読み込み</button>':''}`
  return
 }
 const decks=deckManagerDecks(),player=deckManager.pickerPlayer
 const tiles=decks.map(deck=>`<button class="dm-picker-deck" data-action="choose-saved-deck" data-deck-id="${escapeHTML(deck.id)}" data-player="${player}">${deckManagerCoverHTML(deck)}<span><strong>${escapeHTML(deck.name)}</strong><small>${deckManagerTotal(deck)}枚 · ${state.catalog[deck.channelDefinition]?escapeHTML(state.catalog[deck.channelDefinition].name):'チャンネル未設定'}</small></span></button>`).join('')
 byId('modalContent').innerHTML=`<p class="dm-picker-description">${escapeHTML(state.players[player].name)}で使用するデッキを選んでください。</p><div class="dm-picker-grid">${tiles||'<p class="dm-picker-empty">保存したデッキがありません。「デッキ管理」で作成してください。</p>'}</div><div class="modal-footer"><a class="dm-link-button app-link" href="${deckManagerPageURL()}">デッキ管理を開く</a></div>`
}
function handleDeckManagerAction(action,button){
 if(action==='local-deck-transfer-retry'){initializeLocalDeckTransfer();return true}
 if(action==='reload-saved-decks'){showSavedDeckSelector(deckManager.pickerPlayer);return true}
 if(action.startsWith('manager-leave-')){
  if(action==='manager-leave-save')saveDeckManagerAndLeave()
  else if(action==='manager-leave-discard')finishDeckManagerLeave(true)
  else if(action==='manager-leave-stay')finishDeckManagerLeave(false)
  return true
 }
 if(action==='choose-saved-deck'){
  if(!storageReady||deckManager.pickerLoading||replay.mode==='replay')return true
  const deck=deckManagerDecks().find(item=>item.id===button.dataset.deckId);if(!deck){notify('このデッキが見つかりません。',true);renderSavedDeckSelector();return true}
  const player=playerIds.includes(button.dataset.player)?button.dataset.player:deckManager.pickerPlayer,list=(deck.list||[]).filter(item=>deckManagerMain(item.definition)),channel=deckManagerChannel(deck.channelDefinition)?deck.channelDefinition:''
  if(transact(`「${deck.name}」を${state.players[player].name}の山札に設定`,()=>createDeckRaw(player,list,channel))){closeModal();notify(`「${deck.name}」を設定しました。`)}return true
 }
 if(action==='catalog-return-decks'){if(appPage==='deck'){closeModal(false);renderDeckPage()}else goToDeckManager();return true}
 if(appPage!=='deck')return false
 if(action==='deck-export-back'){closeModal();return true}
 if(action==='deck-editor'){closeModal();renderDeckPage();return true}
 if(!action.startsWith('manager-'))return false
 if(!storageReady||deckManager.saving)return true
 if(action==='manager-new')deckManagerOpen()
 
 else if(action==='manager-list')deckManagerLeaveToGallery()
 else if(action==='manager-edit'){const deck=deckManagerDecks().find(item=>item.id===button.dataset.deckId);if(deck)deckManagerOpen(deck)}
 else if(action==='manager-duplicate'){
  const deck=deckManagerDecks().find(item=>item.id===button.dataset.deckId)
  if(deck){const now=new Date().toISOString(),duplicated={...copy(deck),id:uid('deck'),name:`${deck.name} のコピー`.slice(0,120),createdAt:now,updatedAt:now};Promise.resolve(saveNamedDeck(duplicated)).then(result=>{if(result!==false){notify(`「${result?.name||duplicated.name}」を作成しました。`);renderDeckPage()}}).catch(error=>notify(error.message,true))}
 }
 else if(action==='manager-delete'){
  const deck=deckManagerDecks().find(item=>item.id===button.dataset.deckId)
  if(deck&&confirm(`「${deck.name}」を削除しますか？`))Promise.resolve(deleteNamedDeck(deck.id)).then(result=>{if(result===false)return;if(deckManager.draft?.id===deck.id&&!deckManagerDirty()){deckManager.draft=null;try{sessionStorage.removeItem(deckManagerDraftKey)}catch{}}renderDeckPage()}).catch(error=>notify(error.message,true))
 }
 else if(action==='manager-save')saveCurrentNamedDeck()
 else if(action==='manager-export')deckManagerExport(button)
 else if(action==='manager-txt-export')deckManagerExportText()
 else if(action==='manager-txt-import'){if(!deckManager.textImporting){deckManager.exportEpoch++;byId('deckTxtPicker')?.click()}}
 else if(action==='manager-txt-apply')deckManagerApplyTextImport()
 else if(action==='manager-select'){deckManager.selected=button.dataset.definition;deckManagerRefreshParts()}
 else if(action==='manager-plus'||action==='manager-minus'){const id=button.dataset.definition,count=deckManager.draft.list.find(entry=>entry.definition===id)?.count||0;deckManagerSetQuantity(id,count+(action==='manager-plus'?1:-1))}
 else if(action==='manager-set-channel'){const id=button.dataset.definition;if(deckManagerChannel(id)){deckManager.draft.channelDefinition=id;deckManager.selected=id;deckManagerRefreshParts()}}
 else if(action==='manager-clear-channel'){deckManager.draft.channelDefinition='';deckManagerRefreshParts()}
 else if(action==='manager-main-mode'||action==='manager-channel-mode'){deckManager.catalogMode=action==='manager-channel-mode'?'channel':'main';deckManager.type='';deckManager.page=0;deckManager.colors.clear();renderDeckPage()}
 else if(action==='manager-page-prev'||action==='manager-page-next'){deckManager.page+=action==='manager-page-next'?1:-1;deckManagerRefreshParts({contents:false,detail:false})}
 else if(action==='manager-clear-filters'){Object.assign(deckManager,{query:'',type:'',className:'',sort:'registered',sortDirection:'asc',page:0});deckManager.colors.clear();renderDeckPage()}
 return true
}
function handleDeckManagerInput(target){
 if(appPage!=='deck'||!deckManager.draft)return false
 if(target.id==='managerDeckName'){deckManager.draft.name=target.value;deckManagerRefreshParts({catalog:false,contents:false,detail:false});return true}
 if(target.id==='managerSearch'){deckManager.query=target.value;deckManager.page=0;deckManagerRefreshParts({contents:false,detail:false});return true}
 if(target.dataset.managerQuantity){if(target.value.trim()&&deckManagerSetQuantity(target.dataset.managerQuantity,target.value,false)===false)target.value=deckManager.draft.list.find(entry=>entry.definition===target.dataset.managerQuantity)?.count||0;return true}
 return false
}
function handleDeckManagerChange(target){
 if(appPage!=='deck'||!deckManager.draft)return false
 const field={managerType:'type',managerClass:'className',managerSort:'sort',managerSortDirection:'sortDirection'}[target.id]
 if(field){
  deckManager[field]=field==='sort'?(Object.hasOwn(deckManagerSorts,target.value)?target.value:'registered'):field==='sortDirection'?(Object.hasOwn(deckManagerSortDirections,target.value)?target.value:'asc'):target.value
  deckManager.page=0;deckManagerRefreshParts({contents:false,detail:false});return true
 }
 if(target.dataset.managerColor){if(target.checked)deckManager.colors.add(target.dataset.managerColor);else deckManager.colors.delete(target.dataset.managerColor);deckManager.page=0;deckManagerRefreshParts({contents:false,detail:false});return true}
 if(target.dataset.managerQuantity){deckManagerSetQuantity(target.dataset.managerQuantity,target.value);deckManagerRefreshParts();return true}
 return false
}

// Catalog-to-deck adds one copy. Deck-to-catalog removes one copy. Deck-to-deck reorders entries.
let deckManagerDragDefinition = ''
let deckManagerDragOrigin = ''
const deckManagerDragType = 'application/x-nijinana-card-definition'
function deckManagerCanDrag() {
 return appPage === 'deck' && storageReady && deckManager.view === 'editor' && !!deckManager.draft && !deckManager.saving && !byId('modal')?.open
}
function clearDeckManagerDropFeedback() {
 for (const element of document.querySelectorAll('.dm-drop-active,.dm-drop-before,.dm-drop-after')) element.classList.remove('dm-drop-active','dm-drop-before','dm-drop-after')
}
function clearDeckManagerDrag() {
 deckManagerDragDefinition = ''
 deckManagerDragOrigin = ''
 clearDeckManagerDropFeedback()
 for (const element of document.querySelectorAll('.dm-drag-source')) element.classList.remove('dm-drag-source')
}
function handleDeckManagerDragStart(event) {
 const source = event.target.closest('[data-manager-drag-definition]')
 if (!source) return
 clearDeckManagerDrag()
 const id = source.dataset.managerDragDefinition, origin = source.dataset.managerDragOrigin || 'catalog'
 const included = origin === 'deck' ? deckManager.draft?.list.some(entry => entry.definition === id && entry.count > 0)
  : origin === 'channel' ? deckManager.draft?.channelDefinition === id : origin === 'catalog'
 if (!deckManagerCanDrag() || !event.dataTransfer || !included || (!deckManagerMain(id) && !deckManagerChannel(id))) {event.preventDefault(); return}
 deckManagerDragDefinition = id
 deckManagerDragOrigin = origin
 event.dataTransfer.setData(deckManagerDragType, JSON.stringify({definition:id,origin}))
 event.dataTransfer.setData('text/plain', state.catalog[id].name)
 event.dataTransfer.effectAllowed = origin === 'catalog' ? 'copy' : 'move'
 const sourceTile = origin === 'catalog' ? source.closest('.dm-catalog-card') : origin === 'deck' ? source.closest('.dm-included-card') : source
 sourceTile?.classList.add('dm-drag-source')
}
function deckManagerDropTarget(event) {
 if (!deckManagerCanDrag() || !deckManagerDragDefinition) return null
 const target = event.target.closest('[data-manager-drop]')
 if (!target || !byId('deckPage')?.contains(target)) return null
 const id = deckManagerDragDefinition, destination = target.dataset.managerDrop
 if (deckManagerDragOrigin !== 'catalog') {
  const included = deckManagerDragOrigin === 'deck' ? deckManagerMain(id) && deckManager.draft.list.some(entry => entry.definition === id && entry.count > 0)
   : deckManagerDragOrigin === 'channel' && deckManagerChannel(id) && deckManager.draft.channelDefinition === id
  const allowed = destination === 'catalog' || (deckManagerDragOrigin === 'deck' && destination === 'main')
  return included && allowed ? target : null
 }
 if (destination === 'main' ? !deckManagerMain(id) : destination !== 'channel' || !deckManagerChannel(id)) return null
 return target
}
function deckManagerReorderPosition(target, event) {
 if (!Number.isFinite(event.clientX) || !Number.isFinite(event.clientY)) return null
 const tiles = [...target.querySelectorAll('[data-manager-entry]')].map(tile => ({tile,rect:tile.getBoundingClientRect()})).filter(({rect}) => rect.width > 0 && rect.height > 0)
 if (!tiles.length) return null
 const rows = []
 for (const entry of tiles) {
  let row = rows[rows.length - 1]
  if (!row || Math.abs(entry.rect.top - row.top) > 2) {row = {top:entry.rect.top,bottom:entry.rect.bottom,tiles:[]};rows.push(row)}
  row.tiles.push(entry);row.bottom = Math.max(row.bottom,entry.rect.bottom)
 }
 const position = (entry, after) => {
  const index = deckManager.draft.list.findIndex(item => item.definition === entry.tile.dataset.managerEntry)
  return index < 0 ? null : {index:index + Number(after),tile:entry.tile,after}
 }
 if (event.clientY < rows[0].top) return position(tiles[0],false)
 const lastRow = rows[rows.length - 1]
 if (event.clientY > lastRow.bottom) return position(tiles[tiles.length - 1],true)
 let row = lastRow
 for (let index = 0; index < rows.length - 1; index++) {
  if (event.clientY < (rows[index].bottom + rows[index + 1].top) / 2) {row = rows[index];break}
 }
 const before = row.tiles.find(({rect}) => event.clientX < (rect.left + rect.right) / 2)
 return before ? position(before,false) : position(row.tiles[row.tiles.length - 1],true)
}
function deckManagerReorder(id, insertionIndex) {
 if (!deckManagerCanDrag() || !deckManagerMain(id) || !Number.isInteger(insertionIndex)) return false
 const list = deckManager.draft.list, from = list.findIndex(entry => entry.definition === id)
 if (from < 0 || insertionIndex < 0 || insertionIndex > list.length) return false
 const to = insertionIndex > from ? insertionIndex - 1 : insertionIndex
 if (to === from) return true
 const [entry] = list.splice(from,1)
 list.splice(to,0,entry)
 deckManagerRefreshParts({catalog:false,detail:false})
 return true
}
function handleDeckManagerDragOver(event) {
 const target = deckManagerDropTarget(event)
 clearDeckManagerDropFeedback()
 if (!target) {if (deckManagerDragDefinition && event.dataTransfer) event.dataTransfer.dropEffect = 'none'; return}
 event.preventDefault()
 if (event.dataTransfer) event.dataTransfer.dropEffect = deckManagerDragOrigin === 'catalog' ? 'copy' : 'move'
 target.classList.add('dm-drop-active')
 if (deckManagerDragOrigin === 'deck' && target.dataset.managerDrop === 'main') {
  const position = deckManagerReorderPosition(target,event)
  if (position) position.tile.classList.add(position.after ? 'dm-drop-after' : 'dm-drop-before')
 }
}
function handleDeckManagerDrop(event) {
 const target = deckManagerDropTarget(event), id = deckManagerDragDefinition, origin = deckManagerDragOrigin
 if (!target) {clearDeckManagerDrag(); return false}
 event.preventDefault()
 const destination = target.dataset.managerDrop
 const position = origin === 'deck' && destination === 'main' ? deckManagerReorderPosition(target,event) : null
 const transferred = event.dataTransfer?.getData(deckManagerDragType)
 clearDeckManagerDrag()
 if (transferred !== JSON.stringify({definition:id,origin})) return false
 if (origin === 'deck' && destination === 'main') return position ? deckManagerReorder(id,position.index) : false
 if (destination === 'channel' || (destination === 'catalog' && origin === 'channel')) {
  deckManager.draft.channelDefinition = destination === 'channel' ? id : ''
  deckManager.selected = id
  deckManagerRefreshParts()
  return true
 }
 const count = deckManager.draft.list.find(entry => entry.definition === id)?.count || 0
 return deckManagerSetQuantity(id, count + (destination === 'catalog' ? -1 : 1))
}
document.addEventListener('dragstart', handleDeckManagerDragStart)
document.addEventListener('dragenter', handleDeckManagerDragOver)
document.addEventListener('dragover', handleDeckManagerDragOver)
document.addEventListener('dragleave', event => {
 const target = event.target.closest('[data-manager-drop]')
 if (target && !target.contains(event.relatedTarget)) clearDeckManagerDropFeedback()
})
document.addEventListener('drop', handleDeckManagerDrop)
document.addEventListener('dragend', clearDeckManagerDrag)

let deckManagerViewportFrame = null
function updateDeckManagerViewport() {
 if (appPage !== 'deck' || deckManager.view !== 'editor') return
 const root = byId('deckPage'), browser = root?.querySelector('.dm-browser')
 if (!browser || !root.style) return
 const top = browser.getBoundingClientRect().top
 if (Number.isFinite(top)) root.style.setProperty('--dm-editor-top', `${Math.max(12, top)}px`)
}
function scheduleDeckManagerViewport() {
 if (deckManagerViewportFrame !== null) return
 if (typeof globalThis.requestAnimationFrame !== 'function') {updateDeckManagerViewport(); return}
 deckManagerViewportFrame = globalThis.requestAnimationFrame(() => {
  deckManagerViewportFrame = null
  updateDeckManagerViewport()
 })
}
globalThis.addEventListener?.('resize', scheduleDeckManagerViewport)
globalThis.addEventListener?.('scroll', scheduleDeckManagerViewport, {passive:true})

let deckManagerTextImport = null
let deckManagerTextReadEpoch = 0
function deckManagerExportText() {
 if (appPage !== 'deck' || !storageReady || !deckManager.draft || deckManager.saving) return false
 let url = ''
 try {
  const draft = {...deckManager.draft,name:(byId('managerDeckName')?.value ?? deckManager.draft.name).trim() || '新しいデッキ'}
  const content = buildDeckText(draft,state.catalog)
  const blob = new Blob(['\uFEFF',content],{type:'text/plain;charset=utf-8'})
  url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = `NijinanaSimulator_${draft.name.replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_').slice(0,80)||'deck'}.txt`
  document.body.appendChild(anchor)
  try {anchor.click()} finally {anchor.remove()}
  notify('TXTを出力します。')
  return true
 } catch (error) {notify(`TXTを出力できませんでした。${error.message}`,true);return false}
 finally {if(url)setTimeout(()=>URL.revokeObjectURL(url),10000)}
}
function deckManagerTextImportPreview(parsed) {
 const total = parsed.list.reduce((sum,entry)=>sum+entry.count,0),channel = state.catalog[parsed.channelDefinition]
 const issues = parsed.issues.length ? `<div class="dm-text-errors" role="alert"><p>次のカードを照合できませんでした。カード登録を確認してから、TXTを読み込み直してください。</p><ul>${parsed.issues.map(issue=>`<li>${issue.line?`${issue.line}行目：`:''}${escapeHTML(issue.message)}</li>`).join('')}</ul></div>` : ''
 const rows = parsed.list.map(entry=>{const item=state.catalog[entry.definition];return `<tr><td>${escapeHTML(item?.sourceCardKey||'—')}</td><td>${escapeHTML(item?.name||'')}</td><td>${entry.count}</td></tr>`}).join('')
 openModal('TXTデッキの確認',`<div class="dm-text-preview"><h3>${escapeHTML(parsed.name)}</h3>${issues}<div class="dm-text-channel"><span>チャンネル</span><strong>${channel?escapeHTML(channel.name):'未設定'}</strong>${channel?.sourceCardKey?`<small>ID: ${escapeHTML(channel.sourceCardKey)}</small>`:''}</div><p class="dm-text-summary">${parsed.issues.length?'照合できたデッキ':'デッキ'}：${total}枚 · ${parsed.list.length}種類</p><div class="dm-text-list"><table><thead><tr><th>ID</th><th>カード名</th><th>枚数</th></tr></thead><tbody>${rows||'<tr><td colspan="3">カードがありません。</td></tr>'}</tbody></table></div><p class="dm-text-note">新しいデッキとして編集画面に読み込みます。「デッキを保存」で保存できます。</p></div><div class="modal-footer"><button data-action="manager-txt-apply" class="primary" ${parsed.issues.length?'disabled':''}>読み込む</button><button data-action="close-modal">キャンセル</button></div>`,'deck-txt-import')
}
function deckManagerTextImportButtons() {
 for (const button of document.querySelectorAll('[data-action="manager-txt-import"]')) {button.disabled=deckManager.textImporting||deckManager.saving;button.textContent=deckManager.textImporting?'読み込み中…':'TXT読込'}
}
async function deckManagerReadText(file) {
 if (!file || appPage !== 'deck' || !storageReady || deckManager.saving) return false
 const epoch = ++deckManagerTextReadEpoch
 const readContext = JSON.stringify([deckManager.view,deckManager.draft?.id||'',modalKind,deckManager.exportEpoch])
 deckManager.textImporting=true;deckManagerTextImport=null;deckManagerTextImportButtons()
 try {
  if (file.size>1024*1024) throw new Error('TXTは1MB以下にしてください。')
  const text = await file.text()
  if (epoch!==deckManagerTextReadEpoch || !storageReady || deckManager.saving || readContext!==JSON.stringify([deckManager.view,deckManager.draft?.id||'',modalKind,deckManager.exportEpoch])) return false
  const parsed = parseDeckText(text,state.catalog)
  deckManagerTextImport={text,parsed}
  deckManagerTextImportPreview(parsed)
  return true
 } catch(error) {if(epoch===deckManagerTextReadEpoch)notify(`TXTを読み込めませんでした。${error.message}`,true);return false}
 finally {if(epoch===deckManagerTextReadEpoch){deckManager.textImporting=false;deckManagerTextImportButtons()}}
}
async function deckManagerApplyTextImport() {
 if (appPage!=='deck' || !storageReady || deckManager.saving || modalKind!=='deck-txt-import' || !deckManagerTextImport) return false
 const importText=deckManagerTextImport.text
 if(!(await confirmDeckManagerLeave()))return false
 try {
  const parsed = parseDeckText(importText,state.catalog)
  deckManagerTextImport={text:importText,parsed}
  if (parsed.issues.length) {deckManagerTextImportPreview(parsed);return false}
  const now = new Date().toISOString(),draft={id:uid('deck'),name:parsed.name,list:copy(parsed.list),channelDefinition:parsed.channelDefinition,createdAt:now,updatedAt:now}
  if(!(await deckManagerOpen(draft,{unsaved:true})))return false
  if (deckManager.draft?.id!==draft.id) return false
  deckManagerTextImport=null;closeModal(false)
  notify('TXTを読み込みました。内容を確認して「デッキを保存」を押してください。')
  return true
 } catch(error) {notify(`TXTを読み込めませんでした。${error.message}`,true);return false}
}
byId('deckTxtPicker')?.addEventListener('change',async event=>{
 const file=event.target.files[0];event.target.value='';await deckManagerReadText(file)
})

let deckManagerPendingLeave=null
let deckManagerUnloadGuardActive=false
function deckManagerBeforeUnload(event){
 if(appPage!=='deck'||(!deckManagerDirty()&&!deckManager.saving))return
 event.preventDefault();event.returnValue=''
}
function discardDeckManagerDraft(){
 deckManager.exportEpoch++;deckManager.exporting=false;deckManagerTextReadEpoch++;deckManagerTextImport=null;deckManager.textImporting=false
 deckManager.draft=null;deckManager.baseline='';deckManager.selected='';deckManager.view='gallery'
 clearDeckManagerDrag();deckManagerRemember()
}
function confirmDeckManagerLeave(){
 if(appPage!=='deck')return Promise.resolve(true)
 if(deckManager.saving||deckManagerPendingLeave)return Promise.resolve(false)
 if(!deckManagerDirty())return Promise.resolve(true)
 deckManager.exportEpoch++;deckManager.exporting=false;deckManagerTextReadEpoch++;deckManager.textImporting=false
 if(modalKind==='deck-export')clearDeckExportPreview()
 return new Promise(resolve=>{
  deckManagerPendingLeave={resolve,busy:false}
  openModal('未保存の変更',`<div class="dm-leave-confirm"><p>「${escapeHTML(deckManager.draft.name||'新しいデッキ')}」の変更を保存してから離れますか？</p><p id="deckLeaveError" class="dm-leave-error" role="alert"></p></div><div class="modal-footer"><button data-action="manager-leave-save" class="primary">保存して離れる</button><button data-action="manager-leave-discard">破棄して離れる</button><button data-action="manager-leave-stay">編集に戻る</button></div>`,'deck-leave')
 })
}
function deckManagerLeaveBusy(busy){
 for(const button of document.querySelectorAll('#modal [data-action^="manager-leave-"],#modal .modal-heading [data-action="close-modal"]'))button.disabled=busy
}
function finishDeckManagerLeave(leave){
 const pending=deckManagerPendingLeave
 if(!pending||pending.busy)return false
 deckManagerPendingLeave=null;deckManagerLeaveBusy(false)
 if(leave)discardDeckManagerDraft()
 closeModal(false);pending.resolve(leave);return true
}
async function saveDeckManagerAndLeave(){
 const pending=deckManagerPendingLeave
 if(!pending||pending.busy)return false
 pending.busy=true;deckManagerLeaveBusy(true)
 try{
  const saved=await saveCurrentNamedDeck()
  if(deckManagerPendingLeave!==pending)return false
  pending.busy=false;deckManagerLeaveBusy(false)
  if(saved)return finishDeckManagerLeave(true)
  const error=byId('deckLeaveError');if(error)error.textContent='保存できませんでした。「編集に戻る」から内容を確認してください。'
  return false
 }catch(error){
  if(deckManagerPendingLeave===pending){pending.busy=false;deckManagerLeaveBusy(false);const message=byId('deckLeaveError');if(message)message.textContent=`保存できませんでした。${error.message}`}
  return false
 }
}
async function handleDeckManagerNavigation(event){
 if(appPage!=='deck')return
 const link=event.target.closest('a[href]')
 if(!link||event.defaultPrevented||event.button!==0||event.ctrlKey||event.metaKey||event.shiftKey||event.altKey||link.hasAttribute('download')||(link.target&&link.target!=='_self'))return
 const href=link.getAttribute('href');if(!href||href.startsWith('#'))return
 const target=new URL(link.href||href,location.href)
 if(!['http:','https:','file:'].includes(target.protocol))return
 event.preventDefault();event.stopImmediatePropagation()
 if(!(await confirmDeckManagerLeave()))return
 try{
  if(storageReady)await flushStorageSave()
  discardDeckManagerDraft();window.location.assign(target.href)
 }catch(error){notify(`移動できませんでした。${error.message}`,true)}
}
document.addEventListener('click',handleDeckManagerNavigation,true)
byId('modal')?.addEventListener('cancel',event=>{
 if(!deckManagerPendingLeave)return
 event.preventDefault();finishDeckManagerLeave(false)
})
globalThis.addEventListener?.('pagehide',()=>{
 if(appPage!=='deck')return
 const pending=deckManagerPendingLeave;deckManagerPendingLeave=null
 if(pending)pending.resolve(false)
 discardDeckManagerDraft()
})
globalThis.addEventListener?.('pageshow',event=>{
 if(appPage!=='deck'||!event.persisted)return
 if(byId('modal')?.open)closeModal(false)
 else renderDeckPage()
})
