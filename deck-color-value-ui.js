// Deck-editor presentation for the expected calls saved by revealing a color.
function deckManagerColorValueHTML() {
  const entries = [];
  for (const entry of deckManager.draft?.list || []) {
    if (!deckManagerMain(entry.definition)) continue;
    const item = state.catalog[entry.definition];
    const count = deckManagerCount(entry.count);
    if (!count || textCardType(item) !== 'liver') continue;
    entries.push({ count, colors: liverCardColors(item) });
  }

  const stats = NijinanaColorValue.analyze(entries);
  if (!stats.liverCount) {
    return '<p class="dm-color-value-empty">ライバーを追加してください。</p>';
  }
  if (stats.status === 'incomplete') {
    const missing = stats.missingColors.map(color => deckManagerColorNames[color] || color).join('・');
    return `<p class="dm-color-value-empty">${escapeHTML(missing)}の対応ライバーがありません。</p>`;
  }

  const points = stats.points;
  const width = 440;
  const left = 39;
  const right = 424;
  const top = 12;
  const bottom = 164;
  const x = remaining => left + (7 - remaining) * (right - left) / 6;
  const y = value => bottom - value * (bottom - top);
  const formatValue = value => Number(value).toFixed(3);
  const colorNames = colors => nijiColors
    .filter(color => colors.includes(color))
    .map(color => deckManagerColorNames[color] || color)
    .join('・');
  const series = [
    { key: 'maximum', label: '最高', marker: 'square' },
    { key: 'minimum', label: '最低', marker: 'diamond' },
    { key: 'average', label: '平均', marker: 'circle' },
  ];

  const grid = [0, 0.25, 0.5, 0.75, 1].map(value => {
    const label = value === 0 || value === 1 ? String(value) : String(value).replace(/^0/, '');
    return `<line class="dm-color-value-grid" data-color-value-grid="${value}" x1="${left}" x2="${right}" y1="${y(value)}" y2="${y(value)}"/><text class="dm-color-value-tick" data-color-value-y="${value}" x="${left - 9}" y="${y(value) + 5}" text-anchor="end">${label}</text>`;
  }).join('');
  const ticks = points.map(point => `<text class="dm-color-value-tick" data-color-value-x="${point.remaining}" x="${x(point.remaining)}" y="187" text-anchor="middle">${point.remaining}</text>`).join('');
  const lines = series.map(item => {
    const coordinates = points.map(point => `${x(point.remaining)},${y(point[item.key])}`).join(' ');
    const values = points.map(point => `${point.remaining}:${point[item.key]}`).join(',');
    return `<polyline class="dm-color-value-line dm-color-value-${item.key}" data-color-value-points="${values}" points="${coordinates}"/>`;
  }).join('');
  const markers = series.map(item => points.map(point => {
    const centerX = x(point.remaining);
    const centerY = y(point[item.key]);
    const detail = item.key === 'average'
      ? `${point.patterns}通りの平均`
      : `残る裏の色：${colorNames(point[`${item.key}Colors`])}`;
    const calls = `裏だった色の回収：${formatValue(point[`${item.key}Before`])}回→${formatValue(point[`${item.key}After`])}回`;
    const title = `<title>${escapeHTML(`裏${point.remaining}枚・${item.label} ${formatValue(point[item.key])}点／${detail}／${calls}`)}</title>`;
    const className = `dm-color-value-marker dm-color-value-${item.key}`;
    const values = `data-color-value-remaining="${point.remaining}" data-color-value-value="${point[item.key]}"`;
    if (item.marker === 'square') {
      return `<rect class="${className}" ${values} x="${centerX - 4.5}" y="${centerY - 4.5}" width="9" height="9">${title}</rect>`;
    }
    if (item.marker === 'diamond') {
      return `<path class="${className}" ${values} d="M${centerX},${centerY - 4} l4,4 -4,4 -4,-4 Z">${title}</path>`;
    }
    return `<circle class="${className}" ${values} cx="${centerX}" cy="${centerY}" r="2.4">${title}</circle>`;
  }).join('')).join('');
  const totals = [
    { key: 'minimum', subscript: 'min', label: '最低線' },
    { key: 'average', subscript: '', label: '平均線' },
    { key: 'maximum', subscript: 'max', label: '最高線' },
  ].map(item => {
    const total = 7 + points.reduce((sum, point) => sum + point[item.key], 0);
    const symbol = `<span class="dm-color-value-total-symbol"><var>L</var>${item.subscript ? `<sub>${item.subscript}</sub>` : ''}</span>`;
    const title = `L${item.subscript ? `_${item.subscript}` : ''} = 7＋${item.label}の合計（残り7～1枚）`;
    return `<span class="dm-color-value-total-item" data-color-value-total="${item.key}" title="${escapeHTML(title)}">${symbol}<span>= ${formatValue(total)}</span></span>`;
  }).join('');
  return `<div class="dm-color-value-plot"><svg class="dm-color-value-chart" viewBox="0 0 ${width} 200" role="img" aria-label="裏枚数7枚から1枚までの、裏を表にする価値の最高・最低・平均。各点にカーソルを合わせると数値を確認できます。">
      <title>裏の色を表にする価値の推移</title>
      <desc>縦軸は表のカードをにじエリアに送る価値を1点とした評価値、横軸は残りの裏枚数です。残り1枚ではすべて0点になります。</desc>
      ${grid}<line class="dm-color-value-axis" x1="${left}" x2="${right}" y1="${bottom}" y2="${bottom}"/>
      ${lines}${markers}${ticks}
    </svg></div><div class="dm-color-value-total">${totals}</div>`;
}

// Keep one observer and one current plot across full renders and quantity edits.
const deckManagerColorValuePlotLayout = (() => {
  let currentPlot = null;
  let observer = null;
  let frame = null;
  let lastSVG = null;
  let lastSize = '';

  function resizePlot() {
    if (!currentPlot || typeof currentPlot.getBoundingClientRect !== 'function') return;
    const svg = currentPlot.querySelector('.dm-color-value-chart');
    if (!svg) return;
    const bounds = currentPlot.getBoundingClientRect();
    if (!(bounds.width > 0 && bounds.height > 0)) return;
    const size = `${bounds.width}:${bounds.height}`;
    if (svg === lastSVG && size === lastSize) return;

    const width = 440;
    const height = bounds.height * width / bounds.width;
    const left = 39;
    const right = 424;
    const top = 12;
    const bottom = Math.max(top + 1, height - 36);
    const x = remaining => left + (7 - remaining) * (right - left) / 6;
    const y = value => bottom - value * (bottom - top);
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);

    for (const line of svg.querySelectorAll('[data-color-value-grid]')) {
      const position = y(Number(line.getAttribute('data-color-value-grid')));
      line.setAttribute('y1', position);
      line.setAttribute('y2', position);
    }
    for (const tick of svg.querySelectorAll('[data-color-value-y]')) {
      tick.setAttribute('y', y(Number(tick.getAttribute('data-color-value-y'))) + 5);
    }
    for (const tick of svg.querySelectorAll('[data-color-value-x]')) {
      tick.setAttribute('y', bottom + 23);
    }
    const axis = svg.querySelector('.dm-color-value-axis');
    if (axis) {
      axis.setAttribute('y1', bottom);
      axis.setAttribute('y2', bottom);
    }
    for (const line of svg.querySelectorAll('[data-color-value-points]')) {
      const points = line.getAttribute('data-color-value-points').split(',').map(pair => {
        const [remaining, value] = pair.split(':').map(Number);
        return `${x(remaining)},${y(value)}`;
      });
      line.setAttribute('points', points.join(' '));
    }
    for (const marker of svg.querySelectorAll('[data-color-value-value]')) {
      const centerX = x(Number(marker.getAttribute('data-color-value-remaining')));
      const centerY = y(Number(marker.getAttribute('data-color-value-value')));
      const tag = marker.tagName.toLowerCase();
      if (tag === 'rect') marker.setAttribute('y', centerY - 4.5);
      else if (tag === 'circle') marker.setAttribute('cy', centerY);
      else if (tag === 'path') marker.setAttribute('d', `M${centerX},${centerY - 4} l4,4 -4,4 -4,-4 Z`);
    }
    lastSVG = svg;
    lastSize = size;
  }

  function scheduleResize() {
    if (frame !== null) return;
    if (typeof globalThis.requestAnimationFrame !== 'function') {
      resizePlot();
      return;
    }
    frame = globalThis.requestAnimationFrame(() => {
      frame = null;
      resizePlot();
    });
  }

  function initialize(curveRoot) {
    const plot = curveRoot && typeof curveRoot.querySelector === 'function'
      ? curveRoot.querySelector('.dm-color-value-plot') : null;
    if (plot !== currentPlot) {
      if (observer && currentPlot) observer.unobserve(currentPlot);
      currentPlot = plot;
      lastSVG = null;
      lastSize = '';
      if (currentPlot && typeof globalThis.ResizeObserver === 'function') {
        if (!observer) observer = new globalThis.ResizeObserver(entries => {
          if (entries.some(entry => entry.target === currentPlot)) scheduleResize();
        });
        observer.observe(currentPlot);
      }
    }
    if (!currentPlot) {
      if (frame !== null && typeof globalThis.cancelAnimationFrame === 'function') {
        globalThis.cancelAnimationFrame(frame);
      }
      frame = null;
      return;
    }
    resizePlot();
  }

  return { initialize };
})();

function deckManagerInitializeColorValuePlot(curveRoot) {
  const root = curveRoot || (typeof byId === 'function' ? byId('managerColorValueCurve') : null);
  deckManagerColorValuePlotLayout.initialize(root);
}
