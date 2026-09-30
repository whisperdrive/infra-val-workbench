// Chart.js set-up shared by the Model Desk (app/) and the Valuation Desk (engage/): every kind of chart the chart
// tools can ask for, drawn from a spec's shown data (labels and series, sign presentation already applied).
//   line       one line per series
//   bar        clustered columns
//   stacked    columns stacked into a total; a series marked as "line" stays a line on top (e.g. last year's total)
//   area       stacked areas
//   combo      stacked columns for the parts, lines for the series marked as "line"
//   waterfall  one series as a bridge: each value moves the running total; spec.totals are the bars drawn from
//              zero (the start, subtotals, the end); without them a Total bar is added at the end
// A series with axis "right" gets its own axis on the right (a rate beside amounts), never stacked.
const KIND_WORD = {line: "Line", bar: "Bar", stacked: "Stacked", area: "Area", combo: "Combo", waterfall: "Waterfall"};

// The kinds worth offering for a spec, in button order.
function chartKinds(spec) {
  if (spec.kind === "waterfall") return ["waterfall", "bar"];
  const k = ["line", "bar"];
  if (spec.series.length > 1) k.push("stacked", "area");
  if (spec.kind === "combo" || spec.series.some(s => s.as === "line")) k.push("combo");
  return k;
}

// o: {color(i), ink, muted, rule, shade, up, down, total, band, font, fmt(v), unit, plugins, onClick}
// spec.highlight: the categories to mark (their indices), a band behind each: what a review point is about
function chartConfig(spec, d, kind, o) {
  const marks = (spec.highlight || []).filter(i => i >= 0 && i < d.labels.length);
  if (marks.length) o = {...o, plugins: [...(o.plugins || []), {id: "reviewBand", beforeDatasetsDraw(chart) {
    const x = chart.scales.x, a = chart.chartArea, ctx = chart.ctx, n = chart.data.labels.length;
    const w = n > 1 ? Math.abs(x.getPixelForValue(1) - x.getPixelForValue(0)) : a.right - a.left;
    ctx.save(); ctx.fillStyle = o.band || "rgba(224, 180, 0, .2)";
    marks.forEach(i => { const c = x.getPixelForValue(i); ctx.fillRect(c - w / 2, a.top, w, a.bottom - a.top); });
    // the marked categories named at the top of the band (an axis may skip their own labels)
    const name = i => [].concat(chart.data.labels[i]).join(" "), first = Math.min(...marks), last = Math.max(...marks);
    ctx.fillStyle = o.ink; ctx.font = `600 11px ${o.font?.family || "sans-serif"}`; ctx.textAlign = "center"; ctx.textBaseline = "top";
    ctx.fillText(first === last ? name(first) : `${name(first)}–${name(last)}`, (x.getPixelForValue(first) + x.getPixelForValue(last)) / 2, a.top + 2);
    ctx.restore(); }}]};
  const n = d.labels.length, num = v => Number(v).toLocaleString();
  const axis = (extra = {}) => ({ticks: {color: o.muted, callback: num}, grid: {color: o.rule}, border: {display: false}, ...extra});
  const x = {ticks: {color: o.muted, maxRotation: 0, autoSkip: true, maxTicksLimit: 10}, grid: {display: false}, border: {color: o.rule}};
  const base = (datasets, labels, scales, extra = {}) => ({
    type: extra.type || "bar", plugins: o.plugins || [], data: {labels, datasets},
    options: {responsive: true, maintainAspectRatio: false, animation: false, interaction: {mode: "index", intersect: false},
      plugins: {phases: {phases: kind === "waterfall" ? null : d.phases, shade: o.shade, text: o.muted},
        legend: {display: extra.legend ?? d.series.length > 1, labels: {color: o.ink, boxWidth: 12, font: o.font}},
        tooltip: {callbacks: {label: extra.tip || (c => ` ${c.dataset.label}: ${o.fmt(c.raw)}${o.unit}`)}}},
      scales, onClick: o.onClick}});

  if (kind === "waterfall") {
    const s = d.series[0] || {name: "", data: []}, tot = new Set(spec.totals || []);
    const bars = [], colours = [], steps = [];
    let run = 0;
    s.data.forEach((v, i) => {
      v = v ?? 0;
      if (tot.has(i)) { bars.push([0, v]); colours.push(o.total); run = v; }
      else { bars.push([run, run + v]); colours.push(v >= 0 ? o.up : o.down); run += v; }
      steps.push(v);
    });
    const labels = [...d.labels];
    if (!tot.size) { bars.push([0, run]); colours.push(o.total); steps.push(run); labels.push("Total"); tot.add(labels.length - 1); }
    // every bar named (a long name on two lines) and its value written above it: a bridge is read step by step
    const wrap = l => { if (typeof l !== "string" || l.length <= 10 || !l.includes(" ")) return l;
      const w = l.split(" "), half = Math.ceil(w.length / 2); return [w.slice(0, half).join(" "), w.slice(half).join(" ")]; };
    const values = {id: "wfValues", afterDatasetsDraw(chart) {
      const ctx = chart.ctx, meta = chart.getDatasetMeta(0);
      ctx.save(); ctx.font = `600 11.5px ${o.font?.family || "sans-serif"}`; ctx.fillStyle = o.ink; ctx.textAlign = "center"; ctx.textBaseline = "bottom";
      meta.data.forEach((bar, i) => { const v = steps[i]; if (v == null) return;
        ctx.fillText(tot.has(i) ? o.fmt(v) : (v > 0 ? "+" : "") + o.fmt(v), bar.x, Math.min(bar.y, bar.base) - 4); });
      ctx.restore(); }};
    const cfg = base([{label: s.name, data: bars, backgroundColor: colours, borderWidth: 0, steps}], labels.map(wrap),
      {x: {...x, ticks: {...x.ticks, autoSkip: false, maxRotation: 0}}, y: axis({beginAtZero: true, grace: "10%"})},
      {legend: false, tip: c => ` ${[].concat(c.label).join(" ")}: ${o.fmt(c.dataset.steps[c.dataIndex])}${o.unit}`});
    cfg.plugins = [...cfg.plugins, values];
    return cfg;
  }

  const stacking = ["stacked", "area", "combo"].includes(kind);
  let firstArea = null;
  const datasets = d.series.map((s, i) => {
    const c = o.color(i), right = s.axis === "right";  // a series on the right axis is a line unless it says bar
    const asLine = kind === "line" || kind === "area" || (right && s.as !== "bar")
      || ((kind === "combo" || kind === "stacked") && s.as === "line");
    const ds = {label: right ? `${s.name} (right axis)` : s.name, data: s.data, borderColor: c, borderWidth: 2, spanGaps: true,
      yAxisID: right ? "y2" : "y"};
    if (asLine) {
      const area = kind === "area" && !right && s.as !== "line";  // a line on an area chart stays a line
      if (area && firstArea === null) firstArea = i;
      Object.assign(ds, {type: "line", backgroundColor: area ? c + "66" : c + "22", pointRadius: n > 60 || area ? 0 : 2.5,
        pointHoverRadius: 4, tension: 0, fill: area ? (i === firstArea ? "origin" : "-1") : false, order: 0,
        stack: area ? "area" : `line${i}`});  // its own stack: a line on a stacked chart isn't added to the columns
    } else {
      Object.assign(ds, {type: "bar", backgroundColor: c, order: 1, ...(stacking && !right ? {stack: "parts"} : {})});
    }
    return ds;
  });
  const scales = {x: {...x, stacked: stacking}, y: axis({stacked: stacking})};
  if (d.series.some(s => s.axis === "right")) scales.y2 = axis({position: "right", grid: {display: false}});
  return base(datasets, d.labels, scales, {type: kind === "line" || kind === "area" ? "line" : "bar"});
}

// The subtitle's count: "48 quarterly periods", "4 steps" for a bridge, "12 points" when there's no frequency.
function chartCount(spec, n) {
  if (spec.kind === "waterfall") return `${n} step${n === 1 ? "" : "s"}`;
  const f = {monthly: "monthly", quarterly: "quarterly", "semi-annual": "half-yearly", annual: "annual"}[spec.frequency];
  return f ? `${n} ${f} periods` : `${n} points`;
}
