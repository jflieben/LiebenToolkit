(() => {
  // Dependency-free SVG chart components for the dashboard and the native HTML report.
  // Colors are referenced through CSS custom properties (--viz-*) defined in style.css for
  // both themes, so charts restyle with the theme toggle and inline correctly when the
  // dashboard markup is serialized into the standalone report export.
  //
  // Every mark carries a visible text label (palette relief rule) and identity is never
  // color-alone: legends and labels accompany all multi-series visuals.

  function esc(value) {
    return `${value ?? ''}`
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function fmt(n) {
    const v = Number(n) || 0;
    return v.toLocaleString();
  }

  // Fixed categorical assignment (never cycled) and reserved status roles.
  const SERIES = ['var(--viz-s1)', 'var(--viz-s2)', 'var(--viz-s3)', 'var(--viz-s4)', 'var(--viz-s5)', 'var(--viz-s6)', 'var(--viz-s7)', 'var(--viz-s8)'];
  const STATUS_COLOR = {
    Passed: 'var(--viz-good)',
    Failed: 'var(--viz-critical)',
    Error: 'var(--viz-serious)',
    Investigate: 'var(--viz-warning)',
    Skipped: 'var(--viz-neutral)',
    Planned: 'var(--viz-neutral)',
  };

  // Semantic node colors for the Zero Trust sankeys (labels carry the identity; color reinforces).
  const SANKEY_NODE_COLORS = {
    'MFA': 'var(--viz-good)',
    'Compliant': 'var(--viz-good)',
    'Phish resistant': 'var(--viz-good)',
    'Passkey': 'var(--viz-good)',
    'WHfB': 'var(--viz-good)',
    'No MFA': 'var(--viz-critical)',
    'Non-compliant': 'var(--viz-critical)',
    'Single factor': 'var(--viz-critical)',
    'No CA applied': 'var(--viz-critical)',
    'Unmanaged': 'var(--viz-critical)',
    'Phishable': 'var(--viz-serious)',
    'Phone': 'var(--viz-serious)',
    'Authenticator': 'var(--viz-warning)',
    'CA applied': 'var(--viz-s1)',
    'Managed': 'var(--viz-s1)',
    'User sign in': 'var(--viz-s5)',
    'Users': 'var(--viz-s5)',
    'Desktop devices': 'var(--viz-s5)',
    'Mobile devices': 'var(--viz-s5)',
    'Windows': 'var(--viz-s1)',
    'macOS': 'var(--viz-s7)',
    'Android': 'var(--viz-s2)',
    'iOS': 'var(--viz-s3)',
    'Android (Company)': 'var(--viz-s2)',
    'Android (Personal)': 'var(--viz-s2)',
    'iOS (Company)': 'var(--viz-s3)',
    'iOS (Personal)': 'var(--viz-s3)',
    'Entra joined': 'var(--viz-s1)',
    'Entra hybrid joined': 'var(--viz-s8)',
    'Entra registered': 'var(--viz-s3)',
  };

  function statusColor(status) {
    return STATUS_COLOR[status] || 'var(--viz-neutral)';
  }

  // ---- Stat tile -------------------------------------------------------------------------

  function tile(label, value, sub) {
    return `<div class="kpi"><div class="label">${esc(label)}</div><div class="value">${esc(value)}</div>${sub ? `<div class="label">${esc(sub)}</div>` : ''}</div>`;
  }

  // ---- Radial gauge (pillar scorecard) ---------------------------------------------------

  function gauge(opts) {
    const value = Number(opts.value) || 0;
    const total = Number(opts.total) || 0;
    const pct = total > 0 ? value / total : 0;
    const size = 132;
    const r = 52;
    const cx = size / 2;
    const cy = size / 2;
    const circumference = 2 * Math.PI * r;
    const arc = Math.max(0, Math.min(1, pct)) * circumference;
    const pctText = total > 0 ? `${Math.round(pct * 100)}%` : '—';
    const subText = total > 0 ? `${fmt(value)}/${fmt(total)} passed` : 'no scored tests';
    return `
      <div class="viz-gauge" data-tip="${esc(opts.label)}: ${esc(subText)}">
        <svg viewBox="0 0 ${size} ${size}" width="${size}" height="${size}" role="img" aria-label="${esc(opts.label)} ${esc(pctText)} (${esc(subText)})">
          <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--viz-grid)" stroke-width="10"/>
          <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${opts.color || 'var(--viz-s1)'}" stroke-width="10"
            stroke-linecap="round" stroke-dasharray="${arc.toFixed(1)} ${circumference.toFixed(1)}"
            transform="rotate(-90 ${cx} ${cy})"/>
          <text x="${cx}" y="${cy - 2}" text-anchor="middle" class="viz-gauge-value">${esc(pctText)}</text>
          <text x="${cx}" y="${cy + 16}" text-anchor="middle" class="viz-gauge-sub">${esc(subText)}</text>
        </svg>
        <div class="viz-gauge-label">${esc(opts.label)}</div>
      </div>`;
  }

  // ---- Donut -----------------------------------------------------------------------------

  function donut(opts) {
    const segments = (opts.segments || []).filter(s => (Number(s.value) || 0) > 0);
    const total = segments.reduce((a, s) => a + Number(s.value), 0);
    const size = 150;
    const r = 55;
    const cx = size / 2;
    const cy = size / 2;

    let paths = '';
    if (total > 0) {
      let angle = -Math.PI / 2;
      for (const s of segments) {
        const frac = Number(s.value) / total;
        const sweep = frac * 2 * Math.PI;
        const x1 = cx + r * Math.cos(angle);
        const y1 = cy + r * Math.sin(angle);
        const end = angle + sweep;
        const x2 = cx + r * Math.cos(end);
        const y2 = cy + r * Math.sin(end);
        const large = sweep > Math.PI ? 1 : 0;
        // Full-circle single segment: two arcs
        const d = frac >= 0.999
          ? `M ${cx} ${cy - r} A ${r} ${r} 0 1 1 ${cx - 0.01} ${cy - r}`
          : `M ${x1.toFixed(2)} ${y1.toFixed(2)} A ${r} ${r} 0 ${large} 1 ${x2.toFixed(2)} ${y2.toFixed(2)}`;
        const pctLabel = `${Math.round(frac * 100)}%`;
        paths += `<path d="${d}" fill="none" stroke="${s.color}" stroke-width="18" data-tip="${esc(s.label)}: ${fmt(s.value)} (${pctLabel})"/>`;
        // 2px surface gap between segments
        paths += `<line x1="${x2.toFixed(2)}" y1="${y2.toFixed(2)}" x2="${(cx + (r + 12) * Math.cos(end)).toFixed(2)}" y2="${(cy + (r + 12) * Math.sin(end)).toFixed(2)}" stroke="var(--card)" stroke-width="2"/>`;
        angle = end;
      }
    }

    const legend = (opts.segments || []).map(s => `
      <div class="viz-legend-row">
        <span class="viz-swatch" style="background:${s.color}"></span>
        <span class="viz-legend-label">${esc(s.label)}</span>
        <span class="viz-legend-value">${fmt(s.value)}</span>
      </div>`).join('');

    return `
      <div class="viz-donut">
        <svg viewBox="0 0 ${size} ${size}" width="${size}" height="${size}" role="img" aria-label="${esc(opts.title || 'Distribution')}">
          ${paths || `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--viz-grid)" stroke-width="18"/>`}
          <text x="${cx}" y="${cy - 2}" text-anchor="middle" class="viz-gauge-value">${fmt(total)}</text>
          <text x="${cx}" y="${cy + 16}" text-anchor="middle" class="viz-gauge-sub">${esc(opts.centerLabel || 'total')}</text>
        </svg>
        <div class="viz-legend">${legend}</div>
      </div>`;
  }

  // ---- Sankey ----------------------------------------------------------------------------

  function sankeyLayout(links, width, height) {
    const nodes = new Map();
    const ensure = name => {
      if (!nodes.has(name)) nodes.set(name, { name, in: 0, out: 0, level: 0, links: [] });
      return nodes.get(name);
    };
    for (const l of links) {
      ensure(l.source).out += l.value;
      ensure(l.target).in += l.value;
    }
    // Level = longest path from a root, relaxed iteratively (graphs here are tiny DAGs).
    for (let pass = 0; pass < 8; pass += 1) {
      let changed = false;
      for (const l of links) {
        const s = nodes.get(l.source);
        const t = nodes.get(l.target);
        if (t.level < s.level + 1) { t.level = s.level + 1; changed = true; }
      }
      if (!changed) break;
    }

    const maxLevel = Math.max(...Array.from(nodes.values()).map(n => n.level));
    const levels = Array.from({ length: maxLevel + 1 }, () => []);
    for (const n of nodes.values()) {
      n.value = Math.max(n.in, n.out);
      levels[n.level].push(n);
    }

    const nodeWidth = 10;
    const gap = 14;
    const labelSpace = 128; // room for labels right of the final column
    const plotWidth = width - labelSpace;
    const colStep = maxLevel > 0 ? (plotWidth - nodeWidth) / maxLevel : 0;

    const maxColSum = Math.max(...levels.map(col => col.reduce((a, n) => a + n.value, 0)));
    const maxColCount = Math.max(...levels.map(col => col.length));
    const usable = height - 20 - (maxColCount - 1) * gap;
    const scale = maxColSum > 0 ? usable / maxColSum : 0;

    for (const [levelIdx, col] of levels.entries()) {
      const colHeight = col.reduce((a, n) => a + Math.max(2, n.value * scale), 0) + (col.length - 1) * gap;
      let y = Math.max(10, (height - colHeight) / 2);
      for (const n of col) {
        n.x = levelIdx * colStep;
        n.h = Math.max(2, n.value * scale);
        n.y = y;
        n.inOffset = 0;
        n.outOffset = 0;
        y += n.h + gap;
      }
    }

    return { nodes, scale, nodeWidth };
  }

  function sankey(opts) {
    const data = opts.data || {};
    const links = (data.nodes || []).filter(l => (Number(l.value) || 0) > 0)
      .map(l => ({ source: `${l.source}`, target: `${l.target}`, value: Number(l.value) }));
    if (!links.length) {
      return `<div class="viz-empty">${esc(opts.emptyText || 'No data available for this diagram.')}</div>`;
    }

    const width = opts.width || 560;
    const height = opts.height || 240;
    const { nodes, scale, nodeWidth } = sankeyLayout(links, width, height);

    let ribbons = '';
    for (const l of links) {
      const s = nodes.get(l.source);
      const t = nodes.get(l.target);
      const h0 = Math.max(1, l.value * scale);
      const sy = s.y + s.outOffset;
      const ty = t.y + t.inOffset;
      s.outOffset += h0;
      t.inOffset += h0;
      const x0 = s.x + nodeWidth;
      const x1 = t.x;
      const mid = (x0 + x1) / 2;
      const color = SANKEY_NODE_COLORS[t.name] || 'var(--viz-neutral)';
      ribbons += `<path d="M ${x0} ${sy.toFixed(1)} C ${mid} ${sy.toFixed(1)} ${mid} ${ty.toFixed(1)} ${x1} ${ty.toFixed(1)} L ${x1} ${(ty + h0).toFixed(1)} C ${mid} ${(ty + h0).toFixed(1)} ${mid} ${(sy + h0).toFixed(1)} ${x0} ${(sy + h0).toFixed(1)} Z"
        fill="${color}" opacity="0.30" data-tip="${esc(l.source)} → ${esc(l.target)}: ${fmt(l.value)}"/>`;
    }

    let rects = '';
    let labels = '';
    for (const n of nodes.values()) {
      const color = SANKEY_NODE_COLORS[n.name] || 'var(--viz-s1)';
      rects += `<rect x="${n.x}" y="${n.y.toFixed(1)}" width="${nodeWidth}" height="${n.h.toFixed(1)}" rx="3" fill="${color}" data-tip="${esc(n.name)}: ${fmt(n.value)}"/>`;
      const labelX = n.x + nodeWidth + 6;
      const labelY = n.y + n.h / 2 + 4;
      labels += `<text x="${labelX}" y="${labelY.toFixed(1)}" class="viz-node-label">${esc(n.name)} <tspan class="viz-node-value">${fmt(n.value)}</tspan></text>`;
    }

    return `
      <div class="viz-sankey">
        <svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="xMinYMid meet" role="img" aria-label="${esc(opts.title || 'Flow diagram')}">
          ${ribbons}${rects}${labels}
        </svg>
        ${data.description ? `<p class="viz-caption">${esc(data.description)}</p>` : ''}
      </div>`;
  }

  // ---- Heatmap (HTML table doubles as the accessible table view) --------------------------

  const SEQ_STEPS = ['var(--viz-seq-1)', 'var(--viz-seq-2)', 'var(--viz-seq-3)', 'var(--viz-seq-4)', 'var(--viz-seq-5)'];

  function heatmap(opts) {
    const rows = opts.rows || [];
    const cols = opts.cols || [];
    const get = opts.get || (() => 0);
    let max = 0;
    for (const r of rows) for (const c of cols) max = Math.max(max, get(r, c));

    const body = rows.map(r => {
      const cells = cols.map(c => {
        const v = get(r, c);
        if (!max || !v) return `<td class="viz-heat-cell viz-heat-zero">0</td>`;
        const idx = Math.min(SEQ_STEPS.length - 1, Math.floor((v / max) * SEQ_STEPS.length));
        const dark = idx >= 3;
        return `<td class="viz-heat-cell${dark ? ' viz-heat-dark' : ''}" style="background:${SEQ_STEPS[idx]}" data-tip="${esc(r)} × ${esc(c)}: ${fmt(v)}">${fmt(v)}</td>`;
      }).join('');
      return `<tr><th>${esc(r)}</th>${cells}</tr>`;
    }).join('');

    return `
      <div class="md-table-wrap"><table class="viz-heatmap">
        <thead><tr><th></th>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead>
        <tbody>${body}</tbody>
      </table></div>`;
  }

  // ---- Trend lines -----------------------------------------------------------------------

  function trend(opts) {
    const series = (opts.series || []).filter(s => (s.points || []).length);
    if (!series.length) return `<div class="viz-empty">Not enough scans for a trend yet.</div>`;

    const width = opts.width || 620;
    const height = opts.height || 200;
    const padL = 34; const padR = 12; const padT = 12; const padB = 26;
    const plotW = width - padL - padR;
    const plotH = height - padT - padB;
    const nPoints = Math.max(...series.map(s => s.points.length));
    const yMax = opts.yMax != null ? opts.yMax : Math.max(1, ...series.flatMap(s => s.points.map(p => p.value)));

    const x = i => padL + (nPoints > 1 ? (i / (nPoints - 1)) * plotW : plotW / 2);
    const y = v => padT + plotH - (Math.max(0, Math.min(yMax, v)) / yMax) * plotH;

    // Recessive grid: 4 hairlines + y labels
    let grid = '';
    for (let g = 0; g <= 4; g += 1) {
      const gy = padT + (g / 4) * plotH;
      const gv = Math.round(yMax * (1 - g / 4));
      grid += `<line x1="${padL}" y1="${gy}" x2="${width - padR}" y2="${gy}" stroke="var(--viz-grid)" stroke-width="1"/>`;
      grid += `<text x="${padL - 6}" y="${gy + 4}" text-anchor="end" class="viz-axis-label">${gv}${opts.unit || ''}</text>`;
    }

    // Sparse x labels: first, middle, last
    const labels = series[0].points.map(p => p.label);
    let xLabels = '';
    const marks = nPoints <= 3 ? labels.map((_, i) => i) : [0, Math.floor((nPoints - 1) / 2), nPoints - 1];
    for (const i of Array.from(new Set(marks))) {
      if (labels[i] == null) continue;
      xLabels += `<text x="${x(i)}" y="${height - 8}" text-anchor="middle" class="viz-axis-label">${esc(labels[i])}</text>`;
    }

    let paths = '';
    let dots = '';
    for (const s of series) {
      const pts = s.points.map((p, i) => `${x(i).toFixed(1)},${y(p.value).toFixed(1)}`).join(' ');
      paths += `<polyline points="${pts}" fill="none" stroke="${s.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
      s.points.forEach((p, i) => {
        dots += `<circle cx="${x(i).toFixed(1)}" cy="${y(p.value).toFixed(1)}" r="3.5" fill="${s.color}" stroke="var(--card)" stroke-width="2"/>`;
        dots += `<circle cx="${x(i).toFixed(1)}" cy="${y(p.value).toFixed(1)}" r="10" fill="transparent" data-tip="${esc(s.name)} — ${esc(p.label)}: ${p.value}${opts.unit || ''}"/>`;
      });
    }

    const legend = series.length > 1
      ? `<div class="viz-legend viz-legend-inline">${series.map(s => `<span class="viz-legend-row"><span class="viz-swatch" style="background:${s.color}"></span><span class="viz-legend-label">${esc(s.name)}</span></span>`).join('')}</div>`
      : '';

    return `
      <div class="viz-trend">
        ${legend}
        <svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="xMinYMid meet" role="img" aria-label="${esc(opts.title || 'Trend')}">
          ${grid}${paths}${dots}${xLabels}
        </svg>
      </div>`;
  }

  function sparkline(values, width, height) {
    const vals = (values || []).map(v => Number(v) || 0);
    if (vals.length < 2) return '';
    const w = width || 90; const h = height || 22;
    const max = Math.max(1, ...vals);
    const pts = vals.map((v, i) => `${((i / (vals.length - 1)) * (w - 4) + 2).toFixed(1)},${(h - 3 - (v / max) * (h - 6)).toFixed(1)}`).join(' ');
    return `<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" class="viz-sparkline" role="img" aria-label="trend"><polyline points="${pts}" fill="none" stroke="var(--viz-s1)" stroke-width="1.5"/></svg>`;
  }

  // ---- Shared hover tooltip layer ----------------------------------------------------------

  let tooltipEl = null;

  function ensureTooltip() {
    if (tooltipEl) return tooltipEl;
    tooltipEl = document.createElement('div');
    tooltipEl.className = 'viz-tooltip hidden';
    document.body.appendChild(tooltipEl);
    return tooltipEl;
  }

  function attachTooltips(root) {
    const el = ensureTooltip();
    root.addEventListener('mousemove', e => {
      const target = e.target.closest?.('[data-tip]');
      if (!target) { el.classList.add('hidden'); return; }
      el.textContent = target.dataset.tip;
      el.classList.remove('hidden');
      const pad = 12;
      const vw = window.innerWidth;
      let left = e.clientX + pad;
      if (left + el.offsetWidth + pad > vw) left = e.clientX - el.offsetWidth - pad;
      el.style.left = `${left}px`;
      el.style.top = `${e.clientY + pad}px`;
    });
    root.addEventListener('mouseleave', () => el.classList.add('hidden'));
  }

  window.Charts = {
    tile,
    gauge,
    donut,
    sankey,
    heatmap,
    trend,
    sparkline,
    attachTooltips,
    statusColor,
    SERIES,
  };
})();
