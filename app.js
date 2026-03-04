// Utility functions
function n(v) {
  if (v === '' || v === null || v === undefined) return null;
  const x = Number(v);
  return Number.isFinite(x) ? x : null;
}

function fmt(x, d = 4) {
  if (x === null || x === '' || !Number.isFinite(x)) return '—';
  return x.toLocaleString(undefined, { maximumFractionDigits: d, minimumFractionDigits: 0 });
}

function fmtMoney(x) {
  if (x === null || x === '' || !Number.isFinite(x)) return '—';
  return x.toLocaleString(undefined, { style: 'currency', currency: 'USD' });
}

// Tab switching
for (const b of document.querySelectorAll('.tab')) {
  b.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    document.getElementById(b.dataset.tab).classList.add('active');
  });
}

// ========== MARGIN CALCULATOR ==========
// Excel formulas from Margin Calculator sheet:
// B11: =IF(B5<>"",B5,IF(AND(B6<>"",B7<>""),B6-B7,IF(AND(B6<>"",B8<>"",B8<1),B6*(1-B8),IF(AND(B7<>"",B8<>"",B8<1),B7/(B8/(1-B8)),""))
// B12: =IF(B6<>"",B6,IF(AND(B5<>"",B7<>""),B5+B7,IF(AND(B5<>"",B8<>"",B8<1),B5/(1-B8),IF(AND(B7<>"",B8<>"",B8<>0),B7/B8,""))))
// B13: =IF(B7<>"",B7,IF(AND(B12<>"",B11<>""),B12-B11,""))
// B14: =IF(B8<>"",IF(B8>1,B8/100,B8),IF(AND(B13<>"",B12<>"",B12<>0),B13/B12,""))
// B15: =IF(AND(B13<>"",B11<>"",B11<>0),B13/B11,"")

const m = {
  cost: document.getElementById('m_cost'),
  rev: document.getElementById('m_rev'),
  profit: document.getElementById('m_profit'),
  margin: document.getElementById('m_margin'),
  o_cost: document.getElementById('o_cost'),
  o_rev: document.getElementById('o_rev'),
  o_profit: document.getElementById('o_profit'),
  o_margin: document.getElementById('o_margin'),
  o_markup: document.getElementById('o_markup'),
  clear: document.getElementById('m_clear')
};

function isEmpty(v) {
  return v === null || v === '' || v === undefined;
}

function updateMargin() {
  const B5 = n(m.cost.value);
  const B6 = n(m.rev.value);
  const B7 = n(m.profit.value);
  let B8 = n(m.margin.value);
  // Normalize margin input to match Excel behavior: user may enter 80 for 80% or 0.8
  if (B8 !== null && B8 > 1) B8 = B8 / 100;

  // B11 (Cost output)
  let B11;
  if (!isEmpty(B5)) {
    B11 = B5;
  } else if (!isEmpty(B6) && !isEmpty(B7)) {
    B11 = B6 - B7;
  } else if (!isEmpty(B6) && !isEmpty(B8) && B8 < 1) {
    B11 = B6 * (1 - B8);
  } else if (!isEmpty(B7) && !isEmpty(B8) && B8 < 1) {
    B11 = B7 / (B8 / (1 - B8));
  } else {
    B11 = null;
  }

  // B12 (Revenue output)
  let B12;
  if (!isEmpty(B6)) {
    B12 = B6;
  } else if (!isEmpty(B5) && !isEmpty(B7)) {
    B12 = B5 + B7;
  } else if (!isEmpty(B5) && !isEmpty(B8) && B8 < 1) {
    B12 = B5 / (1 - B8);
  } else if (!isEmpty(B7) && !isEmpty(B8) && B8 !== 0) {
    B12 = B7 / B8;
  } else {
    B12 = null;
  }

  // B13 (Profit output)
  let B13;
  if (!isEmpty(B7)) {
    B13 = B7;
  } else if (!isEmpty(B12) && !isEmpty(B11)) {
    B13 = B12 - B11;
  } else {
    B13 = null;
  }

  // B14 (Margin % output)
  let B14;
  if (!isEmpty(B8)) {
    B14 = B8 > 1 ? B8 / 100 : B8;
  } else if (!isEmpty(B13) && !isEmpty(B12) && B12 !== 0) {
    B14 = B13 / B12;
  } else {
    B14 = null;
  }

  // B15 (Markup % output)
  let B15;
  if (!isEmpty(B13) && !isEmpty(B11) && B11 !== 0) {
    B15 = B13 / B11;
  } else {
    B15 = null;
  }

  m.o_cost.textContent = fmtMoney(B11);
  m.o_rev.textContent = fmtMoney(B12);
  m.o_profit.textContent = fmtMoney(B13);
  m.o_margin.textContent = B14 !== null ? fmt(B14 * 100, 2) + '%' : '—';
  m.o_markup.textContent = B15 !== null ? fmt(B15 * 100, 2) + '%' : '—';
}

[m.cost, m.rev, m.profit, m.margin].forEach(el => el.addEventListener('input', updateMargin));
m.clear.addEventListener('click', () => {
  m.cost.value = '';
  m.rev.value = '';
  m.profit.value = '';
  m.margin.value = '';
  updateMargin();
});
updateMargin();

// ========== WEIGHT CALCULATOR ==========
// Excel formulas from Weight Calculator sheet:
// B11: =IF(ROUNDUP(B22*B15,2)>0,ROUNDUP(B22*B15,2)+3,0)
// B13: =IFERROR(_xlfn.SINGLE(_xlfn.XLOOKUP(B5, Materials!A:A, Materials!B:B, 0)), 0)
// B14: =IFERROR(_xlfn.SINGLE(_xlfn.XLOOKUP(B5, Materials!A:A, Materials!C:C, 0)) / 1000000, 0)
// B15: =IF(B3="Metric",B14,B13)
// B17: =B7*B8*B9  (Sheet vol)
// B18: =PI()*(B8/2)^2*B9  (Rod vol)
// B19: =IF(B7>=B8/2,0,PI()*((B8/2)^2-((B8/2)-B7)^2)*B9)  (Round tube)
// B20: =IF(B7>=B8/2,0,(B8^2-(B8-2*B7)^2)*B9)  (Square tube)
// B22: =_xlfn.SWITCH(B4, "Sheet/Bar", B17, "Rod", B18, "Round Tube", B19, "Square Tube", B20, 0)

const w = {
  units: document.getElementById('w_units'),
  shape: document.getElementById('w_shape'),
  material: document.getElementById('w_material'),
  t: document.getElementById('w_t'),
  ww: document.getElementById('w_w'),
  l: document.getElementById('w_l'),
  clear: document.getElementById('w_clear'),
  o_weight: document.getElementById('o_weight'),
  o_d_lb: document.getElementById('o_d_lb'),
  o_d_kg: document.getElementById('o_d_kg'),
  o_d_used: document.getElementById('o_d_used'),
  o_vol: document.getElementById('o_vol'),
  t_unit: document.getElementById('w_t_unit'),
  w_unit: document.getElementById('w_w_unit'),
  l_unit: document.getElementById('w_l_unit')
};

function populateMaterials() {
  for (const mat of MATERIALS) {
    const opt = document.createElement('option');
    opt.value = mat.material;
    opt.textContent = mat.material;
    w.material.appendChild(opt);
  }
  w.material.value = 'Acrylic (PMMA)';
}

function getMat() {
  return MATERIALS.find(x => x.material === w.material.value) || MATERIALS[0];
}

function roundUp(num, decimals) {
  const multiplier = Math.pow(10, decimals);
  return Math.ceil(num * multiplier) / multiplier;
}

function updateWeight() {
  const B3 = w.units.value;
  const B4 = w.shape.value;
  const B5 = w.material.value;
  const B7 = n(w.t.value);
  const B8 = n(w.ww.value);
  const B9 = n(w.l.value);

  const mat = getMat();
  const B13 = mat.density_lb_in3;
  const B14 = mat.density_g_cm3 / 1000000;
  const B15 = B3 === 'Metric' ? B14 : B13;

  const unit = B3 === 'Imperial' ? 'IN' : 'MM';
  const weightUnit = B3 === 'Imperial' ? 'LBS' : 'KGS';
  w.t_unit.textContent = unit;
  w.w_unit.textContent = unit;
  w.l_unit.textContent = unit;

  let B17 = null, B18 = null, B19 = null, B20 = null, B22 = null, B11 = null;

  if (B7 !== null && B8 !== null && B9 !== null) {
    B17 = B7 * B8 * B9;
    B18 = Math.PI * Math.pow(B8 / 2, 2) * B9;
    B19 = B7 >= B8 / 2 ? 0 : Math.PI * (Math.pow(B8 / 2, 2) - Math.pow((B8 / 2) - B7, 2)) * B9;
    B20 = B7 >= B8 / 2 ? 0 : (Math.pow(B8, 2) - Math.pow(B8 - 2 * B7, 2)) * B9;

    switch (B4) {
      case 'Sheet/Bar': B22 = B17; break;
      case 'Rod': B22 = B18; break;
      case 'Round Tube': B22 = B19; break;
      case 'Square Tube': B22 = B20; break;
      default: B22 = 0;
    }

    if (B22 !== null && B15 !== null) {
      const calc = roundUp(B22 * B15, 2);
      B11 = calc > 0 ? calc + 3 : 0;
    }
  }

  w.o_d_lb.textContent = fmt(B13, 5);
  w.o_d_kg.textContent = fmt(B14, 9);
  w.o_d_used.textContent = B15 !== null ? fmt(B15, B3 === 'Metric' ? 9 : 5) : '—';
  w.o_vol.textContent = B22 !== null ? fmt(B22, 4) : '—';
  w.o_weight.textContent = B11 !== null ? fmt(B11, 2) + ' ' + weightUnit : '—';
}

[w.units, w.shape, w.material, w.t, w.ww, w.l].forEach(el => el.addEventListener('input', updateWeight));
w.clear.addEventListener('click', () => {
  w.t.value = '';
  w.ww.value = '';
  w.l.value = '';
  updateWeight();
});

populateMaterials();
updateWeight();


// ========== CUT PLANNER ==========/\/\/
// Simple heuristic cut planner: per selected sheet size, compute parts per sheet, sheets required, yield and waste.

const cp = {
  stockCheckboxes: () => Array.from(document.querySelectorAll('.cp-stock')),
  customEnable: document.getElementById('cp_custom_enable'),
  customW: document.getElementById('cp_custom_w'),
  customL: document.getElementById('cp_custom_l'),
  kerf: document.getElementById('cp_kerf'),
  grain: document.getElementById('cp_grain'),
  partsTable: document.getElementById('cp_parts'),
  addPart: document.getElementById('cp_add_part'),
  clearParts: document.getElementById('cp_clear_parts'),
  results: document.getElementById('cp_results')
};

function cp_getParts() {
  const rows = Array.from(cp.partsTable.querySelector('tbody').rows);
  return rows.map(r => {
    const name = r.querySelector('.cp-name').value.trim() || '(Part)';
    const w = n(r.querySelector('.cp-w').value);
    const l = n(r.querySelector('.cp-l').value);
    const q = n(r.querySelector('.cp-q').value) || 0;
    return { name, w, l, q };
  }).filter(p => p.w && p.l && p.q > 0);
}

function cp_addPartRow(defaults = {}) {
  const tbody = cp.partsTable.querySelector('tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td style="padding:4px 6px;"><input class="cp-name" type="text" style="width:100%;" value="${defaults.name || ''}"></td>
    <td style="padding:4px 6px;text-align:right;"><input class="cp-w" type="number" step="any" style="width:100%;" value="${defaults.w || ''}"></td>
    <td style="padding:4px 6px;text-align:right;"><input class="cp-l" type="number" step="any" style="width:100%;" value="${defaults.l || ''}"></td>
    <td style="padding:4px 6px;text-align:right;"><input class="cp-q" type="number" step="1" style="width:100%;" value="${defaults.q || ''}"></td>
    <td style="padding:4px 6px;text-align:center;"><button class="cp-del-row">Remove</button></td>
  `;
  tbody.appendChild(tr);
}

function cp_sheetConfigs() {
  const kerf = n(cp.kerf.value) || 0;
  const list = [];
  for (const cb of cp.stockCheckboxes()) {
    if (cb.checked) {
      const w = n(cb.dataset.w);
      const l = n(cb.dataset.l);
      if (w && l) list.push({ label: `${w} × ${l}`, w, l, kerf });
    }
  }
  if (cp.customEnable && cp.customEnable.checked) {
    const w = n(cp.customW.value);
    const l = n(cp.customL.value);
    if (w && l) list.push({ label: `${w} × ${l} (custom)`, w, l, kerf });
  }
  return list;
}

function cp_fitOnSheet(sheetW, sheetL, partW, partL, kerf, allowRotate) {
  function fit(w, l) {
    if (!w || !l || w <= 0 || l <= 0) return 0;
    const effW = w + kerf;
    const effL = l + kerf;
    if (effW <= 0 || effL <= 0) return 0;
    const countW = Math.floor((sheetW + kerf) / effW);
    const countL = Math.floor((sheetL + kerf) / effL);
    if (countW <= 0 || countL <= 0) return 0;
    return countW * countL;
  }
  const base = fit(partW, partL);
  if (!allowRotate) return base;
  const rotated = fit(partL, partW);
  return Math.max(base, rotated);
}

function cp_update() {
  if (!cp.results) return;
  const parts = cp_getParts();
  const sheets = cp_sheetConfigs();
  const allowRotate = cp.grain.value === 'rotate';

  if (!parts.length || !sheets.length) {
    cp.results.textContent = 'Select at least one sheet size and enter at least one part with width, length, and quantity.';
    return;
  }

  let html = '<table style="width:100%;border-collapse:collapse;margin-top:6px;">';
  html += '<thead><tr>' +
          '<th style="text-align:left;border-bottom:1px solid rgba(0,0,0,0.15);padding:4px 6px;">Sheet</th>' +
          '<th style="text-align:right;border-bottom:1px solid rgba(0,0,0,0.15);padding:4px 6px;">Sheets req.</th>' +
          '<th style="text-align:right;border-bottom:1px solid rgba(0,0,0,0.15);padding:4px 6px;">Yield %</th>' +
          '<th style="text-align:right;border-bottom:1px solid rgba(0,0,0,0.15);padding:4px 6px;">Waste %</th>' +
          '</tr></thead><tbody>';

  for (const s of sheets) {
    const sheetArea = s.w * s.l;
    let totalPartArea = 0;
    let limitingSheets = 0;

    for (const p of parts) {
      const perSheet = cp_fitOnSheet(s.w, s.l, p.w, p.l, s.kerf, allowRotate);
      if (perSheet <= 0) {
        limitingSheets = Infinity;
        continue;
      }
      const needed = Math.ceil(p.q / perSheet);
      if (needed > limitingSheets) limitingSheets = needed;
      totalPartArea += p.w * p.l * p.q;
    }

    if (!isFinite(limitingSheets) || limitingSheets === 0) {
      html += `<tr><td style="padding:4px 6px;">${s.label}</td><td colspan="3" style="padding:4px 6px;color:#b00;">Parts do not fit on this sheet size with current kerf/grain settings.</td></tr>`;
      continue;
    }

    const usedArea = sheetArea * limitingSheets;
    const yieldPct = usedArea ? (totalPartArea / usedArea * 100) : 0;
    const wastePct = 100 - yieldPct;

    html += `<tr>
      <td style="padding:4px 6px;">${s.label}</td>
      <td style="padding:4px 6px;text-align:right;">${limitingSheets}</td>
      <td style="padding:4px 6px;text-align:right;">${yieldPct.toFixed(1)}%</td>
      <td style="padding:4px 6px;text-align:right;">${wastePct.toFixed(1)}%</td>
    </tr>`;
  }

  html += '</tbody></table>';
  cp.results.innerHTML = html;
}

if (cp.partsTable) {
  cp_addPartRow({ name: 'Panel', w: 24, l: 36, q: 10 });
  cp_addPartRow();

  cp.addPart.addEventListener('click', () => { cp_addPartRow(); cp_update(); });
  cp.clearParts.addEventListener('click', () => {
    const tbody = cp.partsTable.querySelector('tbody');
    tbody.innerHTML = '';
    cp_addPartRow();
    cp_update();
  });

  cp.partsTable.addEventListener('input', cp_update);
  cp.partsTable.addEventListener('click', (e) => {
    if (e.target.classList.contains('cp-del-row')) {
      e.preventDefault();
      const tr = e.target.closest('tr');
      if (tr) tr.remove();
      cp_update();
    }
  });

  cp.stockCheckboxes().forEach(cb => cb.addEventListener('change', cp_update));
  if (cp.customEnable) cp.customEnable.addEventListener('change', cp_update);
  if (cp.customW) cp.customW.addEventListener('input', cp_update);
  if (cp.customL) cp.customL.addEventListener('input', cp_update);
  cp.kerf.addEventListener('input', cp_update);
  cp.grain.addEventListener('change', cp_update);

  cp_update();
}
