/* ── Canteen Manager – App Logic (Firebase/Firestore) ── */
(function () {
  'use strict';

  // ── Helpers ────────────────────────────────────────
  const $ = (s) => document.querySelector(s);
  const $$ = (s) => document.querySelectorAll(s);
  const MONTHS = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];

  function fmt(n) { return '₹' + Number(n).toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 }); }
  function today() { return new Date().toISOString().slice(0, 10); }
  function sanitize(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // ── Generic table sorting ─────────────────────────
  const sortState = {};  // tableId → { col, dir }

  function applySort(arr, tableId, colAccessor) {
    const s = sortState[tableId];
    if (!s) return arr;
    const sorted = [...arr];
    sorted.sort((a, b) => {
      let va = colAccessor(a, s.col);
      let vb = colAccessor(b, s.col);
      if (va == null) va = '';
      if (vb == null) vb = '';
      if (typeof va === 'number' && typeof vb === 'number') return s.dir === 'asc' ? va - vb : vb - va;
      va = String(va).toLowerCase();
      vb = String(vb).toLowerCase();
      return s.dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });
    return sorted;
  }

  function bindSortHeaders(tableId, renderFn) {
    const table = document.getElementById(tableId);
    if (!table) return;
    table.querySelectorAll('th.sortable').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.dataset.col;
        const prev = sortState[tableId];
        if (prev && prev.col === col) {
          prev.dir = prev.dir === 'asc' ? 'desc' : 'asc';
        } else {
          sortState[tableId] = { col, dir: 'asc' };
        }
        // Update header classes
        table.querySelectorAll('th.sortable').forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
        th.classList.add('sort-' + sortState[tableId].dir);
        renderFn();
      });
    });
  }

  // ── XLSX export helper ────────────────────────────
  function exportXlsx(data, headers, fileName) {
    if (!data.length) return toast('No data to export', true);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    // Auto-width columns
    ws['!cols'] = headers.map((h, i) => {
      let maxW = h.length;
      data.forEach(row => { const len = String(row[i] ?? '').length; if (len > maxW) maxW = len; });
      return { wch: Math.min(maxW + 2, 40) };
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, fileName + '_' + today() + '.xlsx');
    toast('Exported to .xlsx');
  }

  // ── In-memory data (synced from Firestore) ────────
  let itemsIn = [];
  let itemsOut = [];
  let expenses = [];
  let sales = [];

  // ── Firestore references ──────────────────────────
  const colItemsIn  = db.collection('itemsIn');
  const colItemsOut = db.collection('itemsOut');
  const colExpenses = db.collection('expenses');
  const colSales    = db.collection('sales');
  const colAuthLogs  = db.collection('authLogs');

  // ── Toast ─────────────────────────────────────────
  function toast(msg, isError) {
    const el = $('#toast');
    el.textContent = msg;
    el.className = 'toast show' + (isError ? ' error' : '');
    clearTimeout(el._t);
    el._t = setTimeout(() => el.className = 'toast', 2800);
  }

  // ── Navigation ────────────────────────────────────
  const navItems = $$('.nav-item');
  const pages = $$('.page');
  const titles = { dashboard:'Dashboard', 'items-in':'Items In', 'items-out':'Items Out', inventory:'Inventory', expenses:'Other Expenses', sales:'Sales', pnl:'P&L Statement', changelog:'Changelog' };

  function navigate(page) {
    navItems.forEach(n => n.classList.toggle('active', n.dataset.page === page));
    pages.forEach(p => { p.classList.toggle('active', p.id === 'page-' + page); });
    $('#pageTitle').textContent = titles[page] || 'Dashboard';
    if (page === 'dashboard') refreshDashboard();
    if (page === 'items-in')  renderItemsIn();
    if (page === 'items-out') renderItemsOut();
    if (page === 'inventory') renderInventory();
    if (page === 'expenses')  renderExpenses();
    if (page === 'sales')     renderSales();
    if (page === 'pnl')       renderPnl();
    if (page === 'changelog') renderAuthLog();
    $('#sidebar').classList.remove('open');
  }

  navItems.forEach(n => n.addEventListener('click', (e) => { e.preventDefault(); navigate(n.dataset.page); }));
  $('#menuToggle').addEventListener('click', () => $('#sidebar').classList.toggle('open'));

  $('#dateDisplay').textContent = new Date().toLocaleDateString('en-IN', { weekday:'long', year:'numeric', month:'long', day:'numeric' });
  $$('input[type="date"]').forEach(inp => inp.value = today());

  // ── Real-time Firestore listeners ─────────────────
  function startListeners() {
    colItemsIn.orderBy('createdAt', 'desc').onSnapshot((snap) => {
      itemsIn = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      refreshActivePage();
    });

    colItemsOut.orderBy('createdAt', 'desc').onSnapshot((snap) => {
      itemsOut = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      refreshActivePage();
    });

    colExpenses.orderBy('createdAt', 'desc').onSnapshot((snap) => {
      expenses = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      refreshActivePage();
    });

    colSales.orderBy('createdAt', 'desc').onSnapshot((snap) => {
      sales = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      refreshActivePage();
    });
  }

  function refreshActivePage() {
    const activePage = document.querySelector('.nav-item.active')?.dataset.page;
    if (activePage === 'dashboard') refreshDashboard();
    if (activePage === 'items-in')  renderItemsIn();
    if (activePage === 'items-out') renderItemsOut();
    if (activePage === 'inventory') renderInventory();
    if (activePage === 'expenses')  renderExpenses();
    if (activePage === 'sales')     renderSales();
    if (activePage === 'pnl')       renderPnl();
    if (activePage === 'changelog') renderAuthLog();
  }

  // ── Items In ──────────────────────────────────────
  // Auto-calculate total cost = qty × rate
  function autoCalcCost() {
    const qty = parseFloat($('#inQty').value) || 0;
    const rate = parseFloat($('#inRate').value) || 0;
    $('#inPrice').value = qty && rate ? (qty * rate).toFixed(2) : '';
  }
  $('#inQty').addEventListener('input', autoCalcCost);
  $('#inRate').addEventListener('input', autoCalcCost);

  $('#formItemsIn').addEventListener('submit', (e) => {
    e.preventDefault();
    const qty = parseFloat($('#inQty').value);
    const rate = parseFloat($('#inRate').value);
    const record = {
      date: $('#inDate').value,
      billNo: $('#inBillNo').value.trim(),
      item: $('#inItem').value.trim(),
      brand: $('#inBrand').value.trim(),
      qty: qty,
      unit: $('#inUnit').value,
      rate: rate,
      cost: qty * rate,
      supplier: $('#inSupplier').value.trim(),
      remark: $('#inRemark').value.trim(),
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    };
    if (!record.item || !record.qty || !record.rate) return toast('Please fill all required fields', true);
    colItemsIn.add(record).then(() => {
      e.target.reset();
      $$('input[type="date"]').forEach(inp => inp.value = today());
      toast('Item added to stock');
    }).catch(() => toast('Failed to save', true));
  });

  function getItemsInFilters() {
    return {
      dateFrom: $('#filterInDateFrom').value,
      dateTo: $('#filterInDateTo').value,
      item: $('#filterInItem').value.trim().toLowerCase(),
      brand: $('#filterInBrand').value.trim().toLowerCase(),
      supplier: $('#filterInSupplier').value.trim().toLowerCase(),
      billNo: $('#filterInBillNo').value.trim().toLowerCase()
    };
  }

  function renderItemsIn() {
    const f = getItemsInFilters();
    const tbody = $('#tableItemsIn tbody');
    let filtered = itemsIn.filter(r => {
      if (f.dateFrom && r.date < f.dateFrom) return false;
      if (f.dateTo && r.date > f.dateTo) return false;
      if (f.item && !r.item.toLowerCase().includes(f.item)) return false;
      if (f.brand && !(r.brand || '').toLowerCase().includes(f.brand)) return false;
      if (f.supplier && !(r.supplier || '').toLowerCase().includes(f.supplier)) return false;
      if (f.billNo && !(r.billNo || '').toLowerCase().includes(f.billNo)) return false;
      return true;
    });
    filtered = applySort(filtered, 'tableItemsIn', (r, col) => {
      if (col === 'qty' || col === 'rate' || col === 'cost') return r[col] || 0;
      return r[col] || '';
    });
    lastFilteredItemsIn = filtered;
    tbody.innerHTML = filtered.length === 0
      ? '<tr><td colspan="11" style="text-align:center;color:var(--text-light);padding:32px">No purchase records yet</td></tr>'
      : filtered.map(r => `<tr>
          <td>${sanitize(r.date)}</td><td>${sanitize(r.billNo || '—')}</td><td>${sanitize(r.item)}</td>
          <td>${sanitize(r.brand || '—')}</td><td>${sanitize(r.supplier || '—')}</td>
          <td>${r.qty}</td><td>${sanitize(r.unit)}</td>
          <td>${fmt(r.rate || 0)}</td><td>${fmt(r.cost)}</td>
          <td>${sanitize(r.remark || '—')}</td>
          <td><button class="btn-delete" data-id="${sanitize(r.id)}" data-type="in">Delete</button></td>
        </tr>`).join('');
    renderItemsInSummary();
  }
  let lastFilteredItemsIn = [];

  // Filter listeners for Items In
  ['filterInDateFrom','filterInDateTo','filterInItem','filterInBrand','filterInSupplier','filterInBillNo'].forEach(id => {
    $('#' + id).addEventListener('input', () => renderItemsIn());
  });
  $('#clearFiltersIn').addEventListener('click', () => {
    ['filterInDateFrom','filterInDateTo','filterInItem','filterInBrand','filterInSupplier','filterInBillNo'].forEach(id => {
      $('#' + id).value = '';
    });
    renderItemsIn();
  });

  // ── Items In: Month-wise Summary Tables ───────────
  function renderItemsInSummary() {
    const yearSel = $('#itemsInFY');
    const years = getFYears();
    const prev = yearSel.value;
    yearSel.innerHTML = years.map(y => `<option value="${y}">FY ${fyLabel(y)}</option>`).join('');
    if (prev && years.includes(parseInt(prev))) yearSel.value = prev;
    drawItemsInSummary();
  }

  function drawItemsInSummary() {
    const fyStart = parseInt($('#itemsInFY').value);

    function fyIdx(dateStr) {
      const d = new Date(dateStr);
      const m = d.getMonth(), y = d.getFullYear();
      if (m >= 3) return y === fyStart ? m - 3 : -1;
      return y === fyStart + 1 ? m + 9 : -1;
    }

    // Build maps: item→{unit, qty[12], amt[12]}, brand→amt[12], supplier→amt[12]
    const itemQtyMap = {};
    const itemAmtMap = {};
    const brandMap = {};
    const supplierMap = {};

    itemsIn.forEach(r => {
      const idx = fyIdx(r.date);
      if (idx < 0) return;
      const key = r.item.toLowerCase();
      const brandKey = (r.brand || 'Unbranded').trim();
      const suppKey = (r.supplier || 'Unknown').trim();

      // Item Qty
      if (!itemQtyMap[key]) itemQtyMap[key] = { name: r.item, unit: r.unit, months: MONTHS.map(() => 0) };
      itemQtyMap[key].months[idx] += r.qty;

      // Item Amt
      if (!itemAmtMap[key]) itemAmtMap[key] = { name: r.item, months: MONTHS.map(() => 0) };
      itemAmtMap[key].months[idx] += r.cost;

      // Brand Amt
      const bk = brandKey.toLowerCase();
      if (!brandMap[bk]) brandMap[bk] = { name: brandKey, months: MONTHS.map(() => 0) };
      brandMap[bk].months[idx] += r.cost;

      // Supplier Amt
      const sk = suppKey.toLowerCase();
      if (!supplierMap[sk]) supplierMap[sk] = { name: suppKey, months: MONTHS.map(() => 0) };
      supplierMap[sk].months[idx] += r.cost;
    });

    lastInSummary = { itemQty: Object.values(itemQtyMap), itemAmt: Object.values(itemAmtMap), brand: Object.values(brandMap), supplier: Object.values(supplierMap) };

    // Helper to render a grouped table
    function renderGroupTable(tbodyId, rows, valFmt, totalRowId) {
      const tbody = $('#' + tbodyId + ' tbody');
      const monthTotals = MONTHS.map(() => 0);
      let grandTotal = 0;
      const hasUnit = rows.length > 0 && rows[0].unit !== undefined;

      tbody.innerHTML = rows.length === 0
        ? `<tr><td colspan="${hasUnit ? 15 : 14}" style="text-align:center;color:var(--text-light);padding:32px">No data for this FY</td></tr>`
        : rows.filter(r => r.months.some(v => v)).map(r => {
            const rowTotal = r.months.reduce((s, v) => s + v, 0);
            r.months.forEach((v, i) => { monthTotals[i] += v; });
            grandTotal += rowTotal;
            return `<tr>
              <td><strong>${sanitize(r.name)}</strong></td>
              ${hasUnit ? `<td>${sanitize(r.unit)}</td>` : ''}
              ${r.months.map(v => `<td>${v ? valFmt(v) : '—'}</td>`).join('')}
              <td><strong>${valFmt(rowTotal)}</strong></td>
            </tr>`;
          }).join('');

      if (totalRowId && grandTotal > 0) {
        $(totalRowId).innerHTML = `
          <td><strong>TOTAL</strong></td>
          ${hasUnit ? '<td></td>' : ''}
          ${monthTotals.map(v => `<td><strong>${v ? valFmt(v) : '—'}</strong></td>`).join('')}
          <td><strong>${valFmt(grandTotal)}</strong></td>`;
      } else if (totalRowId) {
        $(totalRowId).innerHTML = '';
      }
    }

    const fmtQty = v => v % 1 === 0 ? String(v) : v.toFixed(2);

    renderGroupTable('tableInQtyMonthly', Object.values(itemQtyMap), fmtQty, null);
    renderGroupTable('tableInAmtMonthly', Object.values(itemAmtMap), fmt, '#inAmtTotalRow');
    renderGroupTable('tableInBrandMonthly', Object.values(brandMap), fmt, '#inBrandTotalRow');
    renderGroupTable('tableInSupplierMonthly', Object.values(supplierMap), fmt, '#inSupplierTotalRow');
  }
  let lastInSummary = { itemQty: [], itemAmt: [], brand: [], supplier: [] };

  $('#itemsInFY').addEventListener('change', drawItemsInSummary);

  $('#btnExportItemsInSummary').addEventListener('click', () => {
    const s = lastInSummary;
    const fy = $('#itemsInFY').value;
    const rows = [];
    // Sheet 1 data: all 4 tables in one sheet separated by headers
    rows.push(['Quantity of Items Purchased — Month-wise']);
    rows.push(['Item', 'Unit', ...MONTHS, 'Total']);
    s.itemQty.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, r.unit, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);
    rows.push(['Amount of Items Purchased — Month-wise']);
    rows.push(['Item', ...MONTHS, 'Total']);
    s.itemAmt.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);
    rows.push(['Amount Purchased from Brands — Month-wise']);
    rows.push(['Brand', ...MONTHS, 'Total']);
    s.brand.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);
    rows.push(['Amount Purchased from Suppliers — Month-wise']);
    rows.push(['Supplier', ...MONTHS, 'Total']);
    s.supplier.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });

    if (rows.length <= 8) return toast('No data to export', true);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Purchase Summaries');
    XLSX.writeFile(wb, 'canteen_purchase_summaries_FY' + fyLabel(fy) + '_' + today() + '.xlsx');
    toast('Exported to .xlsx');
  });

  // ── Items Out ─────────────────────────────────────
  $('#formItemsOut').addEventListener('submit', (e) => {
    e.preventDefault();
    const record = {
      date: $('#outDate').value,
      item: $('#outItem').value.trim(),
      brand: $('#outBrand').value.trim(),
      supplier: $('#outSupplier').value.trim(),
      qty: parseFloat($('#outQty').value),
      unit: $('#outUnit').value,
      rate: parseFloat($('#outRate').value) || 0,
      amount: parseFloat($('#outPrice').value),
      category: $('#outCategory').value,
      person: $('#outPerson').value,
      customer: $('#outCustomer').value.trim(),
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    };
    if (!record.item || !record.qty || !record.amount) return toast('Please fill all required fields', true);
    colItemsOut.add(record).then(() => {
      e.target.reset();
      $$('input[type="date"]').forEach(inp => inp.value = today());
      toast('Sale recorded');
    }).catch(() => toast('Failed to save', true));
  });

  function getItemsOutFilters() {
    return {
      dateFrom: $('#filterOutDateFrom').value,
      dateTo: $('#filterOutDateTo').value,
      item: $('#filterOutItem').value.trim().toLowerCase(),
      category: $('#filterOutCategory').value,
      person: $('#filterOutPerson').value
    };
  }

  function renderItemsOut() {
    const f = getItemsOutFilters();
    const tbody = $('#tableItemsOut tbody');
    let filtered = itemsOut.filter(r => {
      if (f.dateFrom && r.date < f.dateFrom) return false;
      if (f.dateTo && r.date > f.dateTo) return false;
      if (f.item && !r.item.toLowerCase().includes(f.item)) return false;
      if (f.category && (r.category || '') !== f.category) return false;
      if (f.person && (r.person || '') !== f.person) return false;
      return true;
    });
    filtered = applySort(filtered, 'tableItemsOut', (r, col) => {
      if (col === 'qty' || col === 'amount' || col === 'rate') return r[col] || 0;
      return r[col] || '';
    });
    lastFilteredItemsOut = filtered;
    tbody.innerHTML = filtered.length === 0
      ? '<tr><td colspan="12" style="text-align:center;color:var(--text-light);padding:32px">No sales records yet</td></tr>'
      : filtered.map(r => `<tr>
          <td>${sanitize(r.date)}</td><td>${sanitize(r.item)}</td>
          <td>${sanitize(r.brand || '—')}</td><td>${sanitize(r.supplier || '—')}</td>
          <td>${r.qty}</td><td>${sanitize(r.unit)}</td>
          <td>${fmt(r.rate || 0)}</td><td>${fmt(r.amount)}</td>
          <td>${sanitize(r.category || '—')}</td><td>${sanitize(r.person || '—')}</td>
          <td>${sanitize(r.customer || '—')}</td>
          <td><button class="btn-delete" data-id="${sanitize(r.id)}" data-type="out">Delete</button></td>
        </tr>`).join('');
  }
  let lastFilteredItemsOut = [];

  ['filterOutDateFrom','filterOutDateTo','filterOutItem','filterOutCategory','filterOutPerson'].forEach(id => {
    $('#' + id).addEventListener('input', () => renderItemsOut());
    $('#' + id).addEventListener('change', () => renderItemsOut());
  });
  $('#clearFiltersOut').addEventListener('click', () => {
    ['filterOutDateFrom','filterOutDateTo','filterOutItem'].forEach(id => { $('#' + id).value = ''; });
    $('#filterOutCategory').value = '';
    $('#filterOutPerson').value = '';
    renderItemsOut();
  });

  // ── Expenses ──────────────────────────────────────
  $('#formExpenses').addEventListener('submit', (e) => {
    e.preventDefault();
    const record = {
      date: $('#expDate').value,
      category: $('#expCategory').value,
      amount: parseFloat($('#expAmount').value),
      note: $('#expNote').value.trim(),
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    };
    if (!record.amount) return toast('Please enter an amount', true);
    colExpenses.add(record).then(() => {
      e.target.reset();
      $$('input[type="date"]').forEach(inp => inp.value = today());
      toast('Expense recorded');
    }).catch(() => toast('Failed to save', true));
  });

  const EXP_CATEGORIES = ['Salary','Rent','Electricity','Gas','Water','Maintenance','Equipment','Transport','Miscellaneous'];

  function getExpFilters() {
    return {
      dateFrom: $('#filterExpDateFrom').value,
      dateTo: $('#filterExpDateTo').value,
      category: $('#filterExpCategory').value,
      desc: $('#filterExpDesc').value.trim().toLowerCase()
    };
  }

  function renderExpenses() {
    const f = getExpFilters();
    const tbody = $('#tableExpenses tbody');
    let filtered = expenses.filter(r => {
      if (f.dateFrom && r.date < f.dateFrom) return false;
      if (f.dateTo && r.date > f.dateTo) return false;
      if (f.category && r.category !== f.category) return false;
      if (f.desc && !(r.note || '').toLowerCase().includes(f.desc)) return false;
      return true;
    });
    filtered = applySort(filtered, 'tableExpenses', (r, col) => {
      if (col === 'amount') return r[col] || 0;
      return r[col] || '';
    });
    lastFilteredExpenses = filtered;
    tbody.innerHTML = filtered.length === 0
      ? '<tr><td colspan="5" style="text-align:center;color:var(--text-light);padding:32px">No expenses recorded yet</td></tr>'
      : filtered.map(r => `<tr>
          <td>${sanitize(r.date)}</td><td>${sanitize(r.category)}</td><td>${fmt(r.amount)}</td>
          <td>${sanitize(r.note || '—')}</td>
          <td><button class="btn-delete" data-id="${sanitize(r.id)}" data-type="exp">Delete</button></td>
        </tr>`).join('');

    renderExpMonthly();
  }
  let lastFilteredExpenses = [];

  ['filterExpDateFrom','filterExpDateTo','filterExpDesc'].forEach(id => {
    $('#' + id).addEventListener('input', () => renderExpenses());
  });
  $('#filterExpCategory').addEventListener('change', () => renderExpenses());
  $('#clearFiltersExp').addEventListener('click', () => {
    ['filterExpDateFrom','filterExpDateTo','filterExpDesc'].forEach(id => { $('#' + id).value = ''; });
    $('#filterExpCategory').value = '';
    renderExpenses();
  });

  // ── Month-wise Category Expense ───────────────────
  function renderExpMonthly() {
    const yearSel = $('#expMonthFY');
    const years = getFYears();
    const prev = yearSel.value;
    yearSel.innerHTML = years.map(y => `<option value="${y}">FY ${fyLabel(y)}</option>`).join('');
    if (prev && years.includes(parseInt(prev))) yearSel.value = prev;

    drawExpMonthly();
  }

  let lastExpMonthlyData = [];

  function drawExpMonthly() {
    const fyStart = parseInt($('#expMonthFY').value);
    // Build a map: category → [12 months]
    const catMap = {};
    EXP_CATEGORIES.forEach(c => { catMap[c] = MONTHS.map(() => 0); });

    function fyIdx(dateStr) {
      const d = new Date(dateStr);
      const m = d.getMonth(), y = d.getFullYear();
      if (m >= 3) return y === fyStart ? m - 3 : -1;
      return y === fyStart + 1 ? m + 9 : -1;
    }

    expenses.forEach(r => {
      const idx = fyIdx(r.date);
      if (idx < 0) return;
      if (!catMap[r.category]) catMap[r.category] = MONTHS.map(() => 0);
      catMap[r.category][idx] += r.amount;
    });

    const tbody = $('#tableExpMonthly tbody');
    const monthTotals = MONTHS.map(() => 0);
    let grandTotal = 0;
    lastExpMonthlyData = [];

    tbody.innerHTML = EXP_CATEGORIES.map(cat => {
      const row = catMap[cat];
      const rowTotal = row.reduce((s, v) => s + v, 0);
      if (rowTotal === 0) return ''; // skip categories with no data
      row.forEach((v, i) => { monthTotals[i] += v; });
      grandTotal += rowTotal;
      lastExpMonthlyData.push({ category: cat, months: [...row], total: rowTotal });
      return `<tr>
        <td><strong>${sanitize(cat)}</strong></td>
        ${row.map(v => `<td>${v ? fmt(v) : '—'}</td>`).join('')}
        <td><strong>${fmt(rowTotal)}</strong></td>
      </tr>`;
    }).join('');

    if (grandTotal === 0) {
      tbody.innerHTML = '<tr><td colspan="14" style="text-align:center;color:var(--text-light);padding:32px">No expenses for this financial year</td></tr>';
      $('#expMonthlyTotalRow').innerHTML = '';
    } else {
      // Add totals row to export data
      lastExpMonthlyData.push({ category: 'TOTAL', months: [...monthTotals], total: grandTotal });
      $('#expMonthlyTotalRow').innerHTML = `
        <td><strong>TOTAL</strong></td>
        ${monthTotals.map(v => `<td><strong>${v ? fmt(v) : '—'}</strong></td>`).join('')}
        <td><strong>${fmt(grandTotal)}</strong></td>`;
    }
  }

  $('#expMonthFY').addEventListener('change', drawExpMonthly);

  // ── Delete handler (delegated) ────────────────────
  document.addEventListener('click', (e) => {
    if (!e.target.classList.contains('btn-delete')) return;
    if (!confirm('Delete this record?')) return;
    const id = e.target.dataset.id;
    const type = e.target.dataset.type;
    let promise;
    if (type === 'in')  promise = colItemsIn.doc(id).delete();
    if (type === 'out') promise = colItemsOut.doc(id).delete();
    if (type === 'exp') promise = colExpenses.doc(id).delete();
    if (type === 'sale') promise = colSales.doc(id).delete();
    if (promise) promise.then(() => toast('Record deleted')).catch(() => toast('Failed to delete', true));
  });

  // ── Inventory ─────────────────────────────────────
  function buildInventory() {
    const inv = {};
    itemsIn.forEach(r => {
      const key = r.item.toLowerCase();
      if (!inv[key]) inv[key] = { name: r.item, brand: r.brand || '', unit: r.unit, supplier: r.supplier || '', qtyIn: 0, qtyOut: 0, totalCost: 0 };
      inv[key].qtyIn += r.qty;
      inv[key].totalCost += r.cost;
      if (r.brand && !inv[key].brand) inv[key].brand = r.brand;
      if (r.supplier && !inv[key].supplier) inv[key].supplier = r.supplier;
    });
    itemsOut.forEach(r => {
      const key = r.item.toLowerCase();
      if (!inv[key]) inv[key] = { name: r.item, brand: '', unit: r.unit, supplier: '', qtyIn: 0, qtyOut: 0, totalCost: 0 };
      inv[key].qtyOut += r.qty;
    });
    return Object.values(inv).map(i => {
      const balance = Math.max(0, i.qtyIn - i.qtyOut);
      const avgCost = i.qtyIn > 0 ? i.totalCost / i.qtyIn : 0;
      let status;
      if (balance <= 0) status = 'Out';
      else if (balance < i.qtyIn * 0.2) status = 'Low';
      else status = 'OK';
      return { ...i, balance, avgCost, value: balance * avgCost, status };
    });
  }

  function getInvFilters() {
    return {
      item: $('#filterInvItem').value.trim().toLowerCase(),
      brand: $('#filterInvBrand').value.trim().toLowerCase(),
      supplier: $('#filterInvSupplier').value.trim().toLowerCase(),
      status: $('#filterInvStatus').value
    };
  }

  function renderInventory() {
    const f = getInvFilters();
    let data = buildInventory().filter(i => {
      if (f.item && !i.name.toLowerCase().includes(f.item)) return false;
      if (f.brand && !i.brand.toLowerCase().includes(f.brand)) return false;
      if (f.supplier && !i.supplier.toLowerCase().includes(f.supplier)) return false;
      if (f.status && i.status !== f.status) return false;
      return true;
    });
    data = applySort(data, 'tableInventory', (r, col) => {
      if (['qtyIn','qtyOut','balance','avgCost','value'].includes(col)) return r[col] || 0;
      return r[col] || '';
    });
    lastFilteredInventory = data;
    const tbody = $('#tableInventory tbody');
    let lowCount = 0;

    tbody.innerHTML = data.length === 0
      ? '<tr><td colspan="10" style="text-align:center;color:var(--text-light);padding:32px">No inventory data</td></tr>'
      : data.map(i => {
          let cls;
          if (i.status === 'Out') cls = 'status-out';
          else if (i.status === 'Low') { cls = 'status-low'; lowCount++; }
          else cls = 'status-ok';
          return `<tr>
            <td><strong>${sanitize(i.name)}</strong></td>
            <td>${sanitize(i.brand || '—')}</td><td>${sanitize(i.supplier || '—')}</td>
            <td>${i.qtyIn.toFixed(2)}</td><td>${i.qtyOut.toFixed(2)}</td>
            <td><strong>${i.balance.toFixed(2)}</strong></td><td>${sanitize(i.unit)}</td>
            <td>${fmt(i.avgCost)}</td><td>${fmt(i.value)}</td>
            <td><span class="status-badge ${cls}">${i.status}</span></td>
          </tr>`;
        }).join('');

    // Count low stock from unfiltered data for accurate stats
    const allData = buildInventory();
    $('#invTotalItems').textContent = allData.length;
    $('#invLowStock').textContent = allData.filter(i => i.status === 'Low').length;
    const totalVal = allData.reduce((s, i) => s + i.value, 0);
    $('#invTotalValue').textContent = fmt(totalVal);
  }
  let lastFilteredInventory = [];

  ['filterInvItem','filterInvBrand','filterInvSupplier'].forEach(id => {
    $('#' + id).addEventListener('input', () => renderInventory());
  });
  $('#filterInvStatus').addEventListener('change', () => renderInventory());
  $('#clearFiltersInv').addEventListener('click', () => {
    ['filterInvItem','filterInvBrand','filterInvSupplier'].forEach(id => { $('#' + id).value = ''; });
    $('#filterInvStatus').value = '';
    renderInventory();
  });

  // ── Export handlers ────────────────────────────────
  $('#btnExportInventory').addEventListener('click', () => {
    const data = lastFilteredInventory;
    exportXlsx(
      data.map(i => [i.name, i.brand, i.supplier, i.qtyIn, i.qtyOut, i.balance, i.unit, +i.avgCost.toFixed(2), +i.value.toFixed(2), i.status]),
      ['Item','Brand','Supplier','Qty In','Qty Out','Balance','Unit','Rate/Unit','Stock Value','Status'],
      'canteen_inventory'
    );
  });

  $('#btnExportItemsIn').addEventListener('click', () => {
    exportXlsx(
      lastFilteredItemsIn.map(r => [r.date, r.billNo || '', r.item, r.brand || '', r.supplier || '', r.qty, r.unit, r.rate || 0, r.cost, r.remark || '']),
      ['Date','Bill No','Item','Brand','Supplier','Qty','Unit','Rate','Cost','Remark'],
      'canteen_items_in'
    );
  });

  $('#btnExportItemsOut').addEventListener('click', () => {
    exportXlsx(
      lastFilteredItemsOut.map(r => [r.date, r.item, r.brand || '', r.supplier || '', r.qty, r.unit, r.rate || 0, r.amount, r.category || '', r.person || '', r.customer || '']),
      ['Date','Item','Brand','Supplier','Qty','Unit','Rate','Amount','Category','Person','Remark'],
      'canteen_items_out'
    );
  });

  $('#btnExportExpenses').addEventListener('click', () => {
    exportXlsx(
      lastFilteredExpenses.map(r => [r.date, r.category, r.amount, r.note || '']),
      ['Date','Category','Amount','Description'],
      'canteen_expenses'
    );
  });

  // ── Sales ─────────────────────────────────────────
  $('#formSales').addEventListener('submit', (e) => {
    e.preventDefault();
    const record = {
      date: $('#saleDate').value,
      type: $('#saleType').value,
      amount: parseFloat($('#saleAmount').value),
      details: $('#saleDetails').value.trim(),
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    };
    if (!record.amount) return toast('Please enter an amount', true);
    colSales.add(record).then(() => {
      e.target.reset();
      $$('input[type="date"]').forEach(inp => inp.value = today());
      toast('Sale recorded');
    }).catch(() => toast('Failed to save', true));
  });

  function getSalesFilters() {
    return {
      dateFrom: $('#filterSaleDateFrom').value,
      dateTo: $('#filterSaleDateTo').value,
      type: $('#filterSaleType').value,
      details: $('#filterSaleDetails').value.trim().toLowerCase()
    };
  }

  function renderSales() {
    const f = getSalesFilters();
    const tbody = $('#tableSales tbody');
    let filtered = sales.filter(r => {
      if (f.dateFrom && r.date < f.dateFrom) return false;
      if (f.dateTo && r.date > f.dateTo) return false;
      if (f.type && r.type !== f.type) return false;
      if (f.details && !(r.details || '').toLowerCase().includes(f.details)) return false;
      return true;
    });
    filtered = applySort(filtered, 'tableSales', (r, col) => {
      if (col === 'amount') return r[col] || 0;
      return r[col] || '';
    });
    lastFilteredSales = filtered;
    tbody.innerHTML = filtered.length === 0
      ? '<tr><td colspan="5" style="text-align:center;color:var(--text-light);padding:32px">No sales recorded yet</td></tr>'
      : filtered.map(r => `<tr>
          <td>${sanitize(r.date)}</td><td>${sanitize(r.type)}</td><td>${fmt(r.amount)}</td>
          <td>${sanitize(r.details || '—')}</td>
          <td><button class="btn-delete" data-id="${sanitize(r.id)}" data-type="sale">Delete</button></td>
        </tr>`).join('');
  }
  let lastFilteredSales = [];

  ['filterSaleDateFrom','filterSaleDateTo','filterSaleDetails'].forEach(id => {
    $('#' + id).addEventListener('input', () => renderSales());
  });
  $('#filterSaleType').addEventListener('change', () => renderSales());
  $('#clearFiltersSales').addEventListener('click', () => {
    ['filterSaleDateFrom','filterSaleDateTo','filterSaleDetails'].forEach(id => { $('#' + id).value = ''; });
    $('#filterSaleType').value = '';
    renderSales();
  });

  $('#btnExportSales').addEventListener('click', () => {
    exportXlsx(
      lastFilteredSales.map(r => [r.date, r.type, r.amount, r.details || '']),
      ['Date','Sale Type','Amount','Details'],
      'canteen_sales'
    );
  });

  // ── P&L Computation (Financial Year: Apr–Mar) ────
  // fyStart = the calendar year in which April falls
  // e.g. FY 2025-26 → fyStart=2025 covers Apr 2025 – Mar 2026
  function computeMonthlyPnl(fyStart) {
    const monthly = MONTHS.map(() => ({ revenue: 0, itemCost: 0, otherExp: 0 }));

    function fyIndex(dateStr) {
      const d = new Date(dateStr);
      const m = d.getMonth(); // 0=Jan .. 11=Dec
      const y = d.getFullYear();
      // Apr(3)=idx0 .. Dec(11)=idx8, Jan(0)=idx9 .. Mar(2)=idx11
      if (m >= 3) { // Apr-Dec → belongs to FY starting in same year
        return y === fyStart ? m - 3 : -1;
      } else { // Jan-Mar → belongs to FY starting in previous year
        return y === fyStart + 1 ? m + 9 : -1;
      }
    }

    itemsOut.forEach(r => {
      const idx = fyIndex(r.date);
      if (idx >= 0) monthly[idx].revenue += r.amount;
    });
    itemsIn.forEach(r => {
      const idx = fyIndex(r.date);
      if (idx >= 0) monthly[idx].itemCost += r.cost;
    });
    expenses.forEach(r => {
      const idx = fyIndex(r.date);
      if (idx >= 0) monthly[idx].otherExp += r.amount;
    });
    return monthly.map(m => ({ ...m, totalCost: m.itemCost + m.otherExp, net: m.revenue - m.itemCost - m.otherExp }));
  }

  function getFYears() {
    const fys = new Set();
    [...itemsIn, ...itemsOut, ...expenses].forEach(r => {
      const d = new Date(r.date);
      const m = d.getMonth();
      // If Jan-Mar, FY started previous year; if Apr-Dec, FY started same year
      fys.add(m >= 3 ? d.getFullYear() : d.getFullYear() - 1);
    });
    if (fys.size === 0) {
      const now = new Date();
      fys.add(now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1);
    }
    return [...fys].sort((a, b) => b - a);
  }

  function fyLabel(y) { return y + '–' + String(y + 1).slice(2); }

  // ── P&L Rendering ────────────────────────────────
  let pnlChart = null;
  let lastPnlData = [];

  function renderPnl() {
    const yearSel = $('#pnlYear');
    const years = getFYears();
    yearSel.innerHTML = years.map(y => `<option value="${y}">FY ${fyLabel(y)}</option>`).join('');

    function draw() {
      const year = parseInt(yearSel.value);
      const data = computeMonthlyPnl(year);
      lastPnlData = data.map((m, i) => ({
        month: MONTHS[i], revenue: m.revenue, itemCost: m.itemCost, otherExp: m.otherExp,
        totalCost: m.totalCost, net: m.net,
        margin: m.revenue > 0 ? +((m.net / m.revenue) * 100).toFixed(1) : 0
      }));
      const pnlCols = ['month','revenue','itemCost','otherExp','totalCost','net','margin'];
      let pnlRows = lastPnlData;
      const ps = sortState['tablePnl'];
      if (ps) {
        pnlRows = [...pnlRows].sort((a, b) => {
          const key = pnlCols[parseInt(ps.col)] || 'month';
          const va = a[key], vb = b[key];
          if (typeof va === 'number' && typeof vb === 'number') return ps.dir === 'asc' ? va - vb : vb - va;
          return ps.dir === 'asc' ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
        });
      }
      const totRev = data.reduce((s, m) => s + m.revenue, 0);
      const totCost = data.reduce((s, m) => s + m.totalCost, 0);
      const totNet = totRev - totCost;

      $('#pnlRevenue').textContent = fmt(totRev);
      $('#pnlExpenses').textContent = fmt(totCost);
      const netEl = $('#pnlNet');
      netEl.textContent = fmt(totNet);
      netEl.className = 'stat-value ' + (totNet >= 0 ? 'profit' : 'loss');

      const tbody = $('#tablePnl tbody');
      tbody.innerHTML = pnlRows.map(r => {
        const marginStr = r.margin ? r.margin + '%' : '—';
        return `<tr>
          <td>${r.month}</td>
          <td>${fmt(r.revenue)}</td><td>${fmt(r.itemCost)}</td>
          <td>${fmt(r.otherExp)}</td><td>${fmt(r.totalCost)}</td>
          <td class="${r.net >= 0 ? 'profit' : 'loss'}">${fmt(r.net)}</td>
          <td>${marginStr}</td>
        </tr>`;
      }).join('');

      const totalMargin = totRev > 0 ? ((totNet / totRev) * 100).toFixed(1) : '—';
      const totItemCost = data.reduce((s, m) => s + m.itemCost, 0);
      const totOtherExp = data.reduce((s, m) => s + m.otherExp, 0);
      const row = $('#pnlTotalRow');
      row.innerHTML = `
        <td><strong>TOTAL</strong></td>
        <td><strong>${fmt(totRev)}</strong></td><td><strong>${fmt(totItemCost)}</strong></td>
        <td><strong>${fmt(totOtherExp)}</strong></td><td><strong>${fmt(totCost)}</strong></td>
        <td class="${totNet >= 0 ? 'profit' : 'loss'}"><strong>${fmt(totNet)}</strong></td>
        <td><strong>${totalMargin}${totalMargin !== '—' ? '%' : ''}</strong></td>`;

      if (pnlChart) pnlChart.destroy();
      pnlChart = new Chart($('#chartPnlDetailed'), {
        type: 'bar',
        data: {
          labels: MONTHS,
          datasets: [
            { label: 'Revenue', data: data.map(m => m.revenue), backgroundColor: '#22c55e', borderRadius: 6, order: 2 },
            { label: 'Total Costs', data: data.map(m => m.totalCost), backgroundColor: '#ef4444', borderRadius: 6, order: 2 },
            { label: 'Net P&L', data: data.map(m => m.net), type: 'line', borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,.1)', fill: true, tension: .4, pointRadius: 5, pointBackgroundColor: '#6366f1', order: 1 }
          ]
        },
        options: {
          responsive: true,
          interaction: { mode: 'index', intersect: false },
          plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + fmt(ctx.raw) } } },
          scales: { y: { beginAtZero: true, ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } }
        }
      });
    }

    yearSel.addEventListener('change', draw);
    draw();
  }

  // ── Dashboard ─────────────────────────────────────
  let dashChart1 = null, dashChart2 = null;

  function refreshDashboard() {
    const totalRevenue = itemsOut.reduce((s, r) => s + r.amount, 0);
    const totalItemCost = itemsIn.reduce((s, r) => s + r.cost, 0);
    const totalExpenses = expenses.reduce((s, r) => s + r.amount, 0);
    const totalCost = totalItemCost + totalExpenses;
    const netProfit = totalRevenue - totalCost;

    $('#statRevenue').textContent = fmt(totalRevenue);
    $('#statCost').textContent = fmt(totalCost);
    const profitEl = $('#statProfit');
    profitEl.textContent = fmt(netProfit);
    profitEl.className = 'stat-value ' + (netProfit >= 0 ? 'profit' : 'loss');
    $('#statInventory').textContent = buildInventory().length;

    const now = new Date();
    const currentFY = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
    const monthly = computeMonthlyPnl(currentFY);

    if (dashChart1) dashChart1.destroy();
    dashChart1 = new Chart($('#chartMonthlyPnl'), {
      type: 'bar',
      data: {
        labels: MONTHS,
        datasets: [
          { label: 'Revenue', data: monthly.map(m => m.revenue), backgroundColor: '#22c55e', borderRadius: 4 },
          { label: 'Costs', data: monthly.map(m => m.totalCost), backgroundColor: '#ef4444', borderRadius: 4 },
          { label: 'Net', data: monthly.map(m => m.net), type: 'line', borderColor: '#6366f1', tension: .4, pointRadius: 4, pointBackgroundColor: '#6366f1', fill: false }
        ]
      },
      options: {
        responsive: true,
        plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + fmt(ctx.raw) } } },
        scales: { y: { ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } }
      }
    });

    if (dashChart2) dashChart2.destroy();
    dashChart2 = new Chart($('#chartRevenueCost'), {
      type: 'doughnut',
      data: {
        labels: ['Revenue', 'Item Costs', 'Other Expenses'],
        datasets: [{ data: [totalRevenue, totalItemCost, totalExpenses], backgroundColor: ['#22c55e', '#ef4444', '#f59e0b'], borderWidth: 0, spacing: 4, borderRadius: 6 }]
      },
      options: {
        responsive: true,
        cutout: '65%',
        plugins: {
          legend: { position: 'bottom' },
          tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } }
        }
      }
    });

    const recent = [
      ...itemsIn.slice(0, 10).map(r => ({ date: r.date, type: 'Purchase', item: r.item, qty: r.qty, amount: -r.cost })),
      ...itemsOut.slice(0, 10).map(r => ({ date: r.date, type: 'Sale', item: r.item, qty: r.qty, amount: r.amount })),
      ...expenses.slice(0, 10).map(r => ({ date: r.date, type: 'Expense', item: r.category, qty: '—', amount: -r.amount }))
    ].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 15);

    const tbody = $('#recentTransactions tbody');
    tbody.innerHTML = recent.length === 0
      ? '<tr><td colspan="5" style="text-align:center;color:var(--text-light);padding:32px">No transactions yet. Start adding items!</td></tr>'
      : recent.map(r => `<tr>
          <td>${sanitize(r.date)}</td>
          <td><span class="status-badge ${r.type === 'Sale' ? 'status-ok' : r.type === 'Purchase' ? 'status-low' : 'status-out'}">${r.type}</span></td>
          <td>${sanitize(r.item)}</td><td>${r.qty}</td>
          <td class="${r.amount >= 0 ? 'profit' : 'loss'}">${fmt(Math.abs(r.amount))}</td>
        </tr>`).join('');
  }

  // ── Clear Data ────────────────────────────────────
  $('#clearDataBtn').addEventListener('click', async () => {
    if (!confirm('Are you sure? This will delete ALL canteen data permanently from the cloud.')) return;
    try {
      const batch = db.batch();
      const [snap1, snap2, snap3, snap4] = await Promise.all([
        colItemsIn.get(), colItemsOut.get(), colExpenses.get(), colSales.get()
      ]);
      snap1.docs.forEach(d => batch.delete(d.ref));
      snap2.docs.forEach(d => batch.delete(d.ref));
      snap3.docs.forEach(d => batch.delete(d.ref));
      snap4.docs.forEach(d => batch.delete(d.ref));
      await batch.commit();
      toast('All data cleared');
    } catch {
      toast('Failed to clear data', true);
    }
  });

  // ── Auth & Init ───────────────────────────────────
  const loginScreen = $('#loginScreen');
  const loginForm = $('#loginForm');
  const loginError = $('#loginError');
  const loginBtn = $('#loginBtn');

  function showApp(user) {
    const email = user.email;
    loginScreen.classList.add('hidden');
    $('#loadingOverlay').classList.add('hidden');
    $('#sidebarUser').textContent = email;
    $('#userBadge').textContent = email;
    startListeners();
    navigate('dashboard');
  }

  function showLogin() {
    loginScreen.classList.remove('hidden');
    $('#loadingOverlay').classList.add('hidden');
  }

  // Check if already logged in
  auth.onAuthStateChanged((user) => {
    if (user) {
      showApp(user);
    } else {
      showLogin();
    }
  });

  // Login form submit
  loginForm.addEventListener('submit', (e) => {
    e.preventDefault();
    loginError.textContent = '';
    loginBtn.disabled = true;
    loginBtn.textContent = 'Signing in...';

    const email = $('#loginEmail').value.trim();
    const password = $('#loginPassword').value;

    auth.signInWithEmailAndPassword(email, password)
      .then((cred) => {
        logAuthEvent(cred.user.email, 'Login');
        showApp(cred.user);
      })
      .catch((err) => {
        loginBtn.disabled = false;
        loginBtn.textContent = 'Sign In';
        if (err.code === 'auth/user-not-found' || err.code === 'auth/wrong-password' || err.code === 'auth/invalid-credential') {
          loginError.textContent = 'Invalid email or password.';
        } else if (err.code === 'auth/too-many-requests') {
          loginError.textContent = 'Too many attempts. Try again later.';
        } else {
          loginError.textContent = 'Login failed. Please try again.';
        }
      });
  });

  // Logout
  $('#logoutBtn').addEventListener('click', () => {
    const email = auth.currentUser?.email;
    logAuthEvent(email, 'Logout').finally(() => {
      auth.signOut().then(() => {
        itemsIn = []; itemsOut = []; expenses = []; sales = [];
        showLogin();
        toast('Logged out');
      });
    });
  });

  // ── Auth Activity Log ─────────────────────────────
  let authLogs = [];

  function logAuthEvent(email, action) {
    return colAuthLogs.add({
      email: email || 'unknown',
      action: action,
      timestamp: firebase.firestore.FieldValue.serverTimestamp()
    }).catch(() => {});
  }

  colAuthLogs.orderBy('timestamp', 'desc').onSnapshot((snap) => {
    authLogs = snap.docs.map(d => {
      const data = d.data();
      return { id: d.id, ...data, ts: data.timestamp ? data.timestamp.toDate() : new Date() };
    });
    const activePage = document.querySelector('.nav-item.active')?.dataset.page;
    if (activePage === 'changelog') renderAuthLog();
  });

  function renderAuthLog() {
    const tbody = $('#tableAuthLog tbody');
    let list = applySort(authLogs, 'tableAuthLog', (r, col) => {
      if (col === 'timestamp') return r.ts.getTime();
      return r[col] || '';
    });
    lastFilteredAuthLog = list;
    tbody.innerHTML = list.length === 0
      ? '<tr><td colspan="3" style="text-align:center;color:var(--text-light);padding:32px">No activity recorded yet</td></tr>'
      : list.map(r => {
          const dt = r.ts.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' })
                   + ' ' + r.ts.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit', hour12:true });
          const cls = r.action === 'Login' ? 'status-ok' : 'status-out';
          return `<tr>
            <td>${sanitize(dt)}</td>
            <td>${sanitize(r.email)}</td>
            <td><span class="status-badge ${cls}">${sanitize(r.action)}</span></td>
          </tr>`;
        }).join('');
  }
  let lastFilteredAuthLog = [];

  $('#btnExportAuthLog').addEventListener('click', () => {
    exportXlsx(
      lastFilteredAuthLog.map(r => {
        const dt = r.ts.toLocaleDateString('en-IN') + ' ' + r.ts.toLocaleTimeString('en-IN');
        return [dt, r.email, r.action];
      }),
      ['Date & Time','User','Action'],
      'canteen_auth_log'
    );
  });

  // ── P&L Export ────────────────────────────────────
  $('#btnExportPnl').addEventListener('click', () => {
    const fy = $('#pnlYear').value;
    exportXlsx(
      lastPnlData.map(r => [r.month, r.revenue, r.itemCost, r.otherExp, r.totalCost, r.net, r.margin ? r.margin + '%' : '—']),
      ['Month','Revenue','Item Costs','Other Expenses','Total Costs','Net P&L','Margin %'],
      'canteen_pnl_FY' + fyLabel(fy)
    );
  });

  // ── Month-wise Expense Export ─────────────────────
  $('#btnExportExpMonthly').addEventListener('click', () => {
    const fy = $('#expMonthFY').value;
    if (!lastExpMonthlyData.length) return toast('No data to export', true);
    exportXlsx(
      lastExpMonthlyData.map(r => [r.category, ...r.months, r.total]),
      ['Category', ...MONTHS, 'Total'],
      'canteen_monthly_expenses_FY' + fyLabel(fy)
    );
  });

  // ── Bind sort headers ─────────────────────────────
  bindSortHeaders('tableItemsIn', renderItemsIn);
  bindSortHeaders('tableItemsOut', renderItemsOut);
  bindSortHeaders('tableInventory', renderInventory);
  bindSortHeaders('tableExpenses', renderExpenses);
  bindSortHeaders('tableSales', renderSales);
  bindSortHeaders('tablePnl', renderPnl);
  bindSortHeaders('tableAuthLog', renderAuthLog);

})();
