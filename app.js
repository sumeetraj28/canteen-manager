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
  const titles = { dashboard:'Dashboard', 'items-in':'Items In', 'items-out':'Items Out', inventory:'Inventory', expenses:'Other Expenses', sales:'Sales', reports:'Reports', changelog:'Changelog' };

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
    if (page === 'reports')   renderReport();
    if (page === 'changelog') renderAuthLog();
    $('#sidebar').classList.remove('open');
    $('#sidebarOverlay').classList.remove('active');
  }

  navItems.forEach(n => n.addEventListener('click', (e) => { e.preventDefault(); navigate(n.dataset.page); }));
  $('#menuToggle').addEventListener('click', () => {
    $('#sidebar').classList.toggle('open');
    $('#sidebarOverlay').classList.toggle('active');
  });
  $('#sidebarOverlay').addEventListener('click', () => {
    $('#sidebar').classList.remove('open');
    $('#sidebarOverlay').classList.remove('active');
  });

  // Sidebar collapse/expand toggle (desktop/tablet)
  $('#sidebarCollapseBtn').addEventListener('click', () => {
    const sidebar = $('#sidebar');
    sidebar.classList.toggle('collapsed');
    localStorage.setItem('sidebarCollapsed', sidebar.classList.contains('collapsed') ? '1' : '');
  });
  // Restore collapsed state from localStorage
  if (localStorage.getItem('sidebarCollapsed') === '1') {
    $('#sidebar').classList.add('collapsed');
  }

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
    if (activePage === 'reports')   renderReport();
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

  // ── Editing state ─────────────────────────────────
  let editingInId = null, editingOutId = null, editingExpId = null, editingSaleId = null;

  function cancelEdit(type) {
    if (type === 'in')   { editingInId = null;   $('#formItemsIn').reset();  $('#inSubmitBtn').textContent = '+ Add Item In'; }
    if (type === 'out')  { editingOutId = null;   $('#formItemsOut').reset(); $('#outSubmitBtn').textContent = '+ Add Item Out'; }
    if (type === 'exp')  { editingExpId = null;   $('#formExpenses').reset(); $('#expSubmitBtn').textContent = '+ Add Expense'; }
    if (type === 'sale') { editingSaleId = null;  $('#formSales').reset();    $('#saleSubmitBtn').textContent = '+ Add Sale'; }
    $$('input[type="date"]').forEach(inp => inp.value = today());
  }

  // ── Edit handler (delegated) ──────────────────────
  document.addEventListener('click', (e) => {
    if (!e.target.classList.contains('btn-edit')) return;
    const id = e.target.dataset.id;
    const type = e.target.dataset.type;

    if (type === 'in') {
      const r = itemsIn.find(x => x.id === id);
      if (!r) return;
      editingInId = id;
      $('#inDate').value = r.date; $('#inBillNo').value = r.billNo || ''; $('#inItem').value = r.item;
      $('#inBrand').value = r.brand || ''; $('#inSupplier').value = r.supplier || '';
      $('#inQty').value = r.qty; $('#inUnit').value = r.unit; $('#inRate').value = r.rate || 0;
      $('#inPrice').value = r.cost; $('#inRemark').value = r.remark || '';
      $('#inSubmitBtn').textContent = '✏️ Update Item In';
      $('#page-items-in').scrollIntoView({ behavior: 'smooth' });
    }
    if (type === 'out') {
      const r = itemsOut.find(x => x.id === id);
      if (!r) return;
      editingOutId = id;
      $('#outDate').value = r.date; $('#outItem').value = r.item;
      $('#outBrand').value = r.brand || ''; $('#outSupplier').value = r.supplier || '';
      $('#outQty').value = r.qty; $('#outUnit').value = r.unit;
      $('#outRate').value = r.rate || 0; $('#outPrice').value = r.amount;
      $('#outCategory').value = r.category || 'Cooked'; $('#outPerson').value = r.person || 'Pradeep';
      $('#outCustomer').value = r.customer || '';
      $('#outSubmitBtn').textContent = '✏️ Update Item Out';
      $('#page-items-out').scrollIntoView({ behavior: 'smooth' });
    }
    if (type === 'exp') {
      const r = expenses.find(x => x.id === id);
      if (!r) return;
      editingExpId = id;
      $('#expDate').value = r.date; $('#expCategory').value = r.category;
      $('#expAmount').value = r.amount; $('#expNote').value = r.note || '';
      $('#expSubmitBtn').textContent = '✏️ Update Expense';
      $('#page-expenses').scrollIntoView({ behavior: 'smooth' });
    }
    if (type === 'sale') {
      const r = sales.find(x => x.id === id);
      if (!r) return;
      editingSaleId = id;
      $('#saleDate').value = r.date; $('#saleType').value = r.type;
      $('#saleAmount').value = r.amount; $('#saleDetails').value = r.details || '';
      $('#saleSubmitBtn').textContent = '✏️ Update Sale';
      $('#page-sales').scrollIntoView({ behavior: 'smooth' });
    }
  });

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
    if (editingInId) {
      colItemsIn.doc(editingInId).update(record).then(() => {
        logAuthEvent(auth.currentUser?.email, 'Edited Purchase #' + editingInId.slice(0, 6));
        cancelEdit('in');
        toast('Record updated');
      }).catch(() => toast('Failed to update', true));
    } else {
      colItemsIn.add(record).then(() => {
        e.target.reset();
        $$('input[type="date"]').forEach(inp => inp.value = today());
        toast('Item added to stock');
      }).catch(() => toast('Failed to save', true));
    }
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
          <td><button class="btn-edit" data-id="${sanitize(r.id)}" data-type="in">Edit</button> <button class="btn-delete" data-id="${sanitize(r.id)}" data-type="in">Delete</button></td>
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
    logAuthEvent(auth.currentUser?.email, 'Exported Purchase Summaries XLSX');
  });

  // ── Items Out ─────────────────────────────────────
  // Auto-calculate sale amount = qty × rate
  function autoCalcSaleAmt() {
    const qty = parseFloat($('#outQty').value) || 0;
    const rate = parseFloat($('#outRate').value) || 0;
    $('#outPrice').value = qty && rate ? (qty * rate).toFixed(2) : '';
  }
  $('#outQty').addEventListener('input', autoCalcSaleAmt);
  $('#outRate').addEventListener('input', autoCalcSaleAmt);

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
    if (editingOutId) {
      colItemsOut.doc(editingOutId).update(record).then(() => {
        logAuthEvent(auth.currentUser?.email, 'Edited Sale #' + editingOutId.slice(0, 6));
        cancelEdit('out');
        toast('Record updated');
      }).catch(() => toast('Failed to update', true));
    } else {
      colItemsOut.add(record).then(() => {
        e.target.reset();
        $$('input[type="date"]').forEach(inp => inp.value = today());
        toast('Sale recorded');
      }).catch(() => toast('Failed to save', true));
    }
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
          <td><button class="btn-edit" data-id="${sanitize(r.id)}" data-type="out">Edit</button> <button class="btn-delete" data-id="${sanitize(r.id)}" data-type="out">Delete</button></td>
        </tr>`).join('');
    renderItemsOutSummary();
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

  // ── Items Out: Month-wise Summary Tables ──────────
  function renderItemsOutSummary() {
    const yearSel = $('#itemsOutFY');
    const years = getFYears();
    const prev = yearSel.value;
    yearSel.innerHTML = years.map(y => `<option value="${y}">FY ${fyLabel(y)}</option>`).join('');
    if (prev && years.includes(parseInt(prev))) yearSel.value = prev;
    drawItemsOutSummary();
  }

  function drawItemsOutSummary() {
    const fyStart = parseInt($('#itemsOutFY').value);

    function fyIdx(dateStr) {
      const d = new Date(dateStr);
      const m = d.getMonth(), y = d.getFullYear();
      if (m >= 3) return y === fyStart ? m - 3 : -1;
      return y === fyStart + 1 ? m + 9 : -1;
    }

    const itemQtyMap = {}, itemAmtMap = {}, brandMap = {}, supplierMap = {};
    const categoryMap = {}, personCountMap = {}, personAmtMap = {};

    itemsOut.forEach(r => {
      const idx = fyIdx(r.date);
      if (idx < 0) return;
      const key = r.item.toLowerCase();
      const brandKey = (r.brand || 'Unbranded').trim();
      const suppKey = (r.supplier || 'Unknown').trim();
      const catKey = (r.category || 'Uncategorized').trim();
      const persKey = (r.person || 'Unknown').trim();

      // 1. Item Qty
      if (!itemQtyMap[key]) itemQtyMap[key] = { name: r.item, unit: r.unit, months: MONTHS.map(() => 0) };
      itemQtyMap[key].months[idx] += r.qty;

      // 2. Item Amt
      if (!itemAmtMap[key]) itemAmtMap[key] = { name: r.item, months: MONTHS.map(() => 0) };
      itemAmtMap[key].months[idx] += r.amount;

      // 3. Brand Amt
      const bk = brandKey.toLowerCase();
      if (!brandMap[bk]) brandMap[bk] = { name: brandKey, months: MONTHS.map(() => 0) };
      brandMap[bk].months[idx] += r.amount;

      // 4. Supplier Amt
      const sk = suppKey.toLowerCase();
      if (!supplierMap[sk]) supplierMap[sk] = { name: suppKey, months: MONTHS.map(() => 0) };
      supplierMap[sk].months[idx] += r.amount;

      // 5. Category Amt
      const ck = catKey.toLowerCase();
      if (!categoryMap[ck]) categoryMap[ck] = { name: catKey, months: MONTHS.map(() => 0) };
      categoryMap[ck].months[idx] += r.amount;

      // 6. Person Count
      const pk = persKey.toLowerCase();
      if (!personCountMap[pk]) personCountMap[pk] = { name: persKey, months: MONTHS.map(() => 0) };
      personCountMap[pk].months[idx] += 1;

      // 7. Person Amt
      if (!personAmtMap[pk]) personAmtMap[pk] = { name: persKey, months: MONTHS.map(() => 0) };
      personAmtMap[pk].months[idx] += r.amount;
    });

    lastOutSummary = {
      itemQty: Object.values(itemQtyMap), itemAmt: Object.values(itemAmtMap),
      brand: Object.values(brandMap), supplier: Object.values(supplierMap),
      category: Object.values(categoryMap),
      personCount: Object.values(personCountMap), personAmt: Object.values(personAmtMap)
    };

    function renderOutGroupTable(tbodyId, rows, valFmt, totalRowId) {
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

    const fmtQ = v => v % 1 === 0 ? String(v) : v.toFixed(2);
    const fmtC = v => String(v);

    renderOutGroupTable('tableOutQtyMonthly', Object.values(itemQtyMap), fmtQ, null);
    renderOutGroupTable('tableOutAmtMonthly', Object.values(itemAmtMap), fmt, '#outAmtTotalRow');
    renderOutGroupTable('tableOutBrandMonthly', Object.values(brandMap), fmt, '#outBrandTotalRow');
    renderOutGroupTable('tableOutSupplierMonthly', Object.values(supplierMap), fmt, '#outSupplierTotalRow');
    renderOutGroupTable('tableOutCategoryMonthly', Object.values(categoryMap), fmt, '#outCategoryTotalRow');
    renderOutGroupTable('tableOutPersonCountMonthly', Object.values(personCountMap), fmtC, '#outPersonCountTotalRow');
    renderOutGroupTable('tableOutPersonAmtMonthly', Object.values(personAmtMap), fmt, '#outPersonAmtTotalRow');
  }

  let lastOutSummary = { itemQty: [], itemAmt: [], brand: [], supplier: [], category: [], personCount: [], personAmt: [] };

  $('#itemsOutFY').addEventListener('change', drawItemsOutSummary);

  $('#btnExportItemsOutSummary').addEventListener('click', () => {
    const s = lastOutSummary;
    const fy = $('#itemsOutFY').value;
    const rows = [];

    rows.push(['Quantity of Items Sold — Month-wise']);
    rows.push(['Item', 'Unit', ...MONTHS, 'Total']);
    s.itemQty.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, r.unit, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);

    rows.push(['Amount of Items Sold — Month-wise']);
    rows.push(['Item', ...MONTHS, 'Total']);
    s.itemAmt.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);

    rows.push(['Amount of Items Sold from Brands — Month-wise']);
    rows.push(['Brand', ...MONTHS, 'Total']);
    s.brand.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);

    rows.push(['Amount of Items Sold from Suppliers — Month-wise']);
    rows.push(['Supplier', ...MONTHS, 'Total']);
    s.supplier.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);

    rows.push(['Amount of Items Sold per Category — Month-wise']);
    rows.push(['Category', ...MONTHS, 'Total']);
    s.category.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);

    rows.push(['Number of Times Each Person Takes Out Items — Month-wise']);
    rows.push(['Person', ...MONTHS, 'Total']);
    s.personCount.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    rows.push([]);

    rows.push(['Amount of Items Each Person Takes — Month-wise']);
    rows.push(['Person', ...MONTHS, 'Total']);
    s.personAmt.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });

    if (rows.length <= 14) return toast('No data to export', true);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sales Summaries');
    XLSX.writeFile(wb, 'canteen_sales_summaries_FY' + fyLabel(fy) + '_' + today() + '.xlsx');
    logAuthEvent(auth.currentUser?.email, 'Exported Sales Summaries XLSX');
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
    if (editingExpId) {
      colExpenses.doc(editingExpId).update(record).then(() => {
        logAuthEvent(auth.currentUser?.email, 'Edited Expense #' + editingExpId.slice(0, 6));
        cancelEdit('exp');
        toast('Record updated');
      }).catch(() => toast('Failed to update', true));
    } else {
      colExpenses.add(record).then(() => {
        e.target.reset();
        $$('input[type="date"]').forEach(inp => inp.value = today());
        toast('Expense recorded');
      }).catch(() => toast('Failed to save', true));
    }
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
          <td><button class="btn-edit" data-id="${sanitize(r.id)}" data-type="exp">Edit</button> <button class="btn-delete" data-id="${sanitize(r.id)}" data-type="exp">Delete</button></td>
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
    const typeLabels = { in: 'Purchase', out: 'Sale', exp: 'Expense', sale: 'Sale Record' };
    let promise;
    if (type === 'in')  promise = colItemsIn.doc(id).delete();
    if (type === 'out') promise = colItemsOut.doc(id).delete();
    if (type === 'exp') promise = colExpenses.doc(id).delete();
    if (type === 'sale') promise = colSales.doc(id).delete();
    if (promise) promise.then(() => {
      toast('Record deleted');
      logAuthEvent(auth.currentUser?.email, 'Deleted ' + (typeLabels[type] || type) + ' #' + id.slice(0, 6));
    }).catch(() => toast('Failed to delete', true));
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
    const allData = buildInventory();
    let data = allData.filter(i => {
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

    // Stats from unfiltered data
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
    logAuthEvent(auth.currentUser?.email, 'Exported Inventory XLSX');
  });

  $('#btnExportItemsIn').addEventListener('click', () => {
    exportXlsx(
      lastFilteredItemsIn.map(r => [r.date, r.billNo || '', r.item, r.brand || '', r.supplier || '', r.qty, r.unit, r.rate || 0, r.cost, r.remark || '']),
      ['Date','Bill No','Item','Brand','Supplier','Qty','Unit','Rate','Cost','Remark'],
      'canteen_items_in'
    );
    logAuthEvent(auth.currentUser?.email, 'Exported Items In XLSX');
  });

  $('#btnExportItemsOut').addEventListener('click', () => {
    exportXlsx(
      lastFilteredItemsOut.map(r => [r.date, r.item, r.brand || '', r.supplier || '', r.qty, r.unit, r.rate || 0, r.amount, r.category || '', r.person || '', r.customer || '']),
      ['Date','Item','Brand','Supplier','Qty','Unit','Rate','Amount','Category','Person','Remark'],
      'canteen_items_out'
    );
    logAuthEvent(auth.currentUser?.email, 'Exported Items Out XLSX');
  });

  $('#btnExportExpenses').addEventListener('click', () => {
    exportXlsx(
      lastFilteredExpenses.map(r => [r.date, r.category, r.amount, r.note || '']),
      ['Date','Category','Amount','Description'],
      'canteen_expenses'
    );
    logAuthEvent(auth.currentUser?.email, 'Exported Expenses XLSX');
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
    if (editingSaleId) {
      colSales.doc(editingSaleId).update(record).then(() => {
        logAuthEvent(auth.currentUser?.email, 'Edited Sale Record #' + editingSaleId.slice(0, 6));
        cancelEdit('sale');
        toast('Record updated');
      }).catch(() => toast('Failed to update', true));
    } else {
      colSales.add(record).then(() => {
        e.target.reset();
        $$('input[type="date"]').forEach(inp => inp.value = today());
        toast('Sale recorded');
      }).catch(() => toast('Failed to save', true));
    }
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
          <td><button class="btn-edit" data-id="${sanitize(r.id)}" data-type="sale">Edit</button> <button class="btn-delete" data-id="${sanitize(r.id)}" data-type="sale">Delete</button></td>
        </tr>`).join('');
    renderSalesSummary();
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
    logAuthEvent(auth.currentUser?.email, 'Exported Sales XLSX');
  });

  // ── Sales: Sale Type Month-wise Summary ───────────
  function renderSalesSummary() {
    const yearSel = $('#salesFY');
    const years = getFYears();
    const prev = yearSel.value;
    yearSel.innerHTML = years.map(y => `<option value="${y}">FY ${fyLabel(y)}</option>`).join('');
    if (prev && years.includes(parseInt(prev))) yearSel.value = prev;
    drawSalesSummary();
  }

  function drawSalesSummary() {
    const fyStart = parseInt($('#salesFY').value);

    function fyIdx(dateStr) {
      const d = new Date(dateStr);
      const m = d.getMonth(), y = d.getFullYear();
      if (m >= 3) return y === fyStart ? m - 3 : -1;
      return y === fyStart + 1 ? m + 9 : -1;
    }

    const typeMap = {};
    sales.forEach(r => {
      const idx = fyIdx(r.date);
      if (idx < 0) return;
      const key = (r.type || 'Unknown').trim();
      const lk = key.toLowerCase();
      if (!typeMap[lk]) typeMap[lk] = { name: key, months: MONTHS.map(() => 0) };
      typeMap[lk].months[idx] += r.amount;
    });

    lastSaleTypeSummary = Object.values(typeMap);

    const tbody = $('#tableSaleTypeMonthly tbody');
    const monthTotals = MONTHS.map(() => 0);
    let grandTotal = 0;

    tbody.innerHTML = lastSaleTypeSummary.length === 0
      ? '<tr><td colspan="14" style="text-align:center;color:var(--text-light);padding:32px">No data for this FY</td></tr>'
      : lastSaleTypeSummary.filter(r => r.months.some(v => v)).map(r => {
          const rowTotal = r.months.reduce((s, v) => s + v, 0);
          r.months.forEach((v, i) => { monthTotals[i] += v; });
          grandTotal += rowTotal;
          return `<tr>
            <td><strong>${sanitize(r.name)}</strong></td>
            ${r.months.map(v => `<td>${v ? fmt(v) : '\u2014'}</td>`).join('')}
            <td><strong>${fmt(rowTotal)}</strong></td>
          </tr>`;
        }).join('');

    if (grandTotal > 0) {
      $('#saleTypeTotalRow').innerHTML = `
        <td><strong>TOTAL</strong></td>
        ${monthTotals.map(v => `<td><strong>${v ? fmt(v) : '\u2014'}</strong></td>`).join('')}
        <td><strong>${fmt(grandTotal)}</strong></td>`;
    } else {
      $('#saleTypeTotalRow').innerHTML = '';
    }
  }

  let lastSaleTypeSummary = [];

  $('#salesFY').addEventListener('change', drawSalesSummary);

  $('#btnExportSaleTypeSummary').addEventListener('click', () => {
    const fy = $('#salesFY').value;
    const rows = [['Sale Type', ...MONTHS, 'Total']];
    lastSaleTypeSummary.filter(r => r.months.some(v => v)).forEach(r => {
      rows.push([r.name, ...r.months, r.months.reduce((a, b) => a + b, 0)]);
    });
    if (rows.length <= 1) return toast('No data to export', true);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sale Type Summary');
    XLSX.writeFile(wb, 'canteen_sale_type_summary_FY' + fyLabel(fy) + '_' + today() + '.xlsx');
    logAuthEvent(auth.currentUser?.email, 'Exported Sale Type Summary XLSX');
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

  // ── Dashboard ─────────────────────────────────────
  let dashCharts = [];
  let lastPnlData = [];

  function refreshDashboard() {
    // Destroy old charts
    dashCharts.forEach(c => c.destroy());
    dashCharts = [];

    // FY selector
    const yearSel = $('#dashFY');
    const years = getFYears();
    const prev = yearSel.value;
    yearSel.innerHTML = years.map(y => `<option value="${y}">FY ${fyLabel(y)}</option>`).join('');
    if (prev && years.includes(parseInt(prev))) yearSel.value = prev;

    drawDashboard();
  }

  function drawDashboard() {
    dashCharts.forEach(c => c.destroy());
    dashCharts = [];

    const fyStart = parseInt($('#dashFY').value);
    const monthly = computeMonthlyPnl(fyStart);

    function fyIndex(dateStr) {
      const d = new Date(dateStr);
      const m = d.getMonth(), y = d.getFullYear();
      if (m >= 3) return y === fyStart ? m - 3 : -1;
      return y === fyStart + 1 ? m + 9 : -1;
    }

    // Filter data for selected FY
    const fyItemsOut = itemsOut.filter(r => fyIndex(r.date) >= 0);
    const fyItemsIn = itemsIn.filter(r => fyIndex(r.date) >= 0);
    const fyExpenses = expenses.filter(r => fyIndex(r.date) >= 0);
    const fySales = sales.filter(r => fyIndex(r.date) >= 0);

    const totalRevenue = fyItemsOut.reduce((s, r) => s + r.amount, 0);
    const totalItemCost = fyItemsIn.reduce((s, r) => s + r.cost, 0);
    const totalExpenses = fyExpenses.reduce((s, r) => s + r.amount, 0);
    const totalCost = totalItemCost + totalExpenses;
    const netProfit = totalRevenue - totalCost;
    const margin = totalRevenue > 0 ? ((netProfit / totalRevenue) * 100).toFixed(1) : 0;
    const txnCount = fyItemsIn.length + fyItemsOut.length + fyExpenses.length + fySales.length;

    // Active days (unique dates with transactions)
    const activeDays = new Set([
      ...fyItemsOut.map(r => r.date),
      ...fyItemsIn.map(r => r.date),
      ...fyExpenses.map(r => r.date)
    ]).size || 1;

    // KPI Row 1
    $('#statRevenue').textContent = fmt(totalRevenue);
    $('#statCost').textContent = fmt(totalCost);
    const profitEl = $('#statProfit');
    profitEl.textContent = fmt(netProfit);
    profitEl.className = 'stat-value ' + (netProfit >= 0 ? 'profit' : 'loss');
    const marginEl = $('#statMargin');
    marginEl.textContent = totalRevenue > 0 ? margin + '%' : '—';
    marginEl.className = 'stat-value ' + (netProfit >= 0 ? 'profit' : 'loss');
    $('#statInventory').textContent = buildInventory().length;
    $('#statTxnCount').textContent = txnCount;

    // KPI Row 2
    $('#statAvgRevenue').textContent = fmt(totalRevenue / activeDays);
    $('#statAvgCost').textContent = fmt(totalCost / activeDays);
    $('#statItemCostRatio').textContent = totalRevenue > 0 ? ((totalItemCost / totalRevenue) * 100).toFixed(1) + '%' : '—';
    $('#statExpRatio').textContent = totalRevenue > 0 ? ((totalExpenses / totalRevenue) * 100).toFixed(1) + '%' : '—';

    // Chart 1: Month-wise P&L
    dashCharts.push(new Chart($('#chartMonthlyPnl'), {
      type: 'bar',
      data: {
        labels: MONTHS,
        datasets: [
          { label: 'Revenue', data: monthly.map(m => m.revenue), backgroundColor: '#22c55e', borderRadius: 4 },
          { label: 'Costs', data: monthly.map(m => m.totalCost), backgroundColor: '#ef4444', borderRadius: 4 },
          { label: 'Net P&L', data: monthly.map(m => m.net), type: 'line', borderColor: '#6366f1', tension: .4, pointRadius: 4, pointBackgroundColor: '#6366f1', fill: false }
        ]
      },
      options: {
        responsive: true,
        plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + fmt(ctx.raw) } } },
        scales: { y: { ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } }
      }
    }));

    // Chart 2: Revenue vs Cost doughnut
    dashCharts.push(new Chart($('#chartRevenueCost'), {
      type: 'doughnut',
      data: {
        labels: ['Item Costs', 'Other Expenses', 'Profit'],
        datasets: [{ data: [totalItemCost, totalExpenses, Math.max(0, netProfit)], backgroundColor: ['#ef4444', '#f59e0b', '#22c55e'], borderWidth: 0, spacing: 4, borderRadius: 6 }]
      },
      options: {
        responsive: true, cutout: '65%',
        plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } } }
      }
    }));

    // Top 10 Items by Revenue
    const itemRevMap = {};
    fyItemsOut.forEach(r => {
      const k = r.item.toLowerCase();
      if (!itemRevMap[k]) itemRevMap[k] = { name: r.item, qty: 0, amt: 0 };
      itemRevMap[k].qty += r.qty; itemRevMap[k].amt += r.amount;
    });
    const topRev = Object.values(itemRevMap).sort((a, b) => b.amt - a.amt).slice(0, 10);

    // Chart 3: Top Items by Revenue (horizontal bar)
    dashCharts.push(new Chart($('#chartTopItems'), {
      type: 'bar',
      data: {
        labels: topRev.map(r => r.name),
        datasets: [{ label: 'Revenue', data: topRev.map(r => r.amt), backgroundColor: '#22c55e', borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y', responsive: true,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => fmt(ctx.raw) } } },
        scales: { x: { ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } }
      }
    }));

    // Top 10 Items by Cost
    const itemCostMap = {};
    fyItemsIn.forEach(r => {
      const k = r.item.toLowerCase();
      if (!itemCostMap[k]) itemCostMap[k] = { name: r.item, qty: 0, amt: 0 };
      itemCostMap[k].qty += r.qty; itemCostMap[k].amt += r.cost;
    });
    const topCost = Object.values(itemCostMap).sort((a, b) => b.amt - a.amt).slice(0, 10);

    // Chart 4: Top Items by Cost (horizontal bar)
    dashCharts.push(new Chart($('#chartTopCostItems'), {
      type: 'bar',
      data: {
        labels: topCost.map(r => r.name),
        datasets: [{ label: 'Cost', data: topCost.map(r => r.amt), backgroundColor: '#ef4444', borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y', responsive: true,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => fmt(ctx.raw) } } },
        scales: { x: { ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } }
      }
    }));

    // Chart 5: Category-wise sales
    const catMap = {};
    fyItemsOut.forEach(r => {
      const k = r.category || 'Uncategorized';
      catMap[k] = (catMap[k] || 0) + r.amount;
    });
    const catLabels = Object.keys(catMap), catValues = Object.values(catMap);
    const pieColors = ['#6366f1','#22c55e','#ef4444','#f59e0b','#3b82f6','#ec4899','#14b8a6','#a855f7','#f97316','#84cc16','#06b6d4','#e11d48','#8b5cf6','#10b981','#d946ef'];
    dashCharts.push(new Chart($('#chartCategorySales'), {
      type: 'doughnut',
      data: {
        labels: catLabels,
        datasets: [{ data: catValues, backgroundColor: pieColors.slice(0, catLabels.length), borderWidth: 0, spacing: 3 }]
      },
      options: { responsive: true, cutout: '55%', plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } } } }
    }));

    // Chart 6: Person-wise sales
    const persMap = {};
    fyItemsOut.forEach(r => {
      const k = r.person || 'Unknown';
      persMap[k] = (persMap[k] || 0) + r.amount;
    });
    dashCharts.push(new Chart($('#chartPersonSales'), {
      type: 'doughnut',
      data: {
        labels: Object.keys(persMap),
        datasets: [{ data: Object.values(persMap), backgroundColor: pieColors.slice(0, Object.keys(persMap).length), borderWidth: 0, spacing: 3 }]
      },
      options: { responsive: true, cutout: '55%', plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } } } }
    }));

    // Chart 7: Sale Type breakdown (from sales collection)
    const saleTypeMap = {};
    fySales.forEach(r => {
      const k = r.type || 'Unknown';
      saleTypeMap[k] = (saleTypeMap[k] || 0) + r.amount;
    });
    dashCharts.push(new Chart($('#chartSaleType'), {
      type: 'pie',
      data: {
        labels: Object.keys(saleTypeMap),
        datasets: [{ data: Object.values(saleTypeMap), backgroundColor: pieColors.slice(0, Object.keys(saleTypeMap).length), borderWidth: 0, spacing: 3 }]
      },
      options: { responsive: true, plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } } } }
    }));

    // Chart 8: Other Expenses by category
    const expCatMap = {};
    fyExpenses.forEach(r => {
      const k = r.category || 'Other';
      expCatMap[k] = (expCatMap[k] || 0) + r.amount;
    });
    dashCharts.push(new Chart($('#chartExpCategory'), {
      type: 'doughnut',
      data: {
        labels: Object.keys(expCatMap),
        datasets: [{ data: Object.values(expCatMap), backgroundColor: pieColors.slice(0, Object.keys(expCatMap).length), borderWidth: 0, spacing: 3 }]
      },
      options: { responsive: true, cutout: '55%', plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } } } }
    }));

    // Top 10 Selling Items table
    $('#dashTopSelling tbody').innerHTML = topRev.length === 0
      ? '<tr><td colspan="4" style="text-align:center;color:var(--text-light);padding:24px">No data</td></tr>'
      : topRev.map((r, i) => `<tr><td>${i + 1}</td><td>${sanitize(r.name)}</td><td>${r.qty % 1 === 0 ? r.qty : r.qty.toFixed(2)}</td><td>${fmt(r.amt)}</td></tr>`).join('');

    // Top 10 Purchased Items table
    $('#dashTopPurchased tbody').innerHTML = topCost.length === 0
      ? '<tr><td colspan="4" style="text-align:center;color:var(--text-light);padding:24px">No data</td></tr>'
      : topCost.map((r, i) => `<tr><td>${i + 1}</td><td>${sanitize(r.name)}</td><td>${r.qty % 1 === 0 ? r.qty : r.qty.toFixed(2)}</td><td>${fmt(r.amt)}</td></tr>`).join('');

    // Month-wise Financial Summary table
    let cumulative = 0;
    const tbody = $('#dashMonthlyTable tbody');
    tbody.innerHTML = monthly.map((m, i) => {
      cumulative += m.net;
      const mg = m.revenue > 0 ? ((m.net / m.revenue) * 100).toFixed(1) + '%' : '—';
      return `<tr>
        <td>${MONTHS[i]}</td><td>${fmt(m.revenue)}</td><td>${fmt(m.itemCost)}</td>
        <td>${fmt(m.otherExp)}</td><td>${fmt(m.totalCost)}</td>
        <td class="${m.net >= 0 ? 'profit' : 'loss'}">${fmt(m.net)}</td>
        <td>${mg}</td>
        <td class="${cumulative >= 0 ? 'profit' : 'loss'}">${fmt(cumulative)}</td>
      </tr>`;
    }).join('');

    const totItemC = monthly.reduce((s, m) => s + m.itemCost, 0);
    const totOthE = monthly.reduce((s, m) => s + m.otherExp, 0);
    const totalMg = totalRevenue > 0 ? ((netProfit / totalRevenue) * 100).toFixed(1) + '%' : '—';
    $('#dashMonthlyTotalRow').innerHTML = `
      <td><strong>TOTAL</strong></td><td><strong>${fmt(totalRevenue)}</strong></td>
      <td><strong>${fmt(totItemC)}</strong></td><td><strong>${fmt(totOthE)}</strong></td>
      <td><strong>${fmt(totalCost)}</strong></td>
      <td class="${netProfit >= 0 ? 'profit' : 'loss'}"><strong>${fmt(netProfit)}</strong></td>
      <td><strong>${totalMg}</strong></td>
      <td class="${cumulative >= 0 ? 'profit' : 'loss'}"><strong>${fmt(cumulative)}</strong></td>`;

    // ── P&L Section (merged) ────────────────────────
    // Cumulative P&L
    let pnlCumul = 0;
    lastPnlData = monthly.map((m, i) => {
      pnlCumul += m.net;
      return {
        month: MONTHS[i], revenue: m.revenue, itemCost: m.itemCost, otherExp: m.otherExp,
        totalCost: m.totalCost, net: m.net,
        margin: m.revenue > 0 ? +((m.net / m.revenue) * 100).toFixed(1) : 0,
        cumulative: pnlCumul
      };
    });

    const pnlCols = ['month','revenue','itemCost','otherExp','totalCost','net','margin','cumulative'];
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

    const pnlTotalMargin = totalRevenue > 0 ? ((netProfit / totalRevenue) * 100).toFixed(1) : '—';

    // Active months count
    const activeMonths = monthly.filter(m => m.revenue > 0 || m.totalCost > 0).length || 1;

    // Best & worst months
    const monthsWithData = monthly.map((m, i) => ({ ...m, label: MONTHS[i] })).filter(m => m.revenue > 0 || m.totalCost > 0);
    const bestMonth = monthsWithData.length ? monthsWithData.reduce((a, b) => a.net > b.net ? a : b) : null;
    const worstMonth = monthsWithData.length ? monthsWithData.reduce((a, b) => a.net < b.net ? a : b) : null;

    // Dashboard Best/Worst Month KPIs
    $('#statBestMonth').textContent = bestMonth ? bestMonth.label + ' (' + fmt(bestMonth.net) + ')' : '—';
    $('#statWorstMonth').textContent = worstMonth ? worstMonth.label + ' (' + fmt(worstMonth.net) + ')' : '—';

    // P&L Table
    const pnlTbody = $('#tablePnl tbody');
    pnlTbody.innerHTML = pnlRows.map(r => {
      const marginStr = r.margin ? r.margin + '%' : '—';
      return `<tr>
        <td>${r.month}</td>
        <td>${fmt(r.revenue)}</td><td>${fmt(r.itemCost)}</td>
        <td>${fmt(r.otherExp)}</td><td>${fmt(r.totalCost)}</td>
        <td class="${r.net >= 0 ? 'profit' : 'loss'}">${fmt(r.net)}</td>
        <td>${marginStr}</td>
        <td class="${r.cumulative >= 0 ? 'profit' : 'loss'}">${fmt(r.cumulative)}</td>
      </tr>`;
    }).join('');

    const finalCumul = lastPnlData.length ? lastPnlData[lastPnlData.length - 1].cumulative : 0;
    const pnlTotalRow = $('#pnlTotalRow');
    pnlTotalRow.innerHTML = `
      <td><strong>TOTAL</strong></td>
      <td><strong>${fmt(totalRevenue)}</strong></td><td><strong>${fmt(totalItemCost)}</strong></td>
      <td><strong>${fmt(totalExpenses)}</strong></td><td><strong>${fmt(totalCost)}</strong></td>
      <td class="${netProfit >= 0 ? 'profit' : 'loss'}"><strong>${fmt(netProfit)}</strong></td>
      <td><strong>${pnlTotalMargin !== '—' ? pnlTotalMargin + '%' : '—'}</strong></td>
      <td class="${finalCumul >= 0 ? 'profit' : 'loss'}"><strong>${fmt(finalCumul)}</strong></td>`;

    // P&L Chart 1: Month-wise P&L bar+line
    dashCharts.push(new Chart($('#chartPnlDetailed'), {
      type: 'bar',
      data: {
        labels: MONTHS,
        datasets: [
          { label: 'Revenue', data: monthly.map(m => m.revenue), backgroundColor: '#22c55e', borderRadius: 6, order: 2 },
          { label: 'Item Costs', data: monthly.map(m => m.itemCost), backgroundColor: '#ef4444', borderRadius: 6, order: 2 },
          { label: 'Other Expenses', data: monthly.map(m => m.otherExp), backgroundColor: '#f59e0b', borderRadius: 6, order: 2 },
          { label: 'Net P&L', data: monthly.map(m => m.net), type: 'line', borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,.1)', fill: true, tension: .4, pointRadius: 5, pointBackgroundColor: '#6366f1', order: 1 }
        ]
      },
      options: {
        responsive: true,
        interaction: { mode: 'index', intersect: false },
        plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + fmt(ctx.raw) } } },
        scales: { y: { beginAtZero: true, ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } }
      }
    }));

    // P&L Chart 2: Cost composition doughnut
    dashCharts.push(new Chart($('#chartPnlCostSplit'), {
      type: 'doughnut',
      data: {
        labels: ['Item Costs', 'Other Expenses'],
        datasets: [{ data: [totalItemCost, totalExpenses], backgroundColor: ['#ef4444', '#f59e0b'], borderWidth: 0, spacing: 4, borderRadius: 6 }]
      },
      options: { responsive: true, cutout: '60%', plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } } } }
    }));

    // P&L Chart 3: Cumulative P&L trend
    let pnlC = 0;
    const pnlCumulData = monthly.map(m => { pnlC += m.net; return pnlC; });
    dashCharts.push(new Chart($('#chartPnlCumulative'), {
      type: 'line',
      data: {
        labels: MONTHS,
        datasets: [{
          label: 'Cumulative P&L', data: pnlCumulData,
          borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,.08)', fill: true,
          tension: .4, pointRadius: 5, pointBackgroundColor: pnlCumulData.map(v => v >= 0 ? '#22c55e' : '#ef4444')
        }]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => 'Cumulative: ' + fmt(ctx.raw) } } },
        scales: { y: { ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } }
      }
    }));

    // P&L Chart 4: Margin trend
    const pnlMargins = monthly.map(m => m.revenue > 0 ? +((m.net / m.revenue) * 100).toFixed(1) : 0);
    dashCharts.push(new Chart($('#chartPnlMarginTrend'), {
      type: 'bar',
      data: {
        labels: MONTHS,
        datasets: [{
          label: 'Margin %', data: pnlMargins,
          backgroundColor: pnlMargins.map(v => v >= 0 ? '#22c55e' : '#ef4444'),
          borderRadius: 4
        }]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.raw + '%' } } },
        scales: { y: { ticks: { callback: v => v + '%' } } }
      }
    }));

    // Revenue Sources
    const revSrcMap = {};
    sales.forEach(r => {
      const idx = fyIndex(r.date);
      if (idx < 0) return;
      const k = (r.type || 'Unknown').trim();
      const lk = k.toLowerCase();
      if (!revSrcMap[lk]) revSrcMap[lk] = { name: k, months: MONTHS.map(() => 0) };
      revSrcMap[lk].months[idx] += r.amount;
    });
    if (Object.keys(revSrcMap).length === 0) {
      itemsOut.forEach(r => {
        const idx = fyIndex(r.date);
        if (idx < 0) return;
        const k = (r.category || 'Uncategorized').trim();
        const lk = k.toLowerCase();
        if (!revSrcMap[lk]) revSrcMap[lk] = { name: k, months: MONTHS.map(() => 0) };
        revSrcMap[lk].months[idx] += r.amount;
      });
    }

    const revSrcRows = Object.values(revSrcMap);
    const revGrand = revSrcRows.reduce((s, r) => s + r.months.reduce((a, b) => a + b, 0), 0);
    const revSrcTbody = $('#pnlRevenueSourcesTable tbody');
    const revMonthTotals = MONTHS.map(() => 0);
    revSrcTbody.innerHTML = revSrcRows.length === 0
      ? '<tr><td colspan="15" style="text-align:center;color:var(--text-light);padding:24px">No data</td></tr>'
      : revSrcRows.filter(r => r.months.some(v => v)).map(r => {
          const rowTotal = r.months.reduce((s, v) => s + v, 0);
          r.months.forEach((v, i) => { revMonthTotals[i] += v; });
          const pct = revGrand > 0 ? ((rowTotal / revGrand) * 100).toFixed(1) + '%' : '—';
          return `<tr><td><strong>${sanitize(r.name)}</strong></td>${r.months.map(v => `<td>${v ? fmt(v) : '—'}</td>`).join('')}<td><strong>${fmt(rowTotal)}</strong></td><td>${pct}</td></tr>`;
        }).join('');
    if (revGrand > 0) {
      $('#pnlRevSourceTotalRow').innerHTML = `<td><strong>TOTAL</strong></td>${revMonthTotals.map(v => `<td><strong>${v ? fmt(v) : '—'}</strong></td>`).join('')}<td><strong>${fmt(revGrand)}</strong></td><td><strong>100%</strong></td>`;
    } else { $('#pnlRevSourceTotalRow').innerHTML = ''; }

    // Expense Category Breakdown
    const pnlExpCatMap = {};
    expenses.forEach(r => {
      const idx = fyIndex(r.date);
      if (idx < 0) return;
      const k = (r.category || 'Other').trim();
      const lk = k.toLowerCase();
      if (!pnlExpCatMap[lk]) pnlExpCatMap[lk] = { name: k, months: MONTHS.map(() => 0) };
      pnlExpCatMap[lk].months[idx] += r.amount;
    });
    const expCatRows = Object.values(pnlExpCatMap);
    const expGrand = expCatRows.reduce((s, r) => s + r.months.reduce((a, b) => a + b, 0), 0);
    const expCatTbody = $('#pnlExpCategoryTable tbody');
    const expMonthTotals = MONTHS.map(() => 0);
    expCatTbody.innerHTML = expCatRows.length === 0
      ? '<tr><td colspan="15" style="text-align:center;color:var(--text-light);padding:24px">No data</td></tr>'
      : expCatRows.filter(r => r.months.some(v => v)).map(r => {
          const rowTotal = r.months.reduce((s, v) => s + v, 0);
          r.months.forEach((v, i) => { expMonthTotals[i] += v; });
          const pct = expGrand > 0 ? ((rowTotal / expGrand) * 100).toFixed(1) + '%' : '—';
          return `<tr><td><strong>${sanitize(r.name)}</strong></td>${r.months.map(v => `<td>${v ? fmt(v) : '—'}</td>`).join('')}<td><strong>${fmt(rowTotal)}</strong></td><td>${pct}</td></tr>`;
        }).join('');
    if (expGrand > 0) {
      $('#pnlExpCatTotalRow').innerHTML = `<td><strong>TOTAL</strong></td>${expMonthTotals.map(v => `<td><strong>${v ? fmt(v) : '—'}</strong></td>`).join('')}<td><strong>${fmt(expGrand)}</strong></td><td><strong>100%</strong></td>`;
    } else { $('#pnlExpCatTotalRow').innerHTML = ''; }

    // Recent Transactions
    const recent = [
      ...fyItemsIn.slice(0, 15).map(r => ({ date: r.date, type: 'Purchase', item: r.item, qty: r.qty, amount: -r.cost })),
      ...fyItemsOut.slice(0, 15).map(r => ({ date: r.date, type: 'Sale', item: r.item, qty: r.qty, amount: r.amount })),
      ...fyExpenses.slice(0, 15).map(r => ({ date: r.date, type: 'Expense', item: r.category, qty: '—', amount: -r.amount }))
    ].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 20);

    $('#recentTransactions tbody').innerHTML = recent.length === 0
      ? '<tr><td colspan="5" style="text-align:center;color:var(--text-light);padding:32px">No transactions yet</td></tr>'
      : recent.map(r => `<tr>
          <td>${sanitize(r.date)}</td>
          <td><span class="status-badge ${r.type === 'Sale' ? 'status-ok' : r.type === 'Purchase' ? 'status-low' : 'status-out'}">${r.type}</span></td>
          <td>${sanitize(r.item)}</td><td>${r.qty}</td>
          <td class="${r.amount >= 0 ? 'profit' : 'loss'}">${fmt(Math.abs(r.amount))}</td>
        </tr>`).join('');
  }

  $('#dashFY').addEventListener('change', drawDashboard);

  // ── Reports ─────────────────────────────────────
  let rptCharts = [];

  function getReportDateRange() {
    const period = $('#rptPeriod').value;
    const now = new Date();
    let from, to;

    switch (period) {
      case 'this-week': {
        const day = now.getDay() || 7; // Mon=1
        from = new Date(now); from.setDate(now.getDate() - day + 1); // Monday
        to = new Date(from); to.setDate(from.getDate() + 6);
        break;
      }
      case 'last-week': {
        const day = now.getDay() || 7;
        to = new Date(now); to.setDate(now.getDate() - day); // last Sunday
        from = new Date(to); from.setDate(to.getDate() - 6);
        break;
      }
      case 'this-month':
        from = new Date(now.getFullYear(), now.getMonth(), 1);
        to = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        break;
      case 'last-month':
        from = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        to = new Date(now.getFullYear(), now.getMonth(), 0);
        break;
      case 'this-fy': {
        const fy = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
        from = new Date(fy, 3, 1);
        to = new Date(fy + 1, 2, 31);
        break;
      }
      case 'last-fy': {
        const fy = (now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1) - 1;
        from = new Date(fy, 3, 1);
        to = new Date(fy + 1, 2, 31);
        break;
      }
      case 'custom':
        from = $('#rptFrom').value ? new Date($('#rptFrom').value) : new Date(now.getFullYear(), now.getMonth(), 1);
        to = $('#rptTo').value ? new Date($('#rptTo').value) : now;
        break;
    }
    // Normalize to YYYY-MM-DD strings
    const pad = d => d.toISOString().slice(0, 10);
    return { from: pad(from), to: pad(to) };
  }

  function formatRange(from, to) {
    const opts = { day: 'numeric', month: 'short', year: 'numeric' };
    return new Date(from).toLocaleDateString('en-IN', opts) + '  →  ' + new Date(to).toLocaleDateString('en-IN', opts);
  }

  // Toggle custom date pickers
  $('#rptPeriod').addEventListener('change', () => {
    const show = $('#rptPeriod').value === 'custom';
    $('#rptCustomFrom').style.display = show ? '' : 'none';
    $('#rptCustomTo').style.display = show ? '' : 'none';
  });

  function renderReport() {
    rptCharts.forEach(c => c.destroy());
    rptCharts = [];

    const { from, to } = getReportDateRange();
    $('#rptPeriodLabel').textContent = 'Report period: ' + formatRange(from, to);

    // Filter data within [from, to]
    const inRange = r => r.date >= from && r.date <= to;
    const fIn = itemsIn.filter(inRange);
    const fOut = itemsOut.filter(inRange);
    const fExp = expenses.filter(inRange);
    const fSales = sales.filter(inRange);

    const totalRev = fOut.reduce((s, r) => s + r.amount, 0);
    const totalPurchases = fIn.reduce((s, r) => s + r.cost, 0);
    const totalExp = fExp.reduce((s, r) => s + r.amount, 0);
    const totalCost = totalPurchases + totalExp;
    const net = totalRev - totalCost;
    const margin = totalRev > 0 ? ((net / totalRev) * 100).toFixed(1) : '—';
    const txnCount = fIn.length + fOut.length + fExp.length + fSales.length;
    const activeDays = new Set([...fIn.map(r => r.date), ...fOut.map(r => r.date), ...fExp.map(r => r.date)]).size || 1;

    // KPIs
    $('#rptRevenue').textContent = fmt(totalRev);
    $('#rptPurchases').textContent = fmt(totalPurchases);
    $('#rptExpenses').textContent = fmt(totalExp);
    $('#rptTotalCost').textContent = fmt(totalCost);
    const netEl = $('#rptNet');
    netEl.textContent = fmt(net);
    netEl.className = 'stat-value ' + (net >= 0 ? 'profit' : 'loss');
    const mgEl = $('#rptMargin');
    mgEl.textContent = margin !== '—' ? margin + '%' : '—';
    mgEl.className = 'stat-value ' + (net >= 0 ? 'profit' : 'loss');
    $('#rptTxns').textContent = txnCount;
    $('#rptDays').textContent = activeDays;
    $('#rptAvgRev').textContent = fmt(totalRev / activeDays);
    $('#rptAvgCost').textContent = fmt(totalCost / activeDays);

    // ─── Chart 1: Revenue vs Cost bar ───
    rptCharts.push(new Chart($('#chartRptRevCost'), {
      type: 'bar',
      data: {
        labels: ['Revenue', 'Purchases', 'Other Expenses', 'Net P&L'],
        datasets: [{
          data: [totalRev, totalPurchases, totalExp, net],
          backgroundColor: ['#22c55e', '#ef4444', '#f59e0b', net >= 0 ? '#6366f1' : '#dc2626'],
          borderRadius: 6
        }]
      },
      options: { responsive: true, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => fmt(ctx.raw) } } }, scales: { y: { ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } } }
    }));

    // ─── Chart 2: Cost Breakdown doughnut ───
    rptCharts.push(new Chart($('#chartRptCostBreak'), {
      type: 'doughnut',
      data: {
        labels: ['Item Purchases', 'Other Expenses'],
        datasets: [{ data: [totalPurchases, totalExp], backgroundColor: ['#ef4444', '#f59e0b'], borderWidth: 0, spacing: 4, borderRadius: 6 }]
      },
      options: { responsive: true, cutout: '60%', plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } } } }
    }));

    // ─── Chart 3: Daily Revenue Trend ───
    const dailyMap = {};
    fOut.forEach(r => { dailyMap[r.date] = (dailyMap[r.date] || 0) + r.amount; });
    const allDates = Object.keys(dailyMap).sort();
    rptCharts.push(new Chart($('#chartRptDailyRev'), {
      type: 'line',
      data: {
        labels: allDates.map(d => { const dt = new Date(d); return dt.getDate() + ' ' + dt.toLocaleString('en-IN', { month: 'short' }); }),
        datasets: [{
          label: 'Revenue', data: allDates.map(d => dailyMap[d]),
          borderColor: '#22c55e', backgroundColor: 'rgba(34,197,94,.08)', fill: true,
          tension: .4, pointRadius: 3, pointBackgroundColor: '#22c55e'
        }]
      },
      options: { responsive: true, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => fmt(ctx.raw) } } }, scales: { y: { ticks: { callback: v => '₹' + v.toLocaleString('en-IN') } } } }
    }));

    // ─── Chart 4: Expense Categories pie ───
    const expCatPie = {};
    fExp.forEach(r => { const k = r.category || 'Other'; expCatPie[k] = (expCatPie[k] || 0) + r.amount; });
    const pieColors = ['#6366f1','#22c55e','#ef4444','#f59e0b','#ec4899','#14b8a6','#8b5cf6','#f97316','#06b6d4','#84cc16'];
    const ecLabels = Object.keys(expCatPie);
    rptCharts.push(new Chart($('#chartRptExpCat'), {
      type: 'doughnut',
      data: {
        labels: ecLabels,
        datasets: [{ data: Object.values(expCatPie), backgroundColor: pieColors.slice(0, ecLabels.length), borderWidth: 0, spacing: 3 }]
      },
      options: { responsive: true, cutout: '55%', plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmt(ctx.raw) } } } }
    }));

    // ─── Purchase Summary by Item ───
    const purchMap = {};
    fIn.forEach(r => {
      const key = r.item.toLowerCase();
      if (!purchMap[key]) purchMap[key] = { name: r.item, brand: r.brand || '', supplier: r.supplier || '', qty: 0, unit: r.unit, cost: 0 };
      purchMap[key].qty += r.qty;
      purchMap[key].cost += r.cost;
      if (r.brand && !purchMap[key].brand) purchMap[key].brand = r.brand;
      if (r.supplier && !purchMap[key].supplier) purchMap[key].supplier = r.supplier;
    });
    const purchRows = Object.values(purchMap).sort((a, b) => b.cost - a.cost);
    const purchGrand = purchRows.reduce((s, r) => s + r.cost, 0);
    const purchQtyGrand = purchRows.reduce((s, r) => s + r.qty, 0);
    $('#rptPurchaseTable tbody').innerHTML = purchRows.length === 0
      ? '<tr><td colspan="7" style="text-align:center;color:var(--text-light);padding:24px">No purchases</td></tr>'
      : purchRows.map(r => `<tr><td>${sanitize(r.name)}</td><td>${sanitize(r.brand)}</td><td>${sanitize(r.supplier)}</td><td>${r.qty % 1 === 0 ? r.qty : r.qty.toFixed(2)}</td><td>${r.unit}</td><td>${fmt(r.cost)}</td><td>${r.qty > 0 ? fmt(r.cost / r.qty) : '—'}</td></tr>`).join('');
    $('#rptPurchaseTotalRow').innerHTML = purchRows.length
      ? `<td><strong>TOTAL</strong></td><td></td><td></td><td><strong>${purchQtyGrand % 1 === 0 ? purchQtyGrand : purchQtyGrand.toFixed(2)}</strong></td><td></td><td><strong>${fmt(purchGrand)}</strong></td><td></td>` : '';

    // ─── Purchase Summary by Supplier ───
    const suppMap = {};
    fIn.forEach(r => {
      const k = (r.supplier || 'Unknown').trim().toLowerCase();
      if (!suppMap[k]) suppMap[k] = { name: r.supplier || 'Unknown', items: new Set(), qty: 0, cost: 0 };
      suppMap[k].items.add(r.item.toLowerCase());
      suppMap[k].qty += r.qty;
      suppMap[k].cost += r.cost;
    });
    const suppRows = Object.values(suppMap).sort((a, b) => b.cost - a.cost);
    const suppGrand = suppRows.reduce((s, r) => s + r.cost, 0);
    $('#rptSupplierTable tbody').innerHTML = suppRows.length === 0
      ? '<tr><td colspan="5" style="text-align:center;color:var(--text-light);padding:24px">No data</td></tr>'
      : suppRows.map(r => {
          const pct = suppGrand > 0 ? ((r.cost / suppGrand) * 100).toFixed(1) + '%' : '—';
          return `<tr><td>${sanitize(r.name)}</td><td>${r.items.size}</td><td>${r.qty % 1 === 0 ? r.qty : r.qty.toFixed(2)}</td><td>${fmt(r.cost)}</td><td>${pct}</td></tr>`;
        }).join('');
    $('#rptSupplierTotalRow').innerHTML = suppGrand > 0
      ? `<td><strong>TOTAL</strong></td><td></td><td></td><td><strong>${fmt(suppGrand)}</strong></td><td><strong>100%</strong></td>` : '';

    // ─── Sales Summary by Item ───
    const saleItemMap = {};
    fOut.forEach(r => {
      const key = r.item.toLowerCase();
      if (!saleItemMap[key]) saleItemMap[key] = { name: r.item, category: r.category || '', qty: 0, unit: r.unit, amount: 0 };
      saleItemMap[key].qty += r.qty;
      saleItemMap[key].amount += r.amount;
      if (r.category && !saleItemMap[key].category) saleItemMap[key].category = r.category;
    });
    const saleRows = Object.values(saleItemMap).sort((a, b) => b.amount - a.amount);
    const saleGrand = saleRows.reduce((s, r) => s + r.amount, 0);
    const saleQtyGrand = saleRows.reduce((s, r) => s + r.qty, 0);
    $('#rptSalesTable tbody').innerHTML = saleRows.length === 0
      ? '<tr><td colspan="6" style="text-align:center;color:var(--text-light);padding:24px">No sales</td></tr>'
      : saleRows.map(r => `<tr><td>${sanitize(r.name)}</td><td>${sanitize(r.category)}</td><td>${r.qty % 1 === 0 ? r.qty : r.qty.toFixed(2)}</td><td>${r.unit}</td><td>${fmt(r.amount)}</td><td>${r.qty > 0 ? fmt(r.amount / r.qty) : '—'}</td></tr>`).join('');
    $('#rptSalesTotalRow').innerHTML = saleGrand > 0
      ? `<td><strong>TOTAL</strong></td><td></td><td><strong>${saleQtyGrand % 1 === 0 ? saleQtyGrand : saleQtyGrand.toFixed(2)}</strong></td><td></td><td><strong>${fmt(saleGrand)}</strong></td><td></td>` : '';

    // ─── Sales by Person ───
    const personMap = {};
    fOut.forEach(r => {
      const k = (r.person || 'Unknown').trim().toLowerCase();
      if (!personMap[k]) personMap[k] = { name: r.person || 'Unknown', count: 0, amount: 0 };
      personMap[k].count++;
      personMap[k].amount += r.amount;
    });
    const personRows = Object.values(personMap).sort((a, b) => b.amount - a.amount);
    const personGrand = personRows.reduce((s, r) => s + r.amount, 0);
    $('#rptPersonTable tbody').innerHTML = personRows.length === 0
      ? '<tr><td colspan="4" style="text-align:center;color:var(--text-light);padding:24px">No data</td></tr>'
      : personRows.map(r => {
          const pct = personGrand > 0 ? ((r.amount / personGrand) * 100).toFixed(1) + '%' : '—';
          return `<tr><td>${sanitize(r.name)}</td><td>${r.count}</td><td>${fmt(r.amount)}</td><td>${pct}</td></tr>`;
        }).join('');
    $('#rptPersonTotalRow').innerHTML = personGrand > 0
      ? `<td><strong>TOTAL</strong></td><td><strong>${personRows.reduce((s, r) => s + r.count, 0)}</strong></td><td><strong>${fmt(personGrand)}</strong></td><td><strong>100%</strong></td>` : '';

    // ─── Sales by Category ───
    const catMap = {};
    fOut.forEach(r => {
      const k = (r.category || 'Uncategorized').trim().toLowerCase();
      if (!catMap[k]) catMap[k] = { name: r.category || 'Uncategorized', count: 0, amount: 0 };
      catMap[k].count++;
      catMap[k].amount += r.amount;
    });
    const catRows = Object.values(catMap).sort((a, b) => b.amount - a.amount);
    const catGrand = catRows.reduce((s, r) => s + r.amount, 0);
    $('#rptCategoryTable tbody').innerHTML = catRows.length === 0
      ? '<tr><td colspan="4" style="text-align:center;color:var(--text-light);padding:24px">No data</td></tr>'
      : catRows.map(r => {
          const pct = catGrand > 0 ? ((r.amount / catGrand) * 100).toFixed(1) + '%' : '—';
          return `<tr><td>${sanitize(r.name)}</td><td>${r.count}</td><td>${fmt(r.amount)}</td><td>${pct}</td></tr>`;
        }).join('');
    $('#rptCategoryTotalRow').innerHTML = catGrand > 0
      ? `<td><strong>TOTAL</strong></td><td><strong>${catRows.reduce((s, r) => s + r.count, 0)}</strong></td><td><strong>${fmt(catGrand)}</strong></td><td><strong>100%</strong></td>` : '';

    // ─── Expense Breakdown ───
    const expMap = {};
    fExp.forEach(r => {
      const k = (r.category || 'Other').trim().toLowerCase();
      if (!expMap[k]) expMap[k] = { name: r.category || 'Other', count: 0, amount: 0 };
      expMap[k].count++;
      expMap[k].amount += r.amount;
    });
    const expRows = Object.values(expMap).sort((a, b) => b.amount - a.amount);
    const expGrandR = expRows.reduce((s, r) => s + r.amount, 0);
    $('#rptExpTable tbody').innerHTML = expRows.length === 0
      ? '<tr><td colspan="4" style="text-align:center;color:var(--text-light);padding:24px">No expenses</td></tr>'
      : expRows.map(r => {
          const pct = expGrandR > 0 ? ((r.amount / expGrandR) * 100).toFixed(1) + '%' : '—';
          return `<tr><td>${sanitize(r.name)}</td><td>${r.count}</td><td>${fmt(r.amount)}</td><td>${pct}</td></tr>`;
        }).join('');
    $('#rptExpTotalRow').innerHTML = expGrandR > 0
      ? `<td><strong>TOTAL</strong></td><td><strong>${expRows.reduce((s, r) => s + r.count, 0)}</strong></td><td><strong>${fmt(expGrandR)}</strong></td><td><strong>100%</strong></td>` : '';

    // ─── Daily Breakdown ───
    const dayData = {};
    fOut.forEach(r => {
      if (!dayData[r.date]) dayData[r.date] = { revenue: 0, purchases: 0, expenses: 0 };
      dayData[r.date].revenue += r.amount;
    });
    fIn.forEach(r => {
      if (!dayData[r.date]) dayData[r.date] = { revenue: 0, purchases: 0, expenses: 0 };
      dayData[r.date].purchases += r.cost;
    });
    fExp.forEach(r => {
      if (!dayData[r.date]) dayData[r.date] = { revenue: 0, purchases: 0, expenses: 0 };
      dayData[r.date].expenses += r.amount;
    });
    const dayRows = Object.entries(dayData).sort(([a], [b]) => a.localeCompare(b)).map(([date, d]) => ({
      date, ...d, totalCost: d.purchases + d.expenses, net: d.revenue - d.purchases - d.expenses
    }));
    const dayTotals = dayRows.reduce((t, r) => ({ revenue: t.revenue + r.revenue, purchases: t.purchases + r.purchases, expenses: t.expenses + r.expenses, totalCost: t.totalCost + r.totalCost, net: t.net + r.net }), { revenue: 0, purchases: 0, expenses: 0, totalCost: 0, net: 0 });
    $('#rptDailyTable tbody').innerHTML = dayRows.length === 0
      ? '<tr><td colspan="6" style="text-align:center;color:var(--text-light);padding:24px">No data</td></tr>'
      : dayRows.map(r => `<tr><td>${r.date}</td><td>${fmt(r.revenue)}</td><td>${fmt(r.purchases)}</td><td>${fmt(r.expenses)}</td><td>${fmt(r.totalCost)}</td><td class="${r.net >= 0 ? 'profit' : 'loss'}">${fmt(r.net)}</td></tr>`).join('');
    $('#rptDailyTotalRow').innerHTML = dayRows.length
      ? `<td><strong>TOTAL</strong></td><td><strong>${fmt(dayTotals.revenue)}</strong></td><td><strong>${fmt(dayTotals.purchases)}</strong></td><td><strong>${fmt(dayTotals.expenses)}</strong></td><td><strong>${fmt(dayTotals.totalCost)}</strong></td><td class="${dayTotals.net >= 0 ? 'profit' : 'loss'}"><strong>${fmt(dayTotals.net)}</strong></td>` : '';

    // Show report
    $('#reportOutput').style.display = '';
  }

  $('#btnGenReport').addEventListener('click', renderReport);

  // Export All Report tables to XLSX (multi-sheet)
  $('#btnExportReport').addEventListener('click', () => {
    if ($('#reportOutput').style.display === 'none') return toast('Generate a report first', true);
    const { from, to } = getReportDateRange();
    const wb = XLSX.utils.book_new();
    const addSheet = (tableId, name) => {
      const tbl = document.getElementById(tableId);
      if (tbl) { const ws = XLSX.utils.table_to_sheet(tbl); XLSX.utils.book_append_sheet(wb, ws, name); }
    };
    addSheet('rptPurchaseTable', 'Purchases by Item');
    addSheet('rptSupplierTable', 'Purchases by Supplier');
    addSheet('rptSalesTable', 'Sales by Item');
    addSheet('rptPersonTable', 'Sales by Person');
    addSheet('rptCategoryTable', 'Sales by Category');
    addSheet('rptExpTable', 'Expenses');
    addSheet('rptDailyTable', 'Daily Breakdown');
    XLSX.writeFile(wb, 'canteen_report_' + from + '_to_' + to + '.xlsx');
    toast('Report exported to .xlsx');
    logAuthEvent(auth.currentUser?.email, 'Exported Report XLSX');
  });

  // Export Report as PDF
  $('#btnExportPdf').addEventListener('click', async () => {
    const output = $('#reportOutput');
    if (output.style.display === 'none') return toast('Generate a report first', true);
    const pdfBtn = $('#btnExportPdf');
    pdfBtn.disabled = true;
    pdfBtn.textContent = '⏳ Generating…';
    toast('Generating PDF…');
    try {
      const canvas = await html2canvas(output, { scale: 2, useCORS: true, logging: false });
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jspdf.jsPDF('p', 'mm', 'a4');
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 10;
      const usableW = pageW - margin * 2;
      const imgW = usableW;
      const imgH = (canvas.height * imgW) / canvas.width;
      let y = margin;
      let remaining = imgH;
      const usableH = pageH - margin * 2;
      while (remaining > 0) {
        if (y !== margin) pdf.addPage();
        const srcY = (imgH - remaining) / imgH * canvas.height;
        const sliceH = Math.min(usableH, remaining);
        const srcSliceH = (sliceH / imgH) * canvas.height;
        const sliceCanvas = document.createElement('canvas');
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = srcSliceH;
        sliceCanvas.getContext('2d').drawImage(canvas, 0, srcY, canvas.width, srcSliceH, 0, 0, canvas.width, srcSliceH);
        pdf.addImage(sliceCanvas.toDataURL('image/png'), 'PNG', margin, margin, imgW, sliceH);
        remaining -= usableH;
      }
      const { from, to } = getReportDateRange();
      pdf.save('canteen_report_' + from + '_to_' + to + '.pdf');
      toast('PDF downloaded');
      logAuthEvent(auth.currentUser?.email, 'Downloaded Report PDF');
    } catch (err) {
      console.error(err);
      toast('PDF generation failed', true);
    } finally {
      pdfBtn.disabled = false;
      pdfBtn.textContent = '📄 Download PDF';
    }
  });

  // ── Clear Data ────────────────────────────────────
  $('#clearDataBtn').addEventListener('click', async () => {
    if (!confirm('Are you sure? This will delete ALL canteen data permanently from the cloud.')) return;
    const password = prompt('Re-enter your password to confirm:');
    if (!password) return;
    const user = auth.currentUser;
    if (!user) return toast('Not logged in', true);
    try {
      const credential = firebase.auth.EmailAuthProvider.credential(user.email, password);
      await user.reauthenticateWithCredential(credential);
    } catch {
      return toast('Incorrect password. Reset cancelled.', true);
    }
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

  function classifyAction(action) {
    if (action === 'Login' || action === 'Logout') return 'login';
    if (action.startsWith('Edited')) return 'edit';
    if (action.startsWith('Deleted')) return 'delete';
    if (action.startsWith('Exported')) return 'export';
    if (action.startsWith('Downloaded')) return 'download';
    return 'other';
  }

  const categoryLabels = { login: '🔑 Login/Logout', edit: '✏️ Edit', delete: '🗑️ Delete', export: '📤 Export', download: '📄 Download', other: '❓ Other' };
  const categoryBadgeClass = { login: 'status-ok', edit: 'badge-edit', delete: 'status-out', export: 'badge-export', download: 'badge-download', other: 'status-low' };

  function renderAuthLog() {
    // Classify and count
    const counts = { total: authLogs.length, login: 0, edit: 0, delete: 0, export: 0, download: 0 };
    authLogs.forEach(r => { const c = classifyAction(r.action); if (counts[c] !== undefined) counts[c]++; });
    const s = id => document.getElementById(id);
    s('statTotalActions').textContent = counts.total;
    s('statLogins').textContent = counts.login;
    s('statEdits').textContent = counts.edit;
    s('statDeletes').textContent = counts.delete;
    s('statExports').textContent = counts.export;
    s('statDownloads').textContent = counts.download;

    // Apply filters
    const catFilter = $('#filterLogCategory').value;
    const userFilter = $('#filterLogUser').value.toLowerCase().trim();
    const fromFilter = $('#filterLogFrom').value;
    const toFilter = $('#filterLogTo').value;

    let filtered = authLogs.filter(r => {
      if (catFilter && classifyAction(r.action) !== catFilter) return false;
      if (userFilter && !r.email.toLowerCase().includes(userFilter)) return false;
      if (fromFilter) {
        const d = r.ts.toISOString().slice(0, 10);
        if (d < fromFilter) return false;
      }
      if (toFilter) {
        const d = r.ts.toISOString().slice(0, 10);
        if (d > toFilter) return false;
      }
      return true;
    });

    const tbody = $('#tableAuthLog tbody');
    let list = applySort(filtered, 'tableAuthLog', (r, col) => {
      if (col === 'timestamp') return r.ts.getTime();
      return r[col] || '';
    });
    lastFilteredAuthLog = list;
    tbody.innerHTML = list.length === 0
      ? '<tr><td colspan="4" style="text-align:center;color:var(--text-light);padding:32px">No activity recorded yet</td></tr>'
      : list.map(r => {
          const dt = r.ts.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' })
                   + ' ' + r.ts.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit', hour12:true });
          const cat = classifyAction(r.action);
          const cls = categoryBadgeClass[cat] || 'status-low';
          return `<tr>
            <td>${sanitize(dt)}</td>
            <td>${sanitize(r.email)}</td>
            <td>${sanitize(r.action)}</td>
            <td><span class="status-badge ${cls}">${categoryLabels[cat] || cat}</span></td>
          </tr>`;
        }).join('');
  }
  let lastFilteredAuthLog = [];

  // Filter listeners for activity log
  ['filterLogCategory','filterLogUser','filterLogFrom','filterLogTo'].forEach(id => {
    $('#' + id).addEventListener(id === 'filterLogCategory' ? 'change' : 'input', () => renderAuthLog());
  });
  $('#clearFiltersLog').addEventListener('click', () => {
    $('#filterLogCategory').value = '';
    $('#filterLogUser').value = '';
    $('#filterLogFrom').value = '';
    $('#filterLogTo').value = '';
    renderAuthLog();
  });

  $('#btnExportAuthLog').addEventListener('click', () => {
    exportXlsx(
      lastFilteredAuthLog.map(r => {
        const dt = r.ts.toLocaleDateString('en-IN') + ' ' + r.ts.toLocaleTimeString('en-IN');
        return [dt, r.email, r.action];
      }),
      ['Date & Time','User','Action'],
      'canteen_auth_log'
    );
    logAuthEvent(auth.currentUser?.email, 'Exported Activity Log XLSX');
  });

  // ── P&L Export ────────────────────────────────────
  $('#btnExportPnl').addEventListener('click', () => {
    const fy = $('#dashFY').value;
    exportXlsx(
      lastPnlData.map(r => [r.month, r.revenue, r.itemCost, r.otherExp, r.totalCost, r.net, r.margin ? r.margin + '%' : '—', r.cumulative]),
      ['Month','Revenue','Item Costs','Other Expenses','Total Costs','Net P&L','Margin %','Cumulative'],
      'canteen_pnl_FY' + fyLabel(fy)
    );
    logAuthEvent(auth.currentUser?.email, 'Exported P&L XLSX');
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
    logAuthEvent(auth.currentUser?.email, 'Exported Monthly Expenses XLSX');
  });

  // ── Bind sort headers ─────────────────────────────
  bindSortHeaders('tableItemsIn', renderItemsIn);
  bindSortHeaders('tableItemsOut', renderItemsOut);
  bindSortHeaders('tableInventory', renderInventory);
  bindSortHeaders('tableExpenses', renderExpenses);
  bindSortHeaders('tableSales', renderSales);
  bindSortHeaders('tablePnl', drawDashboard);
  bindSortHeaders('tableAuthLog', renderAuthLog);

})();
