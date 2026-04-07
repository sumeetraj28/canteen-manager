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

  // ── In-memory data (synced from Firestore) ────────
  let itemsIn = [];
  let itemsOut = [];
  let expenses = [];

  // ── Firestore references ──────────────────────────
  const colItemsIn  = db.collection('itemsIn');
  const colItemsOut = db.collection('itemsOut');
  const colExpenses = db.collection('expenses');

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
  const titles = { dashboard:'Dashboard', 'items-in':'Items In', 'items-out':'Items Out', inventory:'Inventory', expenses:'Other Expenses', pnl:'P&L Statement' };

  function navigate(page) {
    navItems.forEach(n => n.classList.toggle('active', n.dataset.page === page));
    pages.forEach(p => { p.classList.toggle('active', p.id === 'page-' + page); });
    $('#pageTitle').textContent = titles[page] || 'Dashboard';
    if (page === 'dashboard') refreshDashboard();
    if (page === 'items-in')  renderItemsIn();
    if (page === 'items-out') renderItemsOut();
    if (page === 'inventory') renderInventory();
    if (page === 'expenses')  renderExpenses();
    if (page === 'pnl')       renderPnl();
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
  }

  function refreshActivePage() {
    const activePage = document.querySelector('.nav-item.active')?.dataset.page;
    if (activePage === 'dashboard') refreshDashboard();
    if (activePage === 'items-in')  renderItemsIn();
    if (activePage === 'items-out') renderItemsOut();
    if (activePage === 'inventory') renderInventory();
    if (activePage === 'expenses')  renderExpenses();
    if (activePage === 'pnl')       renderPnl();
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
    const filtered = itemsIn.filter(r => {
      if (f.dateFrom && r.date < f.dateFrom) return false;
      if (f.dateTo && r.date > f.dateTo) return false;
      if (f.item && !r.item.toLowerCase().includes(f.item)) return false;
      if (f.brand && !(r.brand || '').toLowerCase().includes(f.brand)) return false;
      if (f.supplier && !(r.supplier || '').toLowerCase().includes(f.supplier)) return false;
      if (f.billNo && !(r.billNo || '').toLowerCase().includes(f.billNo)) return false;
      return true;
    });
    tbody.innerHTML = filtered.length === 0
      ? '<tr><td colspan="11" style="text-align:center;color:var(--text-light);padding:32px">No purchase records yet</td></tr>'
      : filtered.map(r => `<tr>
          <td>${sanitize(r.date)}</td><td>${sanitize(r.billNo || '—')}</td><td>${sanitize(r.item)}</td>
          <td>${sanitize(r.brand || '—')}</td><td>${r.qty}</td><td>${sanitize(r.unit)}</td>
          <td>${fmt(r.rate || 0)}</td><td>${fmt(r.cost)}</td><td>${sanitize(r.supplier || '—')}</td>
          <td>${sanitize(r.remark || '—')}</td>
          <td><button class="btn-delete" data-id="${sanitize(r.id)}" data-type="in">Delete</button></td>
        </tr>`).join('');
  }

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

  // ── Items Out ─────────────────────────────────────
  $('#formItemsOut').addEventListener('submit', (e) => {
    e.preventDefault();
    const record = {
      date: $('#outDate').value,
      item: $('#outItem').value.trim(),
      category: $('#outCategory').value,
      person: $('#outPerson').value,
      qty: parseFloat($('#outQty').value),
      unit: $('#outUnit').value,
      amount: parseFloat($('#outPrice').value),
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
    const filtered = itemsOut.filter(r => {
      if (f.dateFrom && r.date < f.dateFrom) return false;
      if (f.dateTo && r.date > f.dateTo) return false;
      if (f.item && !r.item.toLowerCase().includes(f.item)) return false;
      if (f.category && (r.category || '') !== f.category) return false;
      if (f.person && (r.person || '') !== f.person) return false;
      return true;
    });
    tbody.innerHTML = filtered.length === 0
      ? '<tr><td colspan="9" style="text-align:center;color:var(--text-light);padding:32px">No sales records yet</td></tr>'
      : filtered.map(r => `<tr>
          <td>${sanitize(r.date)}</td><td>${sanitize(r.item)}</td>
          <td>${sanitize(r.category || '—')}</td><td>${sanitize(r.person || '—')}</td>
          <td>${r.qty}</td><td>${sanitize(r.unit)}</td>
          <td>${fmt(r.amount)}</td><td>${sanitize(r.customer || '—')}</td>
          <td><button class="btn-delete" data-id="${sanitize(r.id)}" data-type="out">Delete</button></td>
        </tr>`).join('');
  }

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

  function renderExpenses(filter = '') {
    const tbody = $('#tableExpenses tbody');
    const filtered = expenses.filter(r =>
      r.category.toLowerCase().includes(filter.toLowerCase()) ||
      (r.note || '').toLowerCase().includes(filter.toLowerCase())
    );
    tbody.innerHTML = filtered.length === 0
      ? '<tr><td colspan="5" style="text-align:center;color:var(--text-light);padding:32px">No expenses recorded yet</td></tr>'
      : filtered.map(r => `<tr>
          <td>${sanitize(r.date)}</td><td>${sanitize(r.category)}</td><td>${fmt(r.amount)}</td>
          <td>${sanitize(r.note || '—')}</td>
          <td><button class="btn-delete" data-id="${sanitize(r.id)}" data-type="exp">Delete</button></td>
        </tr>`).join('');
  }

  $('#searchExpenses').addEventListener('input', (e) => renderExpenses(e.target.value));

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
    if (promise) promise.then(() => toast('Record deleted')).catch(() => toast('Failed to delete', true));
  });

  // ── Inventory ─────────────────────────────────────
  function buildInventory() {
    const inv = {};
    itemsIn.forEach(r => {
      const key = r.item.toLowerCase();
      if (!inv[key]) inv[key] = { name: r.item, unit: r.unit, qtyIn: 0, qtyOut: 0, totalCost: 0 };
      inv[key].qtyIn += r.qty;
      inv[key].totalCost += r.cost;
    });
    itemsOut.forEach(r => {
      const key = r.item.toLowerCase();
      if (!inv[key]) inv[key] = { name: r.item, unit: r.unit, qtyIn: 0, qtyOut: 0, totalCost: 0 };
      inv[key].qtyOut += r.qty;
    });
    return Object.values(inv).map(i => {
      const balance = Math.max(0, i.qtyIn - i.qtyOut);
      const avgCost = i.qtyIn > 0 ? i.totalCost / i.qtyIn : 0;
      return { ...i, balance, avgCost, value: balance * avgCost };
    });
  }

  function renderInventory(filter = '') {
    const data = buildInventory().filter(i => i.name.toLowerCase().includes(filter.toLowerCase()));
    const tbody = $('#tableInventory tbody');
    let lowCount = 0;

    tbody.innerHTML = data.length === 0
      ? '<tr><td colspan="8" style="text-align:center;color:var(--text-light);padding:32px">No inventory data</td></tr>'
      : data.map(i => {
          let status, cls;
          if (i.balance <= 0) { status = 'Out'; cls = 'status-out'; }
          else if (i.balance < i.qtyIn * 0.2) { status = 'Low'; cls = 'status-low'; lowCount++; }
          else { status = 'OK'; cls = 'status-ok'; }
          return `<tr>
            <td><strong>${sanitize(i.name)}</strong></td>
            <td>${i.qtyIn.toFixed(2)}</td><td>${i.qtyOut.toFixed(2)}</td>
            <td><strong>${i.balance.toFixed(2)}</strong></td><td>${sanitize(i.unit)}</td>
            <td>${fmt(i.avgCost)}</td><td>${fmt(i.value)}</td>
            <td><span class="status-badge ${cls}">${status}</span></td>
          </tr>`;
        }).join('');

    $('#invTotalItems').textContent = data.length;
    $('#invLowStock').textContent = lowCount;
    const totalVal = data.reduce((s, i) => s + i.value, 0);
    $('#invTotalValue').textContent = fmt(totalVal);
  }

  $('#searchInventory').addEventListener('input', (e) => renderInventory(e.target.value));

  // ── Export Inventory CSV ──────────────────────────
  $('#btnExportInventory').addEventListener('click', () => {
    const data = buildInventory();
    if (data.length === 0) return toast('No data to export', true);
    let csv = 'Item,Qty In,Qty Out,Balance,Unit,Avg Cost,Stock Value\n';
    data.forEach(i => {
      csv += `"${i.name}",${i.qtyIn},${i.qtyOut},${i.balance},"${i.unit}",${i.avgCost.toFixed(2)},${i.value.toFixed(2)}\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'canteen_inventory_' + today() + '.csv'; a.click();
    URL.revokeObjectURL(url);
    toast('Inventory exported');
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

  function renderPnl() {
    const yearSel = $('#pnlYear');
    const years = getFYears();
    yearSel.innerHTML = years.map(y => `<option value="${y}">FY ${fyLabel(y)}</option>`).join('');

    function draw() {
      const year = parseInt(yearSel.value);
      const data = computeMonthlyPnl(year);
      const totRev = data.reduce((s, m) => s + m.revenue, 0);
      const totCost = data.reduce((s, m) => s + m.totalCost, 0);
      const totNet = totRev - totCost;

      $('#pnlRevenue').textContent = fmt(totRev);
      $('#pnlExpenses').textContent = fmt(totCost);
      const netEl = $('#pnlNet');
      netEl.textContent = fmt(totNet);
      netEl.className = 'stat-value ' + (totNet >= 0 ? 'profit' : 'loss');

      const tbody = $('#tablePnl tbody');
      tbody.innerHTML = data.map((m, i) => {
        const margin = m.revenue > 0 ? ((m.net / m.revenue) * 100).toFixed(1) : '—';
        return `<tr>
          <td>${MONTHS[i]}</td>
          <td>${fmt(m.revenue)}</td><td>${fmt(m.itemCost)}</td>
          <td>${fmt(m.otherExp)}</td><td>${fmt(m.totalCost)}</td>
          <td class="${m.net >= 0 ? 'profit' : 'loss'}">${fmt(m.net)}</td>
          <td>${margin}${margin !== '—' ? '%' : ''}</td>
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
      const [snap1, snap2, snap3] = await Promise.all([
        colItemsIn.get(), colItemsOut.get(), colExpenses.get()
      ]);
      snap1.docs.forEach(d => batch.delete(d.ref));
      snap2.docs.forEach(d => batch.delete(d.ref));
      snap3.docs.forEach(d => batch.delete(d.ref));
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
    auth.signOut().then(() => {
      itemsIn = []; itemsOut = []; expenses = [];
      showLogin();
      toast('Logged out');
    });
  });

})();
