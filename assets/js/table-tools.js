(function () {
  'use strict';

  const TABLE_SELECTOR = 'table.data-table, table.file-table, table.admin-table, table.training-table, table.leave-table';
  const SKIP_HEADER_RE = /^(aksi|action|opsi|tindakan|)$/i;
  const NUMBER_HEADER_RE = /^(no\.?|nomor)$/i;
  const DEFAULT_PAGE_SIZE = 25;
  const PAGE_SIZES = [25, 50, 100];
  const FILTER_DEBOUNCE_MS = 260;
  const MONTHS = {
    januari: 0, februari: 1, maret: 2, april: 3, mei: 4, juni: 5,
    juli: 6, agustus: 7, september: 8, oktober: 9, november: 10, desember: 11,
    jan: 0, feb: 1, mar: 2, apr: 3, jun: 5, jul: 6, agu: 7, ags: 7, sep: 8, okt: 9, nov: 10, des: 11
  };

  const pendingRoots = new Set();
  let scanQueued = false;

  function normalize(value) {
    return String(value ?? '').replace(/\s+/g, ' ').trim().toLocaleLowerCase('id-ID');
  }

  function debounce(fn, delay) {
    let timer = null;
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => fn(...args), delay);
    };
  }

  function cleanHeaderText(th) {
    const clone = th.cloneNode(true);
    clone.querySelectorAll('.table-sort-indicator').forEach((node) => node.remove());
    return clone.textContent.replace(/\s+/g, ' ').trim();
  }

  function cellText(cell) {
    if (!cell) return '';
    const controls = Array.from(cell.querySelectorAll('select, input, textarea'));
    const controlValues = controls.map((control) => {
      if (control.tagName === 'SELECT') return control.options[control.selectedIndex]?.textContent || control.value || '';
      return control.value || '';
    });
    const clone = cell.cloneNode(true);
    clone.querySelectorAll('select, input, textarea').forEach((node) => node.remove());
    return [clone.textContent, ...controlValues].join(' ').replace(/\s+/g, ' ').trim();
  }

  function parseDate(value) {
    const raw = normalize(value).replace(/,/g, '');
    if (!raw) return null;
    let match = raw.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})\b/);
    if (match) {
      const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
      return Number.isNaN(date.getTime()) ? null : date.getTime();
    }
    match = raw.match(/\b(\d{1,2})\s+([a-z]+)\s+(\d{4})\b/);
    if (match && Object.prototype.hasOwnProperty.call(MONTHS, match[2])) {
      const date = new Date(Number(match[3]), MONTHS[match[2]], Number(match[1]));
      return Number.isNaN(date.getTime()) ? null : date.getTime();
    }
    match = raw.match(/\b([a-z]+)\s+(\d{4})\b/);
    if (match && Object.prototype.hasOwnProperty.call(MONTHS, match[1])) return new Date(Number(match[2]), MONTHS[match[1]], 1).getTime();
    return null;
  }

  function parseNumeric(value) {
    const raw = normalize(value);
    if (!raw) return null;
    const hasNumericSignal = /\d/.test(raw) && (/^(rp\s*)?[+-]?[\d.,]+\s*(%|jam|hari|pegawai|orang)?$/i.test(raw) || /^(rp|[+-]?\d)/i.test(raw));
    if (!hasNumericSignal) return null;
    let token = raw.replace(/rp\s*/gi, '').replace(/[^\d,.-]/g, '');
    if (!token || token === '-' || token === '.' || token === ',') return null;
    const comma = token.lastIndexOf(',');
    const dot = token.lastIndexOf('.');
    if (comma > -1 && dot > -1) {
      if (comma > dot) token = token.replace(/\./g, '').replace(',', '.');
      else token = token.replace(/,/g, '');
    } else if (comma > -1) {
      const decimals = token.length - comma - 1;
      token = decimals > 0 && decimals <= 2 ? token.replace(',', '.') : token.replace(/,/g, '');
    } else if (dot > -1) {
      const chunks = token.split('.');
      if (chunks.length > 2 || (chunks.length === 2 && chunks[1].length === 3)) token = chunks.join('');
    }
    const digitsOnly = token.replace(/[^\d]/g, '');
    if (/^[-+]?\d+$/.test(token) && digitsOnly.length > 15) return null;
    const number = Number(token);
    return Number.isFinite(number) ? number : null;
  }

  function comparable(value) {
    const date = parseDate(value);
    if (date !== null) return { type: 'number', value: date };
    const numeric = parseNumeric(value);
    if (numeric !== null) return { type: 'number', value: numeric };
    return { type: 'text', value: normalize(value) };
  }

  function dataRows(table) {
    if (!table.tBodies.length) return [];
    return Array.from(table.tBodies[0].rows).filter((row) => {
      if (row.classList.contains('table-tools-empty-row')) return false;
      if (row.cells.length === 1 && row.cells[0].hasAttribute('colspan')) return false;
      return true;
    });
  }

  function originalEmptyRows(table) {
    if (!table.tBodies.length) return [];
    return Array.from(table.tBodies[0].rows).filter((row) => !row.classList.contains('table-tools-empty-row') && row.cells.length === 1 && row.cells[0].hasAttribute('colspan'));
  }

  function getColumns(table) {
    const headerRow = table.tHead?.rows?.[0];
    if (!headerRow) return [];
    return Array.from(headerRow.cells).map((th, index) => {
      const label = cleanHeaderText(th);
      return { index, th, label, isNumber: NUMBER_HEADER_RE.test(label), isAction: SKIP_HEADER_RE.test(label) };
    });
  }

  function createToolbar(table) {
    const toolbar = document.createElement('div');
    toolbar.className = 'table-tools-bar';
    toolbar.innerHTML = `
      <label class="table-tools-search-wrap">
        <span class="table-tools-search-icon" aria-hidden="true">⌕</span>
        <input class="table-tools-search" type="search" placeholder="Cari pada tabel..." autocomplete="off" aria-label="Cari pada tabel">
      </label>
      <div class="table-tools-actions">
        <span class="table-tools-count" aria-live="polite"></span>
        <label class="table-tools-page-size-wrap"><span>Tampil</span><select class="table-tools-page-size" aria-label="Jumlah baris per halaman">${PAGE_SIZES.map((size) => `<option value="${size}" ${size === DEFAULT_PAGE_SIZE ? 'selected' : ''}>${size}</option>`).join('')}</select></label>
        <div class="table-tools-pager" aria-label="Navigasi halaman tabel"><button class="table-tools-page-prev" type="button" aria-label="Halaman sebelumnya">‹</button><span class="table-tools-page-label">1 / 1</span><button class="table-tools-page-next" type="button" aria-label="Halaman berikutnya">›</button></div>
        <button class="table-tools-reset" type="button">Reset</button>
      </div>`;
    table.parentNode.insertBefore(toolbar, table);
    return toolbar;
  }

  function createFilterRow(table, columns) {
    const thead = table.tHead;
    if (!thead) return null;
    const row = document.createElement('tr');
    row.className = 'table-tools-filter-row';
    columns.forEach((column) => {
      const cell = document.createElement('th');
      cell.className = 'table-tools-filter-cell';
      if (!column.isAction && !column.isNumber) {
        const input = document.createElement('input');
        input.type = 'search';
        input.className = 'table-tools-column-filter';
        input.placeholder = 'Filter';
        input.autocomplete = 'off';
        input.setAttribute('aria-label', `Filter kolom ${column.label || column.index + 1}`);
        input.dataset.column = String(column.index);
        cell.appendChild(input);
      }
      row.appendChild(cell);
    });
    thead.appendChild(row);
    return row;
  }

  function addSortControls(columns) {
    columns.forEach((column) => {
      if (column.isAction || column.isNumber) return;
      column.th.classList.add('table-tools-sortable');
      column.th.tabIndex = 0;
      column.th.setAttribute('aria-sort', 'none');
      column.th.setAttribute('title', `Urutkan berdasarkan ${column.label}`);
      const indicator = document.createElement('span');
      indicator.className = 'table-sort-indicator';
      indicator.setAttribute('aria-hidden', 'true');
      indicator.textContent = '↕';
      column.th.appendChild(indicator);
    });
  }

  function removeToolsEmptyRow(table) {
    table.tBodies[0]?.querySelector('.table-tools-empty-row')?.remove();
  }

  function renderToolsEmptyRow(table, columnCount) {
    if (!table.tBodies.length || table.tBodies[0].querySelector('.table-tools-empty-row')) return;
    const row = table.tBodies[0].insertRow();
    row.className = 'table-tools-empty-row';
    const cell = row.insertCell();
    cell.colSpan = Math.max(columnCount, 1);
    cell.innerHTML = '<div class="table-tools-empty"><strong>Tidak ada data yang sesuai filter.</strong><span>Ubah kata pencarian atau reset filter tabel.</span></div>';
  }

  function pageState(table, filteredCount) {
    const tools = table.__tableTools;
    const pageSize = Math.max(1, Number(tools.pageSize.value || DEFAULT_PAGE_SIZE));
    const totalPages = Math.max(1, Math.ceil(filteredCount / pageSize));
    tools.page = Math.max(1, Math.min(Number(tools.page || 1), totalPages));
    return { pageSize, totalPages, page: tools.page };
  }

  function updateToolbar(table, filteredCount, totalCount, start, end, totalPages) {
    const tools = table.__tableTools;
    if (!tools) return;
    if (!filteredCount) tools.count.textContent = totalCount ? `0 hasil dari ${totalCount} baris` : '0 baris';
    else if (filteredCount === totalCount) tools.count.textContent = `${start + 1}-${end} dari ${totalCount} baris`;
    else tools.count.textContent = `${start + 1}-${end} dari ${filteredCount} hasil (${totalCount} baris)`;
    tools.pageLabel.textContent = `${tools.page} / ${totalPages}`;
    tools.prev.disabled = tools.page <= 1;
    tools.next.disabled = tools.page >= totalPages;
    tools.pager.hidden = filteredCount <= Number(tools.pageSize.value || DEFAULT_PAGE_SIZE);
  }

  function applyFilters(table, options) {
    const tools = table.__tableTools;
    if (!tools) return;
    if (options?.resetPage) tools.page = 1;
    const rows = dataRows(table);
    const originalEmpty = originalEmptyRows(table);
    const globalQuery = normalize(tools.search.value);
    const columnQueries = new Map();
    tools.filterRow?.querySelectorAll('.table-tools-column-filter').forEach((input) => {
      const query = normalize(input.value);
      if (query) columnQueries.set(Number(input.dataset.column), query);
    });

    removeToolsEmptyRow(table);
    const matched = [];
    rows.forEach((row) => {
      const allText = normalize(Array.from(row.cells).map(cellText).join(' '));
      let matches = !globalQuery || allText.includes(globalQuery);
      if (matches && columnQueries.size) {
        for (const [columnIndex, query] of columnQueries.entries()) {
          if (!normalize(cellText(row.cells[columnIndex])).includes(query)) { matches = false; break; }
        }
      }
      row.dataset.tableFilteredOut = matches ? 'false' : 'true';
      if (matches) matched.push(row);
    });

    const { pageSize, totalPages, page } = pageState(table, matched.length);
    const start = (page - 1) * pageSize;
    const end = Math.min(start + pageSize, matched.length);
    const visibleSet = new Set(matched.slice(start, end));

    rows.forEach((row) => { row.hidden = !visibleSet.has(row); });
    const numberColumn = getColumns(table).find((column) => column.isNumber);
    if (numberColumn) matched.forEach((row, index) => {
      if (row.cells[numberColumn.index]) row.cells[numberColumn.index].textContent = String(index + 1);
    });

    const filtering = Boolean(globalQuery || columnQueries.size);
    originalEmpty.forEach((row) => { row.hidden = filtering && rows.length > 0; });
    if (rows.length > 0 && matched.length === 0) renderToolsEmptyRow(table, getColumns(table).length);
    updateToolbar(table, matched.length, rows.length, start, end, totalPages);
  }

  function sortTable(table, columnIndex, direction) {
    const tbody = table.tBodies[0];
    if (!tbody) return;
    const rows = dataRows(table);
    rows.forEach((row, index) => {
      if (row.dataset.tableOriginalOrder == null || row.dataset.tableOriginalOrder === '') row.dataset.tableOriginalOrder = String(index);
    });
    rows.sort((rowA, rowB) => {
      const a = comparable(cellText(rowA.cells[columnIndex]));
      const b = comparable(cellText(rowB.cells[columnIndex]));
      let result = 0;
      if (a.type === 'number' && b.type === 'number') result = a.value - b.value;
      else result = String(a.value).localeCompare(String(b.value), 'id-ID', { numeric: true, sensitivity: 'base' });
      if (result === 0) result = Number(rowA.dataset.tableOriginalOrder) - Number(rowB.dataset.tableOriginalOrder);
      return direction === 'desc' ? -result : result;
    });
    rows.forEach((row) => tbody.appendChild(row));

    getColumns(table).forEach((column) => {
      if (column.isAction || column.isNumber) return;
      const active = column.index === columnIndex;
      column.th.setAttribute('aria-sort', active ? (direction === 'asc' ? 'ascending' : 'descending') : 'none');
      column.th.classList.toggle('table-tools-sorted-asc', active && direction === 'asc');
      column.th.classList.toggle('table-tools-sorted-desc', active && direction === 'desc');
      const indicator = column.th.querySelector('.table-sort-indicator');
      if (indicator) indicator.textContent = active ? (direction === 'asc' ? '↑' : '↓') : '↕';
    });
    table.dataset.tableSortColumn = String(columnIndex);
    table.dataset.tableSortDirection = direction;
    table.__tableTools.page = 1;
    applyFilters(table);
  }

  function resetTable(table) {
    const tools = table.__tableTools;
    if (!tools) return;
    tools.search.value = '';
    tools.filterRow?.querySelectorAll('.table-tools-column-filter').forEach((input) => { input.value = ''; });
    tools.pageSize.value = String(DEFAULT_PAGE_SIZE);
    tools.page = 1;
    const tbody = table.tBodies[0];
    if (tbody) {
      const rows = dataRows(table);
      rows.sort((a, b) => Number(a.dataset.tableOriginalOrder || 0) - Number(b.dataset.tableOriginalOrder || 0));
      rows.forEach((row) => tbody.appendChild(row));
    }
    delete table.dataset.tableSortColumn;
    delete table.dataset.tableSortDirection;
    getColumns(table).forEach((column) => {
      if (column.isAction || column.isNumber) return;
      column.th.setAttribute('aria-sort', 'none');
      column.th.classList.remove('table-tools-sorted-asc', 'table-tools-sorted-desc');
      const indicator = column.th.querySelector('.table-sort-indicator');
      if (indicator) indicator.textContent = '↕';
    });
    applyFilters(table);
  }

  function bindTable(table) {
    if (table.dataset.tableToolsReady === 'true' || !table.tHead || !table.tBodies.length) return;
    const columns = getColumns(table);
    if (!columns.length) return;

    table.dataset.tableToolsReady = 'true';
    table.classList.add('table-tools-enabled');
    dataRows(table).forEach((row, index) => { row.dataset.tableOriginalOrder = String(index); });

    const toolbar = createToolbar(table);
    const filterRow = createFilterRow(table, columns);
    addSortControls(columns);
    const tools = {
      toolbar,
      search: toolbar.querySelector('.table-tools-search'),
      count: toolbar.querySelector('.table-tools-count'),
      reset: toolbar.querySelector('.table-tools-reset'),
      pageSize: toolbar.querySelector('.table-tools-page-size'),
      pager: toolbar.querySelector('.table-tools-pager'),
      prev: toolbar.querySelector('.table-tools-page-prev'),
      next: toolbar.querySelector('.table-tools-page-next'),
      pageLabel: toolbar.querySelector('.table-tools-page-label'),
      filterRow,
      page: 1
    };
    table.__tableTools = tools;

    const debouncedFilter = debounce(() => applyFilters(table, { resetPage: true }), FILTER_DEBOUNCE_MS);
    tools.search.addEventListener('input', debouncedFilter);
    filterRow?.addEventListener('input', (event) => {
      if (event.target.classList.contains('table-tools-column-filter')) debouncedFilter();
    });
    tools.pageSize.addEventListener('change', () => { tools.page = 1; applyFilters(table); });
    tools.prev.addEventListener('click', () => { tools.page = Math.max(1, tools.page - 1); applyFilters(table); });
    tools.next.addEventListener('click', () => { tools.page += 1; applyFilters(table); });
    tools.reset.addEventListener('click', () => resetTable(table));

    columns.forEach((column) => {
      if (column.isAction || column.isNumber) return;
      const activateSort = () => {
        const currentColumn = Number(table.dataset.tableSortColumn);
        const currentDirection = table.dataset.tableSortDirection;
        const direction = currentColumn === column.index && currentDirection === 'asc' ? 'desc' : 'asc';
        sortTable(table, column.index, direction);
      };
      column.th.addEventListener('click', activateSort);
      column.th.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); activateSort(); }
      });
    });

    applyFilters(table);
  }

  function refreshPreparedTable(table) {
    if (!table.__tableTools) return;
    dataRows(table).forEach((row, index) => {
      if (row.dataset.tableOriginalOrder == null || row.dataset.tableOriginalOrder === '') row.dataset.tableOriginalOrder = String(index);
    });
    applyFilters(table);
  }

  function scan(root = document) {
    if (!root || !root.isConnected && root !== document) return;
    const tables = [];
    if (root.matches?.(TABLE_SELECTOR)) tables.push(root);
    root.querySelectorAll?.(TABLE_SELECTOR).forEach((table) => tables.push(table));
    const ownerTable = root.closest?.(TABLE_SELECTOR);
    if (ownerTable && !tables.includes(ownerTable)) tables.push(ownerTable);
    tables.forEach((table) => table.dataset.tableToolsReady === 'true' ? refreshPreparedTable(table) : bindTable(table));
  }

  function queueScan(root) {
    if (root) pendingRoots.add(root);
    if (scanQueued) return;
    scanQueued = true;
    requestAnimationFrame(() => {
      scanQueued = false;
      const roots = [...pendingRoots];
      pendingRoots.clear();
      if (!roots.length) return;
      roots.forEach(scan);
    });
  }

  function isHelperNode(node) {
    return node?.classList?.contains('table-tools-empty-row') || node?.classList?.contains('table-tools-filter-row') || node?.classList?.contains('table-tools-bar');
  }

  function start() {
    scan(document);
    const observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        if (mutation.type !== 'childList') return;
        const targetTable = mutation.target?.closest?.(TABLE_SELECTOR);
        const changed = [...mutation.addedNodes, ...mutation.removedNodes].filter((node) => node.nodeType === Node.ELEMENT_NODE);
        if (changed.length && changed.every(isHelperNode)) return;
        if (targetTable) queueScan(targetTable);
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType !== Node.ELEMENT_NODE || isHelperNode(node)) return;
          if (node.matches?.(TABLE_SELECTOR) || node.querySelector?.(TABLE_SELECTOR)) queueScan(node);
        });
      });
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();

  window.PusdatinTableTools = { scan, reset: resetTable, applyFilters, sortTable };
})();
