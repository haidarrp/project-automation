(function () {
  'use strict';

  const TABLE_SELECTOR = 'table.data-table, table.file-table, table.admin-table, table.training-table, table.leave-table';
  const SKIP_HEADER_RE = /^(aksi|action|opsi|tindakan|)$/i;
  const NUMBER_HEADER_RE = /^(no\.?|nomor)$/i;
  const MONTHS = {
    januari: 0, februari: 1, maret: 2, april: 3, mei: 4, juni: 5,
    juli: 6, agustus: 7, september: 8, oktober: 9, november: 10, desember: 11,
    jan: 0, feb: 1, mar: 2, apr: 3, jun: 5, jul: 6, agu: 7, ags: 7, sep: 8, okt: 9, nov: 10, des: 11
  };

  let scanQueued = false;
  let internalMutation = false;

  function normalize(value) {
    return String(value ?? '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLocaleLowerCase('id-ID');
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
      if (control.tagName === 'SELECT') {
        return control.options[control.selectedIndex]?.textContent || control.value || '';
      }
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
    if (match && Object.prototype.hasOwnProperty.call(MONTHS, match[1])) {
      return new Date(Number(match[2]), MONTHS[match[1]], 1).getTime();
    }
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
    return Array.from(table.tBodies[0].rows).filter((row) => {
      return !row.classList.contains('table-tools-empty-row') && row.cells.length === 1 && row.cells[0].hasAttribute('colspan');
    });
  }

  function getColumns(table) {
    const headerRow = table.tHead?.rows?.[0];
    if (!headerRow) return [];
    return Array.from(headerRow.cells).map((th, index) => {
      const label = cleanHeaderText(th);
      return {
        index,
        th,
        label,
        isNumber: NUMBER_HEADER_RE.test(label),
        isAction: SKIP_HEADER_RE.test(label)
      };
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
        <button class="table-tools-reset" type="button">Reset filter</button>
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

  function updateCount(table, visible, total) {
    const count = table.__tableTools?.count;
    if (!count) return;
    count.textContent = total ? `${visible} dari ${total} baris` : '0 baris';
  }

  function renumber(table) {
    const columns = getColumns(table);
    const numberColumn = columns.find((column) => column.isNumber);
    if (!numberColumn) return;
    let number = 0;
    dataRows(table).forEach((row) => {
      if (row.hidden) return;
      number += 1;
      if (row.cells[numberColumn.index]) row.cells[numberColumn.index].textContent = String(number);
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

  function applyFilters(table) {
    const tools = table.__tableTools;
    if (!tools) return;
    const rows = dataRows(table);
    const originalEmpty = originalEmptyRows(table);
    const globalQuery = normalize(tools.search.value);
    const columnQueries = new Map();
    tools.filterRow?.querySelectorAll('.table-tools-column-filter').forEach((input) => {
      const query = normalize(input.value);
      if (query) columnQueries.set(Number(input.dataset.column), query);
    });

    removeToolsEmptyRow(table);
    let visible = 0;
    rows.forEach((row) => {
      const allText = normalize(Array.from(row.cells).map(cellText).join(' '));
      let matches = !globalQuery || allText.includes(globalQuery);
      if (matches && columnQueries.size) {
        for (const [columnIndex, query] of columnQueries.entries()) {
          if (!normalize(cellText(row.cells[columnIndex])).includes(query)) {
            matches = false;
            break;
          }
        }
      }
      row.hidden = !matches;
      if (matches) visible += 1;
    });

    const filtering = Boolean(globalQuery || columnQueries.size);
    originalEmpty.forEach((row) => { row.hidden = filtering && rows.length > 0; });
    if (rows.length > 0 && visible === 0) renderToolsEmptyRow(table, getColumns(table).length);
    updateCount(table, visible, rows.length);
    renumber(table);
  }

  function sortTable(table, columnIndex, direction) {
    const tbody = table.tBodies[0];
    if (!tbody) return;
    const rows = dataRows(table);
    rows.forEach((row, index) => {
      if (!row.dataset.tableOriginalOrder) row.dataset.tableOriginalOrder = String(index);
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

    internalMutation = true;
    rows.forEach((row) => tbody.appendChild(row));
    internalMutation = false;

    const columns = getColumns(table);
    columns.forEach((column) => {
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
    applyFilters(table);
  }

  function resetTable(table) {
    const tools = table.__tableTools;
    if (!tools) return;
    tools.search.value = '';
    tools.filterRow?.querySelectorAll('.table-tools-column-filter').forEach((input) => { input.value = ''; });

    const tbody = table.tBodies[0];
    if (tbody) {
      const rows = dataRows(table);
      rows.sort((a, b) => Number(a.dataset.tableOriginalOrder || 0) - Number(b.dataset.tableOriginalOrder || 0));
      internalMutation = true;
      rows.forEach((row) => tbody.appendChild(row));
      internalMutation = false;
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

    table.__tableTools = {
      toolbar,
      search: toolbar.querySelector('.table-tools-search'),
      count: toolbar.querySelector('.table-tools-count'),
      reset: toolbar.querySelector('.table-tools-reset'),
      filterRow
    };

    table.__tableTools.search.addEventListener('input', () => applyFilters(table));
    filterRow?.addEventListener('input', (event) => {
      if (event.target.classList.contains('table-tools-column-filter')) applyFilters(table);
    });
    table.__tableTools.reset.addEventListener('click', () => resetTable(table));

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
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          activateSort();
        }
      });
    });

    applyFilters(table);
  }

  function refreshPreparedTable(table) {
    if (!table.__tableTools) return;
    dataRows(table).forEach((row, index) => {
      if (!row.dataset.tableOriginalOrder) row.dataset.tableOriginalOrder = String(index);
    });
    applyFilters(table);
  }

  function scan(root = document) {
    const tables = [];
    if (root.matches?.(TABLE_SELECTOR)) tables.push(root);
    root.querySelectorAll?.(TABLE_SELECTOR).forEach((table) => tables.push(table));
    tables.forEach((table) => {
      if (table.dataset.tableToolsReady === 'true') refreshPreparedTable(table);
      else bindTable(table);
    });
  }

  function queueScan(root) {
    if (scanQueued) return;
    scanQueued = true;
    requestAnimationFrame(() => {
      scanQueued = false;
      scan(root || document);
    });
  }

  function start() {
    scan(document);
    const observer = new MutationObserver((mutations) => {
      if (internalMutation) return;
      let shouldScan = false;
      for (const mutation of mutations) {
        if (mutation.type !== 'childList' || (!mutation.addedNodes.length && !mutation.removedNodes.length)) continue;
        const changed = [...mutation.addedNodes, ...mutation.removedNodes].filter((node) => node.nodeType === Node.ELEMENT_NODE);
        if (!changed.length) continue;
        const helperOnly = changed.every((node) =>
          node.classList?.contains('table-tools-empty-row') ||
          node.classList?.contains('table-tools-filter-row') ||
          node.classList?.contains('table-tools-bar')
        );
        if (!helperOnly) {
          shouldScan = true;
          break;
        }
      }
      if (shouldScan) queueScan(document);
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();

  window.PusdatinTableTools = { scan, reset: resetTable, applyFilters, sortTable };
})();
