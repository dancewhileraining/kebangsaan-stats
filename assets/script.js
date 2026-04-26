/* Kebangsaan Statistics v2.3.0 */

/* ─── JNews Dark Mode Switcher compatibility ──────────────────────────────
   JNews Dark Mode Switcher adds `.jnews-dark-mode` to <html>.
   We also keep `jeg_dark_mode` as fallback for the JNews built-in toggle.
   A dynamically injected <style> always wins the cascade over theme CSS.   */
(function () {
    var STYLE_ID = 'kbgs-dark-injected';

    /* Primary: jnews-dark-mode on <html> (JNews Dark Mode Switcher plugin)
       Fallback: jeg_dark_mode on <body> (JNews theme built-in toggle)      */
    var DARK_CLASSES = [
        'jnews-dark-mode',  /* JNews Dark Mode Switcher plugin — on <html>  */
        'jeg_dark_mode',    /* JNews theme built-in toggle   — on <body>    */
        'jeg-dark-mode'
    ];

    var DARK_CSS = [
        '.kbgs-root{color:#e5e7eb!important;background:transparent!important}',
        '.kbgs-table-wrap{background:#1e2530!important;border-color:#364152!important}',
        '.kbgs-table{background:#1e2530!important;color:#e5e7eb!important}',
        '.kbgs-table thead tr th{background:#232b38!important;color:#cbd5e1!important;border-color:#364152!important}',
        '.kbgs-table tbody tr{background:#1e2530!important;border-color:#2b3443!important}',
        '.kbgs-table tbody tr:nth-child(even){background:#222a36!important}',
        '.kbgs-table tbody tr:hover{background:#2a3442!important}',
        '.kbgs-table tbody tr:nth-child(even):hover{background:#2a3442!important}',
        '.kbgs-table td{color:#e5e7eb!important;border-color:#2b3443!important;background:transparent!important}',
        '.kbgs-td-no{color:#64748b!important}',
        '.kbgs-td-kb{color:#f1f5f9!important}',
        '.kbgs-num-datang{color:#93c5fd!important}',
        '.kbgs-num-berangkat{color:#86efac!important}',
        '.kbgs-table tfoot tr{background:#232b38!important;border-color:#475569!important}',
        '.kbgs-table tfoot td{color:#cbd5e1!important;border-color:#364152!important;background:transparent!important}',
        '.kbgs-stat-datang{background:#152033!important;border-color:#1e3a8a!important}',
        '.kbgs-stat-berangkat{background:#0f2417!important;border-color:#166534!important}',
        '.kbgs-stat-datang .kbgs-stat-value{color:#93c5fd!important}',
        '.kbgs-stat-berangkat .kbgs-stat-value{color:#86efac!important}',
        '.kbgs-stat-label{color:#94a3b8!important}',
        '.kbgs-stat-sub{color:#64748b!important}',
        '.kbgs-year-tab{background:#232b38!important;border-color:#364152!important;color:#94a3b8!important}',
        '.kbgs-year-tab:hover{background:#2a3442!important;color:#e5e7eb!important}',
        '.kbgs-year-tab.active{background:#60a5fa!important;border-color:#60a5fa!important;color:#fff!important}',
        '.kbgs-tab{color:#94a3b8!important;background:transparent!important}',
        '.kbgs-tab:hover{background:#2a3442!important;color:#e5e7eb!important}',
        '.kbgs-tab.active{color:#60a5fa!important;border-bottom-color:#60a5fa!important}',
        '.kbgs-tabs{border-bottom-color:#364152!important}',
        '.kbgs-search{background-color:#1e2530!important;color:#e5e7eb!important;border-color:#364152!important}',
        '.kbgs-search::placeholder{color:#64748b!important}',
    ].join('');

    function applyDark() {
        if (document.getElementById(STYLE_ID)) return;
        var el = document.createElement('style');
        el.id = STYLE_ID;
        el.textContent = DARK_CSS;
        document.head.appendChild(el);
    }

    function removeDark() {
        var el = document.getElementById(STYLE_ID);
        if (el) el.parentNode.removeChild(el);
    }

    function hasDarkClass(el) {
        if (!el || !el.classList) return false;
        for (var i = 0; i < DARK_CLASSES.length; i++) {
            if (el.classList.contains(DARK_CLASSES[i])) return true;
        }
        return false;
    }

    function sync() {
        /* Check both <html> and <body> — JNews may add the class to either */
        if (hasDarkClass(document.documentElement) || hasDarkClass(document.body)) {
            applyDark();
        } else {
            removeDark();
        }
    }

    function startObserving() {
        var observer = new MutationObserver(sync);
        var opts = { attributes: true, attributeFilter: ['class', 'data-mode', 'data-theme'] };
        observer.observe(document.documentElement, opts);
        if (document.body) observer.observe(document.body, opts);
        sync();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', startObserving);
    } else {
        startObserving();
    }
}());

/* ─── Year tab switching ─── */
function kbgsYearTab(btn, uid) {
    var wrap = document.getElementById(uid);
    var year = btn.getAttribute('data-year');

    wrap.querySelectorAll('.kbgs-year-tab').forEach(function(t) {
        t.classList.remove('active');
        t.setAttribute('aria-selected', 'false');
    });
    wrap.querySelectorAll('.kbgs-year-panel').forEach(function(p) {
        p.classList.remove('active');
    });
    btn.classList.add('active');
    btn.setAttribute('aria-selected', 'true');

    var panel = wrap.querySelector('.kbgs-year-panel[data-year="' + year + '"]');
    if (panel) panel.classList.add('active');

    /* Clear search when switching years */
    var s = wrap.querySelector('.kbgs-search');
    if (s) s.value = '';
}

/* ─── Month tab switching ─── */
function kbgsTab(btn, uid) {
    var wrap  = document.getElementById(uid);
    var month = btn.getAttribute('data-month');

    /* Scope to the year panel that owns this button */
    var scope = btn.closest('.kbgs-year-panel') || wrap;

    scope.querySelectorAll('.kbgs-tab').forEach(function(t) {
        t.classList.remove('active');
        t.setAttribute('aria-selected', 'false');
    });
    scope.querySelectorAll('.kbgs-panel').forEach(function(p) {
        p.classList.remove('active');
    });
    btn.classList.add('active');
    btn.setAttribute('aria-selected', 'true');

    var panel = scope.querySelector('.kbgs-panel[data-month="' + month + '"]');
    if (panel) panel.classList.add('active');

    /* Re-apply search if active */
    var s = wrap.querySelector('.kbgs-search');
    if (s && s.value) kbgsFilter(s);
}

/* ─── Search / filter ─── */
function kbgsFilter(input) {
    var wrap  = input.closest('.kbgs-root');
    var query = input.value.toLowerCase().trim();

    /* Find active year panel, then active month panel within it */
    var yearPanel = wrap.querySelector('.kbgs-year-panel.active') || wrap;
    var panel     = yearPanel.querySelector('.kbgs-panel.active');
    if (!panel) return;

    panel.querySelectorAll('tbody tr').forEach(function(tr) {
        var kb = tr.querySelector('.kbgs-td-kb');
        if (!kb) return;
        tr.classList.toggle('kbgs-hidden',
            kb.textContent.toLowerCase().indexOf(query) === -1);
    });
}

/* ─── Shared: extract table data from the panel containing the button ─── */
function kbgsGetTableData(btn) {
    var panel      = btn.closest('.kbgs-panel');
    var monthLabel = btn.getAttribute('data-month-label') || '';
    var year       = btn.getAttribute('data-year') || '';
    var rows       = [];

    panel.querySelectorAll('tbody tr:not(.kbgs-hidden)').forEach(function(tr) {
        var cells = tr.querySelectorAll('td');
        rows.push([
            cells[0] ? cells[0].textContent.trim() : '',
            cells[1] ? cells[1].textContent.trim() : '',
            cells[2] ? parseInt(cells[2].textContent.replace(/[^0-9]/g, ''), 10) : 0,
            cells[3] ? parseInt(cells[3].textContent.replace(/[^0-9]/g, ''), 10) : 0,
        ]);
    });

    var tfoot = panel.querySelector('tfoot tr');
    var totD = 0, totB = 0;
    if (tfoot) {
        var fc = tfoot.querySelectorAll('td');
        totD = parseInt((fc[1] ? fc[1].textContent : '0').replace(/[^0-9]/g, ''), 10);
        totB = parseInt((fc[2] ? fc[2].textContent : '0').replace(/[^0-9]/g, ''), 10);
    }

    return { monthLabel: monthLabel, year: year, rows: rows, totD: totD, totB: totB };
}

/* ─── Export Excel ─── */
function kbgsExportExcel(btn) {
    var d = kbgsGetTableData(btn);

    var aoa = [
        ['Statistik Kedatangan & Keberangkatan Per Kebangsaan'],
        [d.monthLabel + ' ' + d.year],
        [],
        ['NO', 'KEBANGSAAN', 'KEDATANGAN', 'KEBERANGKATAN'],
    ];
    d.rows.forEach(function(r) { aoa.push(r); });
    aoa.push([]);
    aoa.push(['', 'JUMLAH', d.totD, d.totB]);

    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 6 }, { wch: 36 }, { wch: 16 }, { wch: 18 }];

    XLSX.utils.book_append_sheet(wb, ws, d.monthLabel || 'Data');
    XLSX.writeFile(wb, 'Statistik_Kebangsaan_' + d.monthLabel + '_' + d.year + '.xlsx');
}

/* ─── Export PDF ─── */
function kbgsExportPdf(btn) {
    var d = kbgsGetTableData(btn);
    var { jsPDF } = window.jspdf;
    var doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text('Statistik Kedatangan & Keberangkatan Per Kebangsaan', 14, 18);

    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100);
    doc.text(d.monthLabel + ' ' + d.year, 14, 25);
    doc.setTextColor(0);

    var body = d.rows.map(function(r) {
        return [r[0], r[1], r[2].toLocaleString(), r[3].toLocaleString()];
    });
    body.push(['', 'JUMLAH', d.totD.toLocaleString(), d.totB.toLocaleString()]);

    doc.autoTable({
        startY: 30,
        head: [['NO', 'KEBANGSAAN', 'KEDATANGAN', 'KEBERANGKATAN']],
        body: body,
        styles: { fontSize: 9, cellPadding: 3 },
        headStyles: { fillColor: [248, 250, 252], textColor: [55, 65, 81], fontStyle: 'bold', lineWidth: 0.2, lineColor: [226, 232, 240] },
        columnStyles: {
            0: { halign: 'center', cellWidth: 12 },
            1: { halign: 'left',   cellWidth: 80 },
            2: { halign: 'right',  cellWidth: 40 },
            3: { halign: 'right',  cellWidth: 48 },
        },
        didParseCell: function(data) {
            if (data.row.index === body.length - 1) {
                data.cell.styles.fontStyle = 'bold';
                data.cell.styles.fillColor = [248, 250, 252];
            }
        },
    });

    doc.save('Statistik_Kebangsaan_' + d.monthLabel + '_' + d.year + '.pdf');
}

/* ─── Sort columns ─── */
function kbgsSort(th, colClass) {
    var table = th.closest('table');
    var tbody = table.querySelector('tbody');
    var rows  = Array.from(tbody.querySelectorAll('tr'));
    var asc   = th.getAttribute('data-sort') !== 'asc';

    table.querySelectorAll('.kbgs-th-datang, .kbgs-th-berangkat').forEach(function(h) {
        h.removeAttribute('data-sort');
        var ic = h.querySelector('.kbgs-si');
        if (ic) ic.textContent = '↕';
    });
    th.setAttribute('data-sort', asc ? 'asc' : 'desc');
    var ic = th.querySelector('.kbgs-si');
    if (ic) ic.textContent = asc ? '↑' : '↓';

    rows.sort(function(a, b) {
        var aCell = a.querySelector('.' + colClass);
        var bCell = b.querySelector('.' + colClass);
        if (!aCell || !bCell) return 0;
        var aN = parseFloat(aCell.textContent.replace(/[^0-9.]/g, ''));
        var bN = parseFloat(bCell.textContent.replace(/[^0-9.]/g, ''));
        if (!isNaN(aN) && !isNaN(bN)) return asc ? aN - bN : bN - aN;
        return asc
            ? aCell.textContent.localeCompare(bCell.textContent)
            : bCell.textContent.localeCompare(aCell.textContent);
    });

    rows.forEach(function(r) { tbody.appendChild(r); });

    var n = 1;
    rows.forEach(function(r) {
        if (!r.classList.contains('kbgs-hidden')) {
            var c = r.querySelector('.kbgs-td-no');
            if (c) c.textContent = n++;
        }
    });
}
