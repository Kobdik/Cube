const PIVOT_SHEET = 'Сводка';
const CATALOG_SHEET = 'Измерения';
const DIM_LABELS = {
    0: 'План',
    1: 'Календарь',
    2: 'Движения',
    3: 'Материалы',
    4: 'Склады',
};
const HDR_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF343A40' } };
const BAND_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF495057' } };
const TID_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } };
const SEL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8D4E8' } };
const THIN_BORDER = {
    top: { style: 'thin', color: { argb: 'FFCED4DA' } },
    left: { style: 'thin', color: { argb: 'FFCED4DA' } },
    bottom: { style: 'thin', color: { argb: 'FFCED4DA' } },
    right: { style: 'thin', color: { argb: 'FFCED4DA' } },
};

function dimLabel(ord, dim) {
    // if (DIM_LABELS[ord]) return DIM_LABELS[ord];
    if (dim && dim.root && dim.root.caption) return dim.root.caption;
    return `Измерение ${ord}`;
}

function colLetter(n) {
    let s = '', x = n;
    while (x > 0) {
        s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
        x = Math.floor((x - 1) / 26);
    }
    return s;
}

function quoteSheet(name) {
    return `'${String(name).replace(/'/g, "''")}'`;
}

function a1Range(sheetName, c1, r1, c2, r2) {
    const sh = quoteSheet(sheetName);
    if (c2 == null || r2 == null) return `${sh}!$${colLetter(c1)}$${r1}`;
    return `${sh}!$${colLetter(c1)}$${r1}:$${colLetter(c2)}$${r2}`;
}

function sanitizeExcelName(tid) {
    const core = String(tid ?? '').replace(/[^A-Za-z0-9._]/g, '_');
    let name = `code_${core || 'node'}`;
    if (name.length > 255) name = name.slice(0, 255);
    return name;
}

function uniqueDefinedName(used, tid, ord) {
    const base = sanitizeExcelName(tid);
    const key0 = base.toLowerCase();
    if (!used.has(key0)) {
        used.add(key0);
        return base;
    }
    let name = `${base}_ord${ord}`;
    let n = 2;
    while (used.has(name.toLowerCase()) || name.length > 255) {
        name = `${base}_ord${ord}_${n++}`;
        if (name.length > 255) name = name.slice(0, 255);
        if (n > 99) break;
    }
    used.add(name.toLowerCase());
    return name;
}

function addDefinedName(workbook, locStr, name) {
    // ExcelJS definedNames.add(location, name) — location first.
    workbook.definedNames.add(locStr, name);
}

function nodeOp(node) {
    return operToText(node && node.oper);
}

function operToText(oper) {
    if (oper == null || oper === '') return '';
    if (typeof oper === 'string') return oper;
    try {
        return JSON.stringify(oper);
    } catch {
        return '';
    }
}

function parentTid(node) {
    if (!node || !node.parent) return '';
    return node.parent.tid || '';
}

function isCollapsed(node) {
    return !!(node && node.fst_child && !node.exp);
}

function outlineLev(node) {
    const lev = (node && node.lev) ? node.lev : 0;
    return Math.max(0, Math.min(7, lev));
}

function styleHeaderCell(cell, value, fill = HDR_FILL) {
    cell.value = value;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10, name: 'Calibri' };
    cell.fill = fill;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = THIN_BORDER;
}

function buildPivotWorkbook(ExcelJS, engine) {
    if (!ExcelJS || typeof ExcelJS.Workbook !== 'function') {
        throw new Error('ExcelJS is not loaded');
    }
    let rows, cols;
    try {
        rows = engine.getVerRows(0);
        cols = engine.getHorRows(0);
    } catch (err) {
        throw new Error('matrix is not ready', { cause: err });
    }
    if (!Array.isArray(rows) || !Array.isArray(cols)) {
        throw new Error('matrix is not ready');
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'MOLAP';
    workbook.created = new Date();
    const usedNames = new Set();
    const nameByKey = new Map();

    const headerCols = 4;
    const headerRows = 3;
    const firstDataCol = headerCols + 1;
    const firstDataRow = headerRows + 1;
    const lastDataCol = headerCols + cols.length;
    const lastDataRow = headerRows + rows.length;
    const hasDataCols = cols.length > 0;
    const hasDataRows = rows.length > 0;

    const pivot = workbook.addWorksheet(PIVOT_SHEET, {
        views: [{
            state: 'frozen',
            xSplit: headerCols,
            ySplit: headerRows,
            topLeftCell: colLetter(firstDataCol) + firstDataRow,
            activeCell: colLetter(firstDataCol) + firstDataRow,
        }],
        properties: {
            outlineLevelRow: 7,
            outlineLevelCol: 7,
        },
    });
    // ExcelJS writes sheetPr/outlinePr from outlineProperties, not outlineSummaryBelow.
    pivot.properties.outlineProperties = {
        summaryBelow: false,
        summaryRight: false,
    };

    pivot.getColumn(1).width = 18;
    pivot.getColumn(2).width = 32;
    pivot.getColumn(3).width = 14;
    pivot.getColumn(4).width = 10;

    pivot.mergeCells(1, 1, 1, headerCols);
    styleHeaderCell(pivot.getCell(1, 1), 'Движения материалов', BAND_FILL);
    pivot.getCell(1, 1).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    for (let c = 2; c <= headerCols; c++) {
        pivot.getCell(1, c).fill = BAND_FILL;
        pivot.getCell(1, c).border = THIN_BORDER;
    }

    styleHeaderCell(pivot.getCell(2, 1), '');
    styleHeaderCell(pivot.getCell(2, 2), '');
    styleHeaderCell(pivot.getCell(2, 3), '');
    styleHeaderCell(pivot.getCell(2, 4), '');
    styleHeaderCell(pivot.getCell(3, 1), 'Код');
    styleHeaderCell(pivot.getCell(3, 2), 'Наименование');
    styleHeaderCell(pivot.getCell(3, 3), 'Измерение');
    styleHeaderCell(pivot.getCell(3, 4), 'Уровень');
    pivot.getRow(1).height = 22;
    pivot.getRow(2).height = 18;
    pivot.getRow(3).height = 36;

    let bandStart = 0;
    for (let i = 0; i <= cols.length; i++) {
        if (!cols.length) break;
        const endBand = i === cols.length || (i > 0 && cols[i].ord !== cols[bandStart].ord);
        if (endBand) {
            const c1 = firstDataCol + bandStart;
            const c2 = firstDataCol + i - 1;
            const label = dimLabel(cols[bandStart].ord, engine.getDim(cols[bandStart].ord));
            if (c2 > c1) pivot.mergeCells(1, c1, 1, c2);
            styleHeaderCell(pivot.getCell(1, c1), label, BAND_FILL);
            for (let c = c1 + 1; c <= c2; c++) {
                pivot.getCell(1, c).fill = BAND_FILL;
                pivot.getCell(1, c).border = THIN_BORDER;
            }
            bandStart = i;
        }
        if (i === cols.length) break;
        const colNode = cols[i];
        const excelCol = firstDataCol + i;
        const tidCell = pivot.getCell(2, excelCol);
        tidCell.value = colNode.tid;
        tidCell.font = { size: 9, name: 'Calibri', color: { argb: 'FF212529' } };
        tidCell.fill = TID_FILL;
        tidCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        tidCell.border = THIN_BORDER;
        const capCell = pivot.getCell(3, excelCol);
        capCell.value = colNode.caption;
        capCell.font = {
            bold: true,
            italic: isCollapsed(colNode),
            size: 9,
            name: 'Calibri',
            color: { argb: 'FFFFFFFF' },
        };
        capCell.fill = colNode.sel ? SEL_FILL : HDR_FILL;
        if (colNode.sel) {
            capCell.font = { bold: true, italic: isCollapsed(colNode), size: 9, name: 'Calibri', color: { argb: 'FF212529' } };
        }
        capCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        capCell.border = THIN_BORDER;
        const column = pivot.getColumn(excelCol);
        column.width = 14;
        column.outlineLevel = outlineLev(colNode);
    }

    rows.forEach((rowNode, i) => {
        const excelRow = firstDataRow + i;
        const row = pivot.getRow(excelRow);
        row.outlineLevel = outlineLev(rowNode);
        row.height = 18;
        const collapsed = isCollapsed(rowNode);
        const selected = !!rowNode.sel;
        const dim = engine.getDim(rowNode.ord);
        const codeCell = row.getCell(1);
        const nameCell = row.getCell(2);
        const dimCell = row.getCell(3);
        const levCell = row.getCell(4);
        codeCell.value = rowNode.tid;
        nameCell.value = rowNode.caption;
        dimCell.value = dimLabel(rowNode.ord, dim);
        levCell.value = rowNode.lev;
        const rowFont = {
            name: 'Calibri',
            size: 10,
            italic: collapsed,
            bold: !!rowNode.fst_child,
        };
        [codeCell, nameCell, dimCell, levCell].forEach((cell, idx) => {
            cell.font = rowFont;
            cell.border = THIN_BORDER;
            cell.alignment = {
                vertical: 'middle',
                horizontal: idx === 3 ? 'center' : 'left',
                indent: idx === 1 ? Math.max(0, (rowNode.lev || 1) - 1) : 0,
            };
            if (selected) cell.fill = SEL_FILL;
        });
        cols.forEach((colNode, j) => {
            const cell = row.getCell(firstDataCol + j);
            let val = 0;
            if (typeof rowNode.row === 'number' && typeof colNode.col === 'number') {
                val = engine.getVal(rowNode.row, colNode.col, 1);
            }
            cell.value = val;
            cell.numFmt = '#,##0';
            cell.font = { name: 'Calibri', size: 10 };
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
            cell.border = THIN_BORDER;
            if (selected || colNode.sel) cell.fill = SEL_FILL;
        });
    });

    if (hasDataCols && hasDataRows) {
        rows.forEach((rowNode, i) => {
            const excelRow = firstDataRow + i;
            const name = uniqueDefinedName(usedNames, rowNode.tid, rowNode.ord);
            nameByKey.set(`${rowNode.ord}\0${rowNode.tid}`, name);
            addDefinedName(workbook, a1Range(PIVOT_SHEET, firstDataCol, excelRow, lastDataCol, excelRow), name);
        });
        cols.forEach((colNode, j) => {
            const excelCol = firstDataCol + j;
            const name = uniqueDefinedName(usedNames, colNode.tid, colNode.ord);
            nameByKey.set(`${colNode.ord}\0${colNode.tid}`, name);
            addDefinedName(workbook, a1Range(PIVOT_SHEET, excelCol, firstDataRow, excelCol, lastDataRow), name);
        });
    } else if (hasDataRows) {
        rows.forEach((rowNode, i) => {
            const excelRow = firstDataRow + i;
            const name = uniqueDefinedName(usedNames, rowNode.tid, rowNode.ord);
            nameByKey.set(`${rowNode.ord}\0${rowNode.tid}`, name);
            addDefinedName(workbook, a1Range(PIVOT_SHEET, 1, excelRow, headerCols, excelRow), name);
        });
    } else if (hasDataCols) {
        cols.forEach((colNode, j) => {
            const excelCol = firstDataCol + j;
            const name = uniqueDefinedName(usedNames, colNode.tid, colNode.ord);
            nameByKey.set(`${colNode.ord}\0${colNode.tid}`, name);
            addDefinedName(workbook, a1Range(PIVOT_SHEET, excelCol, 2, excelCol, 3), name);
        });
    }

    const catalog = workbook.addWorksheet(CATALOG_SHEET, {
        views: [{ state: 'frozen', ySplit: 1, activeCell: 'A2' }],
    });
    const catHeaders = ['Измерение', 'ord', 'Код', 'Наименование', 'Описание', 'Родитель', 'Уровень', 'oper', 'range name'];
    catHeaders.forEach((title, idx) => styleHeaderCell(catalog.getCell(1, idx + 1), title));
    catalog.getColumn(1).width = 14;
    catalog.getColumn(2).width = 8;
    catalog.getColumn(3).width = 18;
    catalog.getColumn(4).width = 28;
    catalog.getColumn(5).width = 36;
    catalog.getColumn(6).width = 16;
    catalog.getColumn(7).width = 10;
    catalog.getColumn(8).width = 20;
    catalog.getColumn(9).width = 22;
    catalog.getRow(1).height = 20;

    const dims = engine.getDims();
    let catRow = 2;
    for (const dim of dims) {
        if (!dim) continue;
        const nodes = typeof engine.listDn === 'function'
            ? engine.listDn(dim, null, false)
            : [dim.root, ...dim.nodes.values()];
        for (const node of nodes) {
            if (!node) continue;
            const key = `${node.ord}\0${node.tid}`;
            let name = nameByKey.get(key);
            const tidCellRef = a1Range(CATALOG_SHEET, 3, catRow);
            if (!name) {
                name = uniqueDefinedName(usedNames, node.tid, node.ord);
                nameByKey.set(key, name);
                addDefinedName(workbook, tidCellRef, name);
            }
            const row = catalog.getRow(catRow);
            row.getCell(1).value = dimLabel(node.ord, dim);
            row.getCell(2).value = node.ord;
            row.getCell(3).value = node.tid;
            row.getCell(4).value = node.caption || '';
            row.getCell(5).value = node.title || '';
            row.getCell(6).value = parentTid(node);
            row.getCell(7).value = node.lev;
            row.getCell(8).value = nodeOp(node);
            row.getCell(9).value = name;
            row.outlineLevel = outlineLev(node);
            row.height = 16;
            for (let c = 1; c <= 9; c++) {
                const cell = row.getCell(c);
                cell.font = { name: 'Calibri', size: 10, italic: isCollapsed(node) };
                cell.border = THIN_BORDER;
                cell.alignment = { vertical: 'middle', indent: c === 4 ? Math.max(0, (node.lev || 1) - 1) : 0 };
            }
            catRow += 1;
        }
    }
    catalog.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: Math.max(1, catRow - 1), column: 9 },
    };
    catalog.properties.outlineProperties = {
        summaryBelow: false,
        summaryRight: false,
    };

    return workbook;
}

const SOURCE_ORDS_SHEET = 'Измерения';
const SOURCE_ROWS_SHEET = 'Данные';
const ORD_COLS = ['ord', 'tid', 'caption', 'title', 'oper'];

function fillHeaderRow(sheet, titles) {
    titles.forEach((title, idx) => styleHeaderCell(sheet.getCell(1, idx + 1), title));
    sheet.getRow(1).height = 20;
}

function styleBodyCell(cell, value, align = 'left') {
    cell.value = value;
    cell.font = { name: 'Calibri', size: 10 };
    cell.border = THIN_BORDER;
    cell.alignment = { vertical: 'middle', horizontal: align };
}

function buildSourceWorkbook(ExcelJS, ords, rows) {
    if (!ExcelJS || typeof ExcelJS.Workbook !== 'function') {
        throw new Error('ExcelJS is not loaded');
    }
    if (!Array.isArray(ords) || !Array.isArray(rows)) {
        throw new Error('source json is not ready');
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'MOLAP';
    workbook.created = new Date();

    const ordsSheet = workbook.addWorksheet(SOURCE_ORDS_SHEET, {
        views: [{ state: 'frozen', ySplit: 1, activeCell: 'A2' }],
    });
    fillHeaderRow(ordsSheet, ORD_COLS);
    ordsSheet.getColumn(1).width = 8;
    ordsSheet.getColumn(2).width = 18;
    ordsSheet.getColumn(3).width = 28;
    ordsSheet.getColumn(4).width = 36;
    ordsSheet.getColumn(5).width = 48;
    ords.forEach((node, i) => {
        const excelRow = i + 2;
        const row = ordsSheet.getRow(excelRow);
        styleBodyCell(row.getCell(1), node.ord ?? '', 'center');
        styleBodyCell(row.getCell(2), node.tid ?? '');
        styleBodyCell(row.getCell(3), node.caption ?? '');
        styleBodyCell(row.getCell(4), node.title ?? '');
        styleBodyCell(row.getCell(5), operToText(node.oper));
        row.height = 16;
    });
    if (ords.length) {
        ordsSheet.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: ords.length + 1, column: ORD_COLS.length },
        };
    }

    let dimCount = 0;
    for (const rec of rows) {
        const n = rec && Array.isArray(rec.keys) ? rec.keys.length : 0;
        if (n > dimCount) dimCount = n;
    }
    const dimNames = Array.from({ length: dimCount }, (_, i) => `d${i}`);
    const rowsHeaders = dimNames.concat(['val']);
    const dataSheet = workbook.addWorksheet(SOURCE_ROWS_SHEET, {
        views: [{ state: 'frozen', ySplit: 1, activeCell: 'A2' }],
    });
    fillHeaderRow(dataSheet, rowsHeaders);
    rowsHeaders.forEach((title, idx) => {
        dataSheet.getColumn(idx + 1).width = title === 'val' ? 14 : 16;
    });
    rows.forEach((rec, i) => {
        const excelRow = i + 2;
        const row = dataSheet.getRow(excelRow);
        const keys = rec && Array.isArray(rec.keys) ? rec.keys : [];
        dimNames.forEach((_, d) => {
            styleBodyCell(row.getCell(d + 1), keys[d] ?? '');
        });
        const valCell = row.getCell(dimCount + 1);
        styleBodyCell(valCell, rec && rec.val != null ? rec.val : 0, 'right');
        valCell.numFmt = '#,##0';
        row.height = 16;
    });
    const lastDataRow = Math.max(2, rows.length + 1);
    if (rows.length) {
        dataSheet.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: lastDataRow, column: rowsHeaders.length },
        };
    }
    dimNames.forEach((name, d) => {
        const col = d + 1;
        addDefinedName(workbook, a1Range(SOURCE_ROWS_SHEET, col, 2, col, lastDataRow), name);
    });
    if (dimCount >= 0) {
        const valCol = dimCount + 1;
        addDefinedName(workbook, a1Range(SOURCE_ROWS_SHEET, valCol, 2, valCol, lastDataRow), 'val');
    }

    return workbook;
}

function cellScalar(value) {
    if (value == null || value === '') return '';
    if (typeof value === 'number' || typeof value === 'boolean') return value;
    if (typeof value === 'string') return value;
    if (value instanceof Date) {
        const y = value.getFullYear();
        const m = String(value.getMonth() + 1).padStart(2, '0');
        const d = String(value.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }
    if (Array.isArray(value.richText)) {
        return value.richText.map((part) => part.text || '').join('');
    }
    if (value.text != null) return value.text;
    if (value.result != null) return cellScalar(value.result);
    return String(value);
}

function headerMap(row) {
    const map = {};
    if (!row) return map;
    row.eachCell((cell, colNumber) => {
        const name = String(cellScalar(cell.value)).trim();
        if (name) map[name] = colNumber;
    });
    return map;
}

function sheetCell(sheet, rowNumber, colNumber) {
    if (!colNumber) return '';
    return cellScalar(sheet.getRow(rowNumber).getCell(colNumber).value);
}

function parseNumber(raw) {
    if (typeof raw === 'number' && Number.isFinite(raw)) return raw;
    if (raw == null || raw === '') return 0;
    const text = String(raw).replace(/\u00a0/g, ' ').replace(/\s/g, '').replace(',', '.');
    const num = Number(text);
    return Number.isFinite(num) ? num : 0;
}

function findSheetByHeaders(workbook, required) {
    for (const sheet of workbook.worksheets) {
        const map = headerMap(sheet.getRow(1));
        if (required.every((name) => name in map)) return { sheet, map };
    }
    return null;
}

function parseSourceWorkbook(workbook) {
    if (!workbook) throw new Error('workbook is empty');
    const ordsFound = findSheetByHeaders(workbook, ['ord', 'tid', 'oper'])
        || (workbook.getWorksheet(SOURCE_ORDS_SHEET) && {
            sheet: workbook.getWorksheet(SOURCE_ORDS_SHEET),
            map: headerMap(workbook.getWorksheet(SOURCE_ORDS_SHEET).getRow(1)),
        });
    const dataFound = findSheetByHeaders(workbook, ['val'])
        || (workbook.getWorksheet(SOURCE_ROWS_SHEET) && {
            sheet: workbook.getWorksheet(SOURCE_ROWS_SHEET),
            map: headerMap(workbook.getWorksheet(SOURCE_ROWS_SHEET).getRow(1)),
        });
    if (!ordsFound || !ordsFound.sheet) throw new Error('sheet Измерения not found');
    if (!dataFound || !dataFound.sheet) throw new Error('sheet Данные not found');

    const ords = [];
    const ordsSheet = ordsFound.sheet;
    const ordsMap = ordsFound.map;
    ordsSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const tid = String(sheetCell(ordsSheet, rowNumber, ordsMap.tid)).trim();
        if (!tid) return;
        const ordRaw = sheetCell(ordsSheet, rowNumber, ordsMap.ord);
        ords.push({
            ord: Number(ordRaw),
            tid,
            caption: String(sheetCell(ordsSheet, rowNumber, ordsMap.caption || 0)),
            title: String(sheetCell(ordsSheet, rowNumber, ordsMap.title || 0)),
            oper: sheetCell(ordsSheet, rowNumber, ordsMap.oper),
        });
    });

    const dimCols = Object.keys(dataFound.map)
        .filter((name) => /^d\d+$/.test(name))
        .sort((a, b) => Number(a.slice(1)) - Number(b.slice(1)));
    const rows = [];
    const dataSheet = dataFound.sheet;
    const dataMap = dataFound.map;
    dataSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const keys = dimCols.map((name) => String(sheetCell(dataSheet, rowNumber, dataMap[name])).trim());
        if (keys.every((key) => !key)) return;
        rows.push({
            keys,
            val: parseNumber(sheetCell(dataSheet, rowNumber, dataMap.val)),
        });
    });

    return { ords, rows };
}

export {
    PIVOT_SHEET,
    CATALOG_SHEET,
    SOURCE_ORDS_SHEET,
    SOURCE_ROWS_SHEET,
    sanitizeExcelName,
    uniqueDefinedName,
    buildPivotWorkbook,
    buildSourceWorkbook,
    parseSourceWorkbook,
};
