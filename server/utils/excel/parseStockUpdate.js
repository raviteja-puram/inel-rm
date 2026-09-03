import { readWorkbookBuffer, sheetToRows, mapHeaders, cell } from './readWorkbook.js';

export function parseStockUpdateFlat(buffer) {
  const wb = readWorkbookBuffer(buffer);
  const { rows } = sheetToRows(wb);
  const hasHeader = rows[0] && isNaN(Number(rows[0][1]));
  const start = hasHeader ? 1 : 0;
  const map = hasHeader ? mapHeaders(rows[0]) : {};
  const hasStockHeader = !hasHeader || [
    'stockdelta', 'stock delta', 'receivedqty', 'received qty', 'receiptqty', 'receipt qty',
    'updatestock', 'update stock', 'addstock', 'add stock',
    'stockvalue', 'stock value', 'stock', 'stocks', 'totalstock', 'total stock', 'mm01', 'storestockmm01',
    'mm10', 'storestockmm10', 'wip', 'pp00', 'qr01', 'inspection', 'qad',
    'ibpending', 'plannedinward', 'inward',
  ].some((key) => map[key.replace(/[^a-z0-9]+/gi, '').toLowerCase()] !== undefined);
  const items = [], errors = [], seenIds = new Set();
  if (!hasStockHeader) {
    return {
      items: [],
      errors: ['Missing stock column. Use Total Stock to replace, or Update Stock to add.'],
    };
  }
  for (let i = start; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every((v) => v === null || v === '')) continue;
    const id = String(cell(row, map, 'componentid', 'component id', 'id') ?? row[0] ?? '').trim();
    if (!id) { errors.push(`Row ${i+1}: missing Component ID`); continue; }
    const idKey = id.toUpperCase();
    if (seenIds.has(idKey)) {
      errors.push(`Row ${i+1}: duplicate Component ID ${id} skipped`);
      continue;
    }
    const stockDelta = Number(String(cell(
      row,
      map,
      'stockdelta',
      'stock delta',
      'receivedqty',
      'received qty',
      'receiptqty',
      'receipt qty',
      'updatestock',
      'update stock',
      'addstock',
      'add stock',
    ) ?? '').replace(/,/g, ''));
    if (Number.isFinite(stockDelta) && stockDelta > 0) {
      seenIds.add(idKey);
      items.push({ id, stockDelta });
      continue;
    }
    const rawStockValue =
      cell(
        row,
        map,
        'stockvalue',
        'stock value',
        'stock',
        'stocks',
        'totalstock',
        'total stock',
      ) ?? (hasHeader ? null : row[2] ?? row[1]);
    const stockValue = Number(String(rawStockValue ?? '').replace(/,/g, ''));
    if (
      rawStockValue !== null &&
      rawStockValue !== '' &&
      Number.isFinite(stockValue)
    ) {
      seenIds.add(idKey);
      items.push({ id, stock: stockValue });
      continue;
    }
    const multiBucketUpdate = {
      id,
      storeStockMm01: Number(cell(row, map, 'mm01', 'storestockmm01') ?? row[1]) || 0,
      storeStockMm10: Number(cell(row, map, 'mm10', 'storestockmm10') ?? row[2]) || 0,
      wipStock: Number(cell(row, map, 'wip', 'pp00') ?? row[3]) || 0,
      qr01Stock: Number(cell(row, map, 'qr01') ?? row[4]) || 0,
      inspectionStock: Number(cell(row, map, 'inspection', 'qad') ?? row[5]) || 0,
      ibPending: Number(cell(row, map, 'ibpending') ?? row[6]) || 0,
      inwardQty: Number(cell(row, map, 'plannedinward', 'inward') ?? row[7]) || 0,
    };
    const total = multiBucketUpdate.storeStockMm01 + multiBucketUpdate.storeStockMm10
      + multiBucketUpdate.wipStock + multiBucketUpdate.qr01Stock
      + multiBucketUpdate.inspectionStock + multiBucketUpdate.ibPending
      + multiBucketUpdate.inwardQty;
    if (total === 0) {
      errors.push(`Row ${i+1}: missing Stock Value`);
      continue;
    }
    seenIds.add(idKey);
    items.push({ id, stock: total });
  }
  return { items, errors };
}
