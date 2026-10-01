import ExcelJS from 'exceljs';
import {
  getAllComponents,
  getAllFgProducts,
  getAllProductionPlans,
  getAllProductionRuns,
  getAllStockMovements,
} from './repository.js';
import { buildComponentMetrics } from './calculations.js';
import { getMonthKeyFromDate } from './workingDays.js';

const pad = (value) => String(value).padStart(2, '0');

const localDateKey = (date = new Date()) =>
  `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;

const fmt = (value) => Number(value || 0).toLocaleString('en-IN');

export async function buildDailyExcelReport(dateKey = localDateKey()) {
  const components = getAllComponents();
  const products = getAllFgProducts();
  const productionRuns = getAllProductionRuns();
  const productionPlans = getAllProductionPlans();
  const stockMovements = getAllStockMovements();

  const runs = productionRuns.filter(
    (run) => localDateKey(new Date(run.date)) === dateKey
  );

  const movements = stockMovements.filter(
    (movement) =>
      ['INWARD', 'BULK_UPDATE'].includes(movement.type) &&
      (movement.receiptDate ||
        localDateKey(new Date(movement.createdAt))) === dateKey
  );

  const critical = buildComponentMetrics(
    components,
    products,
    productionRuns,
    productionPlans,
    getMonthKeyFromDate(dateKey)
  ).filter((row) => row.risk === 'critical');

  const workbook = new ExcelJS.Workbook();

  workbook.creator = 'INEL RM';
  workbook.created = new Date();
  workbook.modified = new Date();

  // --------------------------------------------------
  // Summary Sheet
  // --------------------------------------------------

  const summarySheet = workbook.addWorksheet('Summary', 0);

  summarySheet.columns = [
    { header: 'Report Information', key: 'label', width: 30 },
    { header: 'Value', key: 'value', width: 35 },
  ];

  summarySheet.addRows([
    ['Report Date', dateKey],
    ['Stock Inward Lines', movements.length],
    [
      'Total Inward Quantity',
      movements.reduce((sum, row) => sum + Number(row.qty || 0), 0),
    ],
    ['Production Confirmed', runs.length],
    [
      'Total Production Quantity',
      runs.reduce((sum, row) => sum + Number(row.qty || 0), 0),
    ],
    ['Critical Components', critical.length],
  ]);

  summarySheet.getRow(1).font = {
    bold: true,
  };

  summarySheet.getRow(1).alignment = {
    horizontal: 'center',
  };

  // --------------------------------------------------
  // Sheet 1: Stock Inward
  // --------------------------------------------------

  const inwardSheet = workbook.addWorksheet('Stock Inward');

  inwardSheet.columns = [
    { header: 'Receipt / Batch', key: 'receiptBatch', width: 20 },
    { header: 'Material', key: 'material', width: 20 },
    { header: 'Description', key: 'description', width: 35 },
    { header: 'Qty', key: 'qty', width: 15 },
    { header: 'Vendor', key: 'vendor', width: 25 },
    { header: 'Invoice', key: 'invoice', width: 20 },
  ];

  movements.forEach((movement) => {
    inwardSheet.addRow({
      receiptBatch: movement.transactionId || movement.id,
      material: movement.componentId,
      description: movement.componentDesc,
      qty: Number(movement.qty || 0),
      vendor: movement.vendor || '-',
      invoice: movement.invoiceNo || '-',
    });
  });

  // --------------------------------------------------
  // Sheet 2: Production
  // --------------------------------------------------

  const productionSheet = workbook.addWorksheet('Production');

  productionSheet.columns = [
    { header: 'Run ID', key: 'runId', width: 18 },
    { header: 'Confirmed Time', key: 'confirmedTime', width: 22 },
    { header: 'FG Part', key: 'fgPart', width: 20 },
    { header: 'FG Name', key: 'fgName', width: 35 },
    { header: 'FG Qty', key: 'fgQty', width: 15 },
    { header: 'Components', key: 'components', width: 15 },
  ];

  runs.forEach((run) => {
    productionSheet.addRow({
      runId: `PRD-${String(run.id).padStart(6, '0')}`,
      confirmedTime: new Date(run.date).toLocaleString('en-IN'),
      fgPart: run.partNo,
      fgName: run.name,
      fgQty: Number(run.qty || 0),
      components: (run.items || []).length,
    });
  });

  // --------------------------------------------------
  // Sheet 3: Critical Components
  // --------------------------------------------------

  const criticalSheet = workbook.addWorksheet('Critical Components');

  criticalSheet.columns = [
    { header: 'Component', key: 'component', width: 20 },
    { header: 'Description', key: 'description', width: 35 },
    { header: 'Vendor', key: 'vendor', width: 25 },
    { header: 'Stock', key: 'stock', width: 15 },
    { header: 'Required', key: 'required', width: 15 },
    { header: 'Shortage', key: 'shortage', width: 15 },
    { header: 'Coverage', key: 'coverage', width: 15 },
  ];

  critical.forEach((row) => {
    criticalSheet.addRow({
      component: row.id,
      description: row.desc,
      vendor: row.vendor || '-',
      stock: Number(row.totalStock || 0),
      required: Number(row.demandRequirement || 0),
      shortage: Number(row.shortage || 0),
      coverage: row.coverageDays ?? '-',
    });
  });

  // --------------------------------------------------
  // Formatting
  // --------------------------------------------------

  workbook.worksheets.forEach((sheet) => {
    const headerRow = sheet.getRow(1);

    headerRow.font = {
      bold: true,
    };

    headerRow.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };

    headerRow.height = 22;

    sheet.views = [
      {
        state: 'frozen',
        ySplit: 1,
      },
    ];

    sheet.autoFilter = {
      from: 'A1',
      to: `${String.fromCharCode(64 + sheet.columnCount)}1`,
    };

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.alignment = {
          vertical: 'top',
        };
      }
    });
  });

 return {
    buffer: await workbook.xlsx.writeBuffer(),
    filename: `INEL_RM_Evening_Digest_${dateKey}.xlsx`,
  };
}