import { readWorkbookBuffer, mapHeaders, normalizeHeader } from "./readWorkbook.js";
import XLSX from "xlsx";

const number = (value) => Number(String(value ?? "").replace(/,/g, "")) || 0;
const text = (value) => String(value ?? "").trim();

function parseDateValue(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  const textValue = String(value).trim();
  if (!textValue) return null;

  const match = textValue.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (match) {
    const [, day, month, year] = match;
    const parsed = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  const iso = new Date(textValue);
  if (!Number.isNaN(iso.getTime())) return iso;

  return null;
}

function trackMonitoringDate(name, rows) {
  for (const row of rows.slice(0, 6)) {
    for (let index = 0; index < row.length; index += 1) {
      const cellValue = row[index];
      const label = normalizeHeader(cellValue);
      if (!label || !/(updatedon|monitoringdate|date)/.test(label)) continue;
      const candidate = row[index + 1];
      const parsed = parseDateValue(candidate);
      if (parsed) return parsed.toISOString().slice(0, 10);
    }
  }

  const month = rows[0]?.find((cell) => /[a-z]{3}-\d{2}/i.test(text(cell)));
  const match = text(name).match(/^(\d{2})(\d{2})$/);
  if (month && match) {
    const parsed = new Date(`20${text(month).slice(-2)}-${text(month).slice(0, 3)}-${match[1]}`);
    if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  }

  return new Date().toISOString().slice(0, 10);
}

export function parseFgStockWorkbook(buffer) {
  const wb = readWorkbookBuffer(buffer);
  const records = [], errors = [];
  for (const sheetName of wb.SheetNames) {
    const rows = wb.Sheets[sheetName]
      ? XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null })
      : [];
    const headerIndex = rows.findIndex((row) => row.some((cell) => ["partno", "fgpartno", "fgpartnumber"].includes(normalizeHeader(cell))));
    if (headerIndex < 0) continue;
    const headers = mapHeaders(rows[headerIndex]);
    const get = (row, ...keys) => {
      for (const key of keys) {
        const index = headers[normalizeHeader(key)];
        if (index !== undefined && row[index] !== null && row[index] !== "") return row[index];
      }
      return "";
    };
    const requiredHeaders = ["Part No", "Month Plan Qty", "Stock Norms In days", "Total Stock"];
    const missingHeaders = requiredHeaders.filter((header) => headers[normalizeHeader(header)] === undefined);
    if (missingHeaders.length) {
      errors.push(`${sheetName}: missing required column(s): ${missingHeaders.join(", ")}`);
      continue;
    }
    const monitorDate = trackMonitoringDate(sheetName, rows);
    for (let index = headerIndex + 1; index < rows.length; index++) {
      const row = rows[index];
      const partNo = text(get(row, "Part No", "FG Part No", "FG Part Number"));
      if (!partNo) continue;
      const planValue = get(row, "Month Plan Qty", "Monthly Plan", "Plan Qty");
      const totalStockValue = get(row, "Total Stock");
      const plan = number(planValue);
      const totalStock = number(totalStockValue);
      const customer = text(get(row, "Customer Name", "Customer"));
      const stockNormValue = get(row, "Stock Norms In days", "Stock Norm Days");
      const stockNormDays = number(stockNormValue);
      if (text(planValue) === "" || text(stockNormValue) === "" || text(totalStockValue) === "") {
        errors.push(`${sheetName} row ${index + 1}: Part No, Month Plan Qty, Stock Norms In Days, and Total Stock are required`);
        continue;
      }
      records.push({
        monitorDate, partNo, monthKey: monitorDate.slice(0, 7), monthPlanQty: plan,
        adr: number(get(row, "ADR")), stockNormDays, totalStock,
        name: text(get(row, "FG Name", "Name", "Description")), customer,
        customerCategory: text(get(row, "Catageory", "Customer Category")), productGroup: text(get(row, "Group", "Product Group")),
        fgCategory: text(get(row, "Category")),
      });
    }
  }
  return { records, errors };
}
