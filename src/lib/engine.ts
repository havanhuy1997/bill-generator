import ExcelJS from "exceljs";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { amountInWords, roundVnd } from "./numbers";
import type { ImportResult, Invoice } from "./types";

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const xml = (value: string) =>
  new DOMParser().parseFromString(value, "application/xml");
const serialize = (doc: ReturnType<typeof xml>) =>
  new XMLSerializer().serializeToString(doc);
type XmlDoc = ReturnType<typeof xml>;
type XmlElement = ReturnType<XmlDoc["createElement"]>;

function cellValue(cell: ExcelJS.Cell): unknown {
  const value = cell.value;
  if (value && typeof value === "object" && !(value instanceof Date)) {
    if ("formula" in value || "sharedFormula" in value) {
      // ExcelJS's value getter drops falsy results (notably zero); the result
      // accessor reads the saved model value without that lossy copy.
      const result = cell.result;
      if (result === undefined) {
        throw new Error(
          `Ô ${cell.address} có công thức nhưng chưa lưu kết quả. Hãy mở tệp trong Excel, tính toán lại, lưu tệp rồi thử lại.`,
        );
      }
      if (result && typeof result === "object" && "error" in result)
        throw new Error(
          `Ô ${cell.address} chứa lỗi ${result.error}. Hãy sửa tệp nguồn rồi thử lại.`,
        );
      return result;
    }
    if ("richText" in value)
      return value.richText.map((part) => part.text).join("");
    if ("text" in value) return value.text;
    if ("error" in value)
      throw new Error(
        `Ô ${cell.address} chứa lỗi ${value.error}. Hãy sửa tệp nguồn rồi thử lại.`,
      );
  }
  return value;
}

function text(value: unknown): string {
  return value == null ? "" : String(value).trim();
}

function numeric(
  value: unknown,
  row: number,
  column: string,
  optional = false,
): number {
  if ((value == null || value === "") && optional) return 0;
  if (
    (typeof value !== "number" && typeof value !== "string") ||
    text(value) === ""
  ) {
    throw new Error(
      `Dòng ${row}, cột ${column}: vui lòng nhập giá trị số trong Excel.`,
    );
  }
  if (typeof value === "string" && !/^[+-]?\d+(\.\d+)?$/.test(value.trim())) {
    throw new Error(
      `Dòng ${row}, cột ${column}: “${value}” không phải là số. Hãy dùng ô dạng số, không kèm ký hiệu tiền tệ hoặc dấu phân cách hàng nghìn.`,
    );
  }
  const result = Number(value);
  if (!Number.isFinite(result) || Math.abs(result) > Number.MAX_SAFE_INTEGER) {
    throw new Error(
      `Dòng ${row}, cột ${column}: giá trị nằm ngoài phạm vi hỗ trợ.`,
    );
  }
  return result;
}

export async function parseWorkbook(
  buffer: ArrayBuffer,
): Promise<ImportResult> {
  let workbook: ExcelJS.Workbook;
  try {
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
  } catch {
    throw new Error(
      "Không thể đọc tệp. Vui lòng chọn tệp .xlsx hợp lệ, không được mã hóa hoặc bảo vệ bằng mật khẩu.",
    );
  }
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("Tệp Excel không có trang tính nào.");
  const grouped = new Map<string, Invoice>();
  const warnings = new Set<string>();
  let skippedRows = 0;
  let rowCount = 0;
  sheet.eachRow((row, index) => {
    const code = text(cellValue(row.getCell(8)));
    if (code.toLocaleLowerCase("vi-VN") === "số phiếu ckho4") return;
    if (!code) {
      if (
        Array.from({ length: 8 }, (_, i) => row.getCell(i + 2).value).some(
          (v) => v !== null && v !== "",
        )
      )
        skippedRows++;
      return;
    }
    const name = text(cellValue(row.getCell(2)));
    if (!name)
      throw new Error(`Dòng ${index}, cột B: vui lòng nhập tên sản phẩm.`);
    const quantity = numeric(cellValue(row.getCell(4)), index, "D (số lượng)");
    const price = numeric(cellValue(row.getCell(5)), index, "E (đơn giá)");
    const adjustment = numeric(
      cellValue(row.getCell(6)),
      index,
      "F (giảm giá)",
      true,
    );
    if (quantity < 0 || price < 0)
      throw new Error(
        `Dòng ${index}: số lượng và đơn giá phải lớn hơn hoặc bằng 0.`,
      );
    if (!Number.isInteger(adjustment))
      throw new Error(
        `Dòng ${index}, cột F: vui lòng nhập số tiền giảm giá là số nguyên (VND).`,
      );
    const total = roundVnd(quantity * price) + adjustment;
    if (!Number.isSafeInteger(total))
      throw new Error(`Dòng ${index}: thành tiền quá lớn.`);
    const customer = text(cellValue(row.getCell(7)));
    const rawDate = cellValue(row.getCell(9));
    let excelDate: number | string | null = null;
    let date = "";
    if (rawDate instanceof Date) {
      excelDate = rawDate.getTime() / 86400000 + 25569;
      date = new Intl.DateTimeFormat("vi-VN", { timeZone: "UTC" }).format(
        rawDate,
      );
    } else if (rawDate != null) {
      excelDate =
        typeof rawDate === "number"
          ? rawDate + (workbook.properties.date1904 ? 1462 : 0)
          : text(rawDate);
      date =
        typeof excelDate === "number"
          ? new Intl.DateTimeFormat("vi-VN", { timeZone: "UTC" }).format(
              new Date((excelDate - 25569) * 86400000),
            )
          : excelDate;
    }
    if (!customer) warnings.add(`Hóa đơn ${code} chưa có tên khách hàng.`);
    if (!date) warnings.add(`Hóa đơn ${code} chưa có ngày lập.`);
    if (adjustment > 0)
      warnings.add(
        `Hóa đơn ${code} có số tiền giảm giá dương. Giá trị này sẽ được cộng vào tổng tiền. Hãy nhập số âm để giảm số tiền phải trả.`,
      );
    const product = {
      name,
      unit: text(cellValue(row.getCell(3))),
      quantity,
      price,
      adjustment,
      total,
    };
    const invoice = grouped.get(code);
    if (invoice) {
      if (invoice.customer !== customer || invoice.date !== date)
        warnings.add(
          `Hóa đơn ${code} có thông tin khách hàng hoặc ngày lập không đồng nhất. Ứng dụng sẽ dùng thông tin ở dòng đầu tiên.`,
        );
      invoice.products.push(product);
      invoice.total += total;
    } else
      grouped.set(code, {
        code,
        customer,
        date,
        excelDate,
        products: [product],
        total,
      });
    rowCount++;
  });
  const invoices = [...grouped.values()];
  if (!invoices.length)
    throw new Error(
      "Không tìm thấy dòng hóa đơn nào. Hãy sắp xếp các cột B–I như tệp mẫu và nhập mã hóa đơn vào cột H.",
    );
  for (const invoice of invoices) amountInWords(invoice.total);
  if (skippedRows)
    warnings.add(`Sẽ bỏ qua ${skippedRows} dòng không có mã hóa đơn ở cột H.`);
  return { invoices, rowCount, skippedRows, warnings: [...warnings] };
}

function getCell(doc: XmlDoc, ref: string): XmlElement {
  for (const cell of Array.from(doc.getElementsByTagName("c"))) {
    if (cell.getAttribute("r") === ref) return cell;
  }
  const rowIndex = ref.match(/\d+$/)![0];
  const row = Array.from(doc.getElementsByTagName("row")).find(
    (r) => r.getAttribute("r") === rowIndex,
  );
  if (!row) throw new Error(`Mẫu hóa đơn bị thiếu dòng ${rowIndex}.`);
  const cell = doc.createElementNS(NS, "c");
  cell.setAttribute("r", ref);
  const column = ref.replace(/\d+$/, "");
  const after = Array.from(row.childNodes).find(
    (n) =>
      n.nodeType === 1 &&
      (n as XmlElement).getAttribute("r")!.replace(/\d+$/, "") > column,
  );
  row.insertBefore(cell, after ?? null);
  return cell;
}

function writeCell(
  doc: XmlDoc,
  ref: string,
  value: number | string | null,
  formula?: string,
): void {
  const cell = getCell(doc, ref);
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  cell.removeAttribute("t");
  if (formula) {
    const f = doc.createElementNS(NS, "f");
    f.appendChild(doc.createTextNode(formula));
    cell.appendChild(f);
  }
  if (value === null) return;
  if (typeof value === "number") {
    const v = doc.createElementNS(NS, "v");
    v.appendChild(doc.createTextNode(String(value)));
    cell.appendChild(v);
  } else {
    cell.setAttribute("t", "inlineStr");
    const inline = doc.createElementNS(NS, "is");
    const t = doc.createElementNS(NS, "t");
    t.setAttribute("xml:space", "preserve");
    t.appendChild(doc.createTextNode(value));
    inline.appendChild(t);
    cell.appendChild(inline);
  }
}

function shiftRef(ref: string, extra: number): string {
  return ref.replace(
    /(\$?[A-Z]{1,3}\$?)(\d+)/g,
    (_, col: string, row: string) =>
      col + (Number(row) >= 22 ? Number(row) + extra : row),
  );
}

function extendRows(doc: XmlDoc, extra: number): void {
  if (!extra) return;
  const rows = Array.from(doc.getElementsByTagName("row"));
  const prototype = rows.find((row) => row.getAttribute("r") === "21")!;
  for (const row of rows) {
    const index = Number(row.getAttribute("r"));
    if (index >= 22) row.setAttribute("r", String(index + extra));
  }
  for (const el of Array.from(doc.getElementsByTagName("*"))) {
    for (const attr of ["r", "ref", "sqref"]) {
      if (el.hasAttribute(attr) && el.tagName !== "row")
        el.setAttribute(attr, shiftRef(el.getAttribute(attr)!, extra));
    }
    if (el.tagName === "f" && el.textContent)
      el.textContent = shiftRef(el.textContent, extra);
  }
  const sheetData = prototype.parentNode!;
  const next = prototype.nextSibling;
  const merges = doc.getElementsByTagName("mergeCells")[0];
  for (let i = 1; i <= extra; i++) {
    const row = prototype.cloneNode(true) as XmlElement;
    row.setAttribute("r", String(21 + i));
    for (const cell of Array.from(row.getElementsByTagName("c")))
      cell.setAttribute(
        "r",
        cell.getAttribute("r")!.replace(/21$/, String(21 + i)),
      );
    sheetData.insertBefore(row, next);
    const merge = doc.createElementNS(NS, "mergeCell");
    merge.setAttribute("ref", `B${21 + i}:C${21 + i}`);
    merges.appendChild(merge);
  }
  merges.setAttribute(
    "count",
    String(merges.getElementsByTagName("mergeCell").length),
  );
}

export async function renderInvoice(
  template: ArrayBuffer,
  invoice: Invoice,
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(template);
  const doc = xml(await zip.file("xl/worksheets/sheet1.xml")!.async("string"));
  const extra = Math.max(0, invoice.products.length - 10);
  extendRows(doc, extra);
  writeCell(doc, "C6", invoice.excelDate);
  writeCell(doc, "C7", invoice.customer);
  writeCell(doc, `F${29 + extra}`, invoice.code);
  for (let i = 0; i < Math.max(10, invoice.products.length); i++) {
    const row = i + 12;
    const p = invoice.products[i];
    writeCell(doc, `A${row}`, p ? i + 1 : null);
    writeCell(doc, `B${row}`, p?.name ?? null);
    writeCell(doc, `D${row}`, p?.unit ?? null);
    writeCell(doc, `E${row}`, p?.quantity ?? null);
    writeCell(doc, `F${row}`, p?.price ?? null);
    writeCell(doc, `G${row}`, p?.adjustment ?? null);
    // Keep recalculation consistent with the original Python ties-to-even rounding.
    const product = `E${row}*F${row}`;
    writeCell(
      doc,
      `H${row}`,
      p?.total ?? 0,
      `IF(MOD(${product},1)=0.5,ROUND(${product}/2,0)*2,ROUND(${product},0))+G${row}`,
    );
  }
  const sumRow = 22 + extra;
  const last = 21 + extra;
  writeCell(
    doc,
    `E${sumRow}`,
    invoice.products.reduce((sum, p) => sum + p.quantity, 0),
    `SUM(E12:E${last})`,
  );
  writeCell(
    doc,
    `G${sumRow}`,
    invoice.products.reduce((sum, p) => sum + p.adjustment, 0),
    `SUM(G12:G${last})`,
  );
  writeCell(doc, `H${sumRow}`, invoice.total, `SUM(H12:H${last})`);
  writeCell(doc, `A${24 + extra}`, amountInWords(invoice.total));
  zip.file("xl/worksheets/sheet1.xml", serialize(doc));

  const workbook = xml(await zip.file("xl/workbook.xml")!.async("string"));
  const calc = workbook.getElementsByTagName("calcPr")[0];
  if (calc) {
    calc.setAttribute("fullCalcOnLoad", "1");
    calc.setAttribute("forceFullCalc", "1");
  }
  if (extra) {
    for (const name of Array.from(workbook.getElementsByTagName("definedName")))
      name.textContent = shiftRef(name.textContent ?? "", extra);
    const drawingPath = "xl/drawings/drawing1.xml";
    const drawing = xml(await zip.file(drawingPath)!.async("string"));
    for (const row of Array.from(drawing.getElementsByTagName("xdr:row"))) {
      if (Number(row.textContent) >= 21)
        row.textContent = String(Number(row.textContent) + extra);
    }
    zip.file(drawingPath, serialize(drawing));
  }
  zip.file("xl/workbook.xml", serialize(workbook));
  // Formula positions and shared formula groups have changed; Excel rebuilds the chain.
  zip.remove("xl/calcChain.xml");
  for (const path of ["xl/_rels/workbook.xml.rels", "[Content_Types].xml"]) {
    const part = xml(await zip.file(path)!.async("string"));
    for (const el of Array.from(part.getElementsByTagName("*"))) {
      if (
        el.getAttribute("Target") === "calcChain.xml" ||
        el.getAttribute("PartName") === "/xl/calcChain.xml"
      )
        el.parentNode?.removeChild(el);
    }
    zip.file(path, serialize(part));
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export function safeFilename(code: string): string {
  const name = code
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, "_")
    .replace(/[. ]+$/g, "")
    .slice(0, 150);
  return !name || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name)
    ? `invoice_${name}`
    : name;
}

export async function generateArchive(
  template: ArrayBuffer,
  invoices: Invoice[],
  progress?: (done: number, total: number) => void,
): Promise<ArrayBuffer> {
  const zip = new JSZip();
  const filenames = new Set<string>();
  for (let i = 0; i < invoices.length; i++) {
    const base = safeFilename(invoices[i].code);
    let name = base;
    let suffix = 2;
    while (filenames.has(name.toLowerCase())) name = `${base} (${suffix++})`;
    filenames.add(name.toLowerCase());
    zip.file(`${name}.xlsx`, await renderInvoice(template, invoices[i]));
    progress?.(i + 1, invoices.length);
  }
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

export async function createSample(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Đơn hàng");
  sheet.addRow([
    "STT",
    "Tên hàng hóa",
    "ĐVT",
    "Số lượng",
    "Đơn giá",
    "Giảm giá",
    "Thông tin khách hàng",
    "Số phiếu CKHO4",
    "Ngày",
  ]);
  const date = new Date("2026-11-14T00:00:00Z");
  const rows = [
    [
      1,
      "Loa Bluetooth mini",
      "Cái",
      2,
      350000,
      -50000,
      "Nguyễn Minh Anh",
      "CKHO4/26/11/001",
      date,
    ],
    [
      2,
      "Cáp sạc USB-C",
      "Cái",
      3,
      120000,
      0,
      "Nguyễn Minh Anh",
      "CKHO4/26/11/001",
      date,
    ],
    [
      3,
      "Tai nghe không dây",
      "Bộ",
      1,
      890000,
      -90000,
      "Trần Hoàng Nam",
      "CKHO4/26/11/002",
      date,
    ],
    [
      4,
      "Sạc nhanh 30W",
      "Cái",
      2,
      250000,
      -50000,
      "Lê Thu Hà",
      "CKHO4/26/11/003",
      date,
    ],
    [
      5,
      "Ốp điện thoại",
      "Cái",
      1,
      150000,
      0,
      "Lê Thu Hà",
      "CKHO4/26/11/003",
      date,
    ],
  ];
  rows.forEach((row) => sheet.addRow(row));
  sheet.columns.forEach((column, i) => {
    column.width = [7, 30, 10, 12, 16, 16, 26, 26, 16][i];
  });
  sheet.getColumn(9).numFmt = "dd/mm/yyyy";
  for (const col of [5, 6]) sheet.getColumn(col).numFmt = "#,##0";
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF22694F" },
  };
  sheet.getRow(1).height = 28;
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  const bytes = await workbook.xlsx.writeBuffer();
  return Uint8Array.from(new Uint8Array(bytes)).buffer;
}
