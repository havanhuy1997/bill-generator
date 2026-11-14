import { describe, expect, it } from "vitest";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import { amountInWords, roundVnd } from "../src/lib/numbers";
import {
  createSample,
  generateArchive,
  parseWorkbook,
  renderInvoice,
  safeFilename,
} from "../src/lib/engine";
import type { Invoice } from "../src/lib/types";

const template = async () =>
  Uint8Array.from(await readFile("public/templates/invoice.xlsx")).buffer;
const fixtureRow = (
  code: unknown = "CKHO4/26/001",
  overrides: Record<number, unknown> = {},
) => {
  const row: unknown[] = [
    1,
    "Sản phẩm thử",
    "Cái",
    2,
    350000,
    -50000,
    "Nguyễn Minh Anh",
    code,
    new Date("2026-11-14T00:00:00Z"),
  ];
  for (const [column, value] of Object.entries(overrides))
    row[Number(column) - 1] = value;
  return row;
};
async function input(rows: unknown[][], date1904 = false) {
  const wb = new ExcelJS.Workbook();
  wb.properties.date1904 = date1904;
  const sheet = wb.addWorksheet("Orders");
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
  rows.forEach((row) => sheet.addRow(row));
  sheet.getColumn(9).numFmt = "dd/mm/yyyy";
  return Uint8Array.from(new Uint8Array(await wb.xlsx.writeBuffer())).buffer;
}
const sheetDoc = async (bytes: Uint8Array) =>
  new DOMParser().parseFromString(
    await (await JSZip.loadAsync(bytes))
      .file("xl/worksheets/sheet1.xml")!
      .async("string"),
    "application/xml",
  );
const read = async (bytes: Uint8Array) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Uint8Array.from(bytes).buffer);
  return wb.worksheets[0];
};

describe("original import contract", () => {
  it("groups non-adjacent codes, reads the first sheet, and skips missing codes", async () => {
    const buffer = await input([
      fixtureRow("A"),
      fixtureRow("B"),
      fixtureRow("A"),
      fixtureRow(null),
    ]);
    const parsed = await parseWorkbook(buffer);
    expect(
      parsed.invoices.map((i) => [i.code, i.products.length, i.total]),
    ).toEqual([
      ["A", 2, 1300000],
      ["B", 1, 650000],
    ]);
    expect(parsed.rowCount).toBe(3);
    expect(parsed.skippedRows).toBe(1);
    expect(parsed.warnings[0]).toContain("không có mã hóa đơn");
    expect(parsed.invoices[0].date).toBe("14/11/2026");
  });

  it("accepts numeric invoice codes", async () => {
    const parsed = await parseWorkbook(await input([fixtureRow(12345)]));
    expect(parsed.invoices[0].code).toBe("12345");
  });

  it("uses cached formula results, including zero", async () => {
    const parsed = await parseWorkbook(
      await input([
        fixtureRow("A", {
          4: { formula: "1+1", result: 2 },
          6: { formula: "0+0", result: 0 },
        }),
      ]),
    );
    expect(parsed.invoices[0].total).toBe(700000);
  });

  it("rejects formulas with no cached result and Excel errors", async () => {
    await expect(
      parseWorkbook(await input([fixtureRow("A", { 5: { formula: "1+1" } })])),
    ).rejects.toThrow("chưa lưu kết quả");
    await expect(
      parseWorkbook(
        await input([fixtureRow("A", { 5: { error: "#VALUE!" } })]),
      ),
    ).rejects.toThrow("#VALUE!");
  });

  it("rejects malformed numeric values instead of generating wrong amounts", async () => {
    for (const value of [null, "350,000", "350000 VND", "abc"])
      await expect(
        parseWorkbook(await input([fixtureRow("A", { 5: value })])),
      ).rejects.toThrow("Dòng 2");
    await expect(
      parseWorkbook(await input([fixtureRow("A", { 4: -2 })])),
    ).rejects.toThrow("lớn hơn hoặc bằng 0");
  });

  it("defaults blank discounts to zero, adds signed adjustments, and flags positive ones", async () => {
    const parsed = await parseWorkbook(
      await input([fixtureRow("A", { 6: null }), fixtureRow("B", { 6: 1000 })]),
    );
    expect(parsed.invoices.map((i) => i.total)).toEqual([700000, 701000]);
    expect(parsed.warnings[0]).toContain("giảm giá dương");
  });

  it("preserves first-row customer and date and reports inconsistencies", async () => {
    const parsed = await parseWorkbook(
      await input([
        fixtureRow("A"),
        fixtureRow("A", { 7: "Different customer" }),
      ]),
    );
    expect(parsed.invoices[0].customer).toBe("Nguyễn Minh Anh");
    expect(parsed.warnings[0]).toContain("dòng đầu tiên");
  });

  it("normalizes 1904 workbooks to the template’s 1900 date system", async () => {
    const normal = await parseWorkbook(await input([fixtureRow()]));
    const mac = await parseWorkbook(await input([fixtureRow()], true));
    expect(mac.invoices[0].excelDate).toBe(normal.invoices[0].excelDate);
  });

  it("reports invalid workbooks, no invoices, and excessive numeric totals", async () => {
    await expect(
      parseWorkbook(new TextEncoder().encode("not excel").buffer),
    ).rejects.toThrow(".xlsx hợp lệ");
    await expect(parseWorkbook(await input([]))).rejects.toThrow(
      "Không tìm thấy dòng hóa đơn",
    );
    await expect(
      parseWorkbook(await input([fixtureRow("A", { 4: 1e12, 5: 1e12 })])),
    ).rejects.toThrow("quá lớn");
  });
});

describe("amounts and Vietnamese words", () => {
  it("matches the original Python ties-to-even rule", () => {
    expect([0.5, 1.5, 2.5, 3.5, 4.5, 2.49, 2.51].map(roundVnd)).toEqual([
      0, 2, 2, 4, 4, 2, 3,
    ]);
  });
  it("reads invoice amounts including embedded zero groups", () => {
    expect(amountInWords(0)).toBe("Không đồng.");
    expect(amountInWords(101)).toBe("Một trăm lẽ một đồng.");
    expect(amountInWords(121)).toBe("Một trăm hai mươi mốt đồng.");
    expect(amountInWords(805)).toBe("Tám trăm lẽ năm đồng.");
    expect(amountInWords(1010001)).toBe(
      "Một triệu không trăm mười nghìn không trăm lẽ một đồng.",
    );
    expect(amountInWords(115205201211)).toBe(
      "Một trăm mười lăm tỷ hai trăm lẽ năm triệu hai trăm lẽ một nghìn hai trăm mười một đồng.",
    );
    expect(amountInWords(-5000)).toBe("Âm năm nghìn đồng.");
  });
});

describe("original branded template output", () => {
  it("writes invoice data, cached totals, formulas, and wording; clears sample data", async () => {
    const invoice = (await parseWorkbook(await input([fixtureRow()])))
      .invoices[0];
    const bytes = await renderInvoice(await template(), invoice);
    const sheet = await read(bytes);
    expect(sheet.getCell("C7").value).toBe(invoice.customer);
    expect(sheet.getCell("F29").value).toBe(invoice.code);
    expect(sheet.getCell("B12").value).toBe("Sản phẩm thử");
    expect(sheet.getCell("B13").value).toBeNull();
    expect(sheet.getCell("H12").result).toBe(650000);
    expect(sheet.getCell("H22").result).toBe(650000);
    expect(sheet.getCell("A24").value).toBe("Sáu trăm năm mươi nghìn đồng.");
    expect(sheet.getCell("H12").formula).toContain("MOD(E12*F12,1)=0.5");
  });

  it("preserves styles, embedded images, QR, printer settings, and original notes byte-for-byte", async () => {
    const original = await JSZip.loadAsync(await template());
    const invoice = (await parseWorkbook(await createSample())).invoices[0];
    const output = await JSZip.loadAsync(
      await renderInvoice(await template(), invoice),
    );
    for (const path of [
      "xl/styles.xml",
      "xl/media/image1.png",
      "xl/media/image2.png",
      "xl/drawings/drawing1.xml",
      "xl/printerSettings/printerSettings1.bin",
      "xl/comments1.xml",
      "xl/sharedStrings.xml",
    ]) {
      expect(await output.file(path)!.async("uint8array")).toEqual(
        await original.file(path)!.async("uint8array"),
      );
    }
    expect(output.file("xl/calcChain.xml")).toBeNull();
    expect(
      await output.file("xl/_rels/workbook.xml.rels")!.async("string"),
    ).not.toContain("calcChain.xml");
  });

  it("expands beyond ten rows while moving footer, totals, merged cells, print area, and QR", async () => {
    const invoice = (
      await parseWorkbook(
        await input(Array.from({ length: 13 }, () => fixtureRow())),
      )
    ).invoices[0];
    const bytes = await renderInvoice(await template(), invoice);
    const sheet = await read(bytes);
    expect(sheet.getCell("B24").value).toBe("Sản phẩm thử");
    expect(sheet.getCell("A25").value).toBe("Thanh toán");
    expect(sheet.getCell("H25").result).toBe(8450000);
    expect(sheet.getCell("H25").formula).toBe("SUM(H12:H24)");
    expect(sheet.getCell("A27").value).toBe(amountInWords(8450000));
    expect(sheet.getCell("F32").value).toBe(invoice.code);
    expect(sheet.getCell("B24").isMerged).toBe(true);
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file("xl/workbook.xml")!.async("string")).toContain(
      "$H$40",
    );
    expect(
      await zip.file("xl/drawings/drawing1.xml")!.async("string"),
    ).toContain("<xdr:row>38</xdr:row>");
    const doc = await sheetDoc(bytes);
    const refs = Array.from(doc.getElementsByTagName("row")).map((row) =>
      row.getAttribute("r"),
    );
    expect(new Set(refs).size).toBe(refs.length);
    expect(refs.map(Number)).toEqual(
      [...refs.map(Number)].sort((a, b) => a - b),
    );
  });

  it("retains every invoice when sanitized filenames collide and neutralizes path characters", async () => {
    const invoices = (
      await parseWorkbook(
        await input([
          fixtureRow("../A"),
          fixtureRow("..\\A"),
          fixtureRow("CON"),
        ]),
      )
    ).invoices;
    const archive = await JSZip.loadAsync(
      await generateArchive(await template(), invoices),
    );
    expect(Object.keys(archive.files)).toEqual([
      ".._A.xlsx",
      ".._A (2).xlsx",
      "invoice_CON.xlsx",
    ]);
    expect(safeFilename("A/B:C?")).toBe("A_B_C_");
  });

  it("creates a useful sample and exports reference fixtures for independent Python verification", async () => {
    const buffer = await createSample();
    expect(buffer).toBeInstanceOf(ArrayBuffer);
    const parsed = await parseWorkbook(buffer);
    expect(parsed.invoices).toHaveLength(3);
    expect(parsed.rowCount).toBe(5);
    expect(
      parsed.invoices.reduce((sum, invoice) => sum + invoice.total, 0),
    ).toBe(2410000);
    const directory = ".qa";
    await mkdir(directory, { recursive: true });
    await writeFile(`${directory}/sample-orders.xlsx`, new Uint8Array(buffer));
    await writeFile(
      `${directory}/invoices.zip`,
      new Uint8Array(await generateArchive(await template(), parsed.invoices)),
    );
    const expanded = (
      await parseWorkbook(
        await input(Array.from({ length: 13 }, () => fixtureRow())),
      )
    ).invoices[0];
    await writeFile(
      `${directory}/expanded.xlsx`,
      await renderInvoice(await template(), expanded),
    );
    const rounding: Invoice = {
      ...parsed.invoices[0],
      products: [
        {
          name: "Half VND",
          unit: "Cái",
          quantity: 1,
          price: 2.5,
          adjustment: 0,
          total: 2,
        },
      ],
      total: 2,
    };
    await writeFile(
      `${directory}/rounding.xlsx`,
      await renderInvoice(await template(), rounding),
    );
  });
});
