export interface Product {
  name: string;
  unit: string;
  quantity: number;
  price: number;
  adjustment: number;
  total: number;
}

export interface Invoice {
  code: string;
  customer: string;
  date: string;
  excelDate: number | string | null;
  products: Product[];
  total: number;
}

export interface ImportResult {
  invoices: Invoice[];
  rowCount: number;
  skippedRows: number;
  warnings: string[];
}

export interface Batch {
  id: string;
  name: string;
  createdAt: string;
  invoiceCount: number;
  productCount: number;
  total: number;
  zip: Blob;
}

export type WorkerRequest =
  | { type: "parse"; buffer: ArrayBuffer }
  | { type: "generate"; invoices: Invoice[]; template: ArrayBuffer }
  | { type: "sample" };

export type WorkerResponse =
  | { type: "parsed"; result: ImportResult }
  | { type: "progress"; done: number; total: number }
  | { type: "generated"; buffer: ArrayBuffer }
  | { type: "sample"; buffer: ArrayBuffer }
  | { type: "error"; message: string };
