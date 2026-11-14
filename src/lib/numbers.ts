// Python's round(value, 0), used by the original Django application.
export function roundVnd(value: number): number {
  const floor = Math.floor(value);
  if (value - floor === 0.5) return floor % 2 === 0 ? floor : floor + 1;
  return Math.round(value);
}

const digits = [
  "không",
  "một",
  "hai",
  "ba",
  "bốn",
  "năm",
  "sáu",
  "bảy",
  "tám",
  "chín",
];

function readGroup(value: number, full: boolean): string {
  const hundreds = Math.floor(value / 100);
  const tens = Math.floor((value % 100) / 10);
  const ones = value % 10;
  const words: string[] = [];
  if (hundreds || full) words.push(digits[hundreds], "trăm");
  if (tens > 1) words.push(digits[tens], "mươi");
  else if (tens === 1) words.push("mười");
  else if (ones && (hundreds || full)) words.push("lẽ");
  if (ones)
    words.push(
      ones === 1 && tens > 1
        ? "mốt"
        : ones === 5 && tens > 0
          ? "lăm"
          : digits[ones],
    );
  return words.join(" ");
}

export function amountInWords(value: number): string {
  if (!Number.isSafeInteger(value))
    throw new Error(
      "Tổng tiền hóa đơn phải là số nguyên (VND) và nằm trong phạm vi hỗ trợ.",
    );
  if (value === 0) return "Không đồng.";
  let remaining = Math.abs(value);
  const groups: number[] = [];
  while (remaining > 0) {
    groups.push(remaining % 1000);
    remaining = Math.floor(remaining / 1000);
  }
  const units = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"];
  const result: string[] = [];
  for (let i = groups.length - 1; i >= 0; i--) {
    if (groups[i])
      result.push(
        readGroup(groups[i], i < groups.length - 1 && groups[i] < 100),
        units[i],
      );
  }
  const text = (value < 0 ? "âm " : "") + result.filter(Boolean).join(" ");
  return text.charAt(0).toUpperCase() + text.slice(1) + " đồng.";
}

export const money = (value: number) =>
  new Intl.NumberFormat("vi-VN", { maximumFractionDigits: 2 }).format(value);
