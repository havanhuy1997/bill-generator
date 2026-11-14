import { useEffect, useRef, useState } from "react";
import {
  ArrowDownToLine,
  ArrowLeft,
  ArrowRight,
  Check,
  CheckCheck,
  ChevronRight,
  CircleHelp,
  Clock3,
  FileCheck2,
  FileSpreadsheet,
  Files,
  FileText,
  FolderArchive,
  HardDrive,
  History,
  Info,
  LoaderCircle,
  LockKeyhole,
  Plus,
  Search,
  ShieldCheck,
  Sparkles,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import type { LucideIcon } from "lucide-react";
import { amountInWords, money } from "./lib/numbers";
import { deleteBatch, download, listBatches, saveBatch } from "./lib/storage";
import type {
  Batch,
  ImportResult,
  Invoice,
  WorkerRequest,
  WorkerResponse,
} from "./lib/types";

type Page = "generate" | "history" | "guide";
type Status = "idle" | "reading" | "generating";
const zipName = (name: string) => `${name.replace(/\.xlsx$/i, "")}-hoa-don.zip`;
const formatDate = (value: string) =>
  new Intl.DateTimeFormat("vi-VN", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(value));

function App() {
  const [page, setPage] = useState<Page>("generate");
  const [batches, setBatches] = useState<Batch[]>([]);
  const [result, setResult] = useState<ImportResult | null>(null);
  const [filename, setFilename] = useState("");
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [storageWarning, setStorageWarning] = useState("");
  const [dragging, setDragging] = useState(false);
  const [query, setQuery] = useState("");
  const [preview, setPreview] = useState<Invoice | null>(null);
  const [completed, setCompleted] = useState<Batch | null>(null);
  const [sampleLoading, setSampleLoading] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const activeWorker = useRef<Worker | null>(null);
  const [removed, setRemoved] = useState<Batch | null>(null);
  const busy = status !== "idle" || sampleLoading;

  useEffect(() => {
    let mounted = true;
    listBatches()
      .then((items) => {
        if (mounted) setBatches(items);
      })
      .catch(() => {
        if (mounted)
          setStorageWarning(
            "Không thể truy cập bộ nhớ trình duyệt. Bạn vẫn có thể tạo và tải hóa đơn, nhưng lịch sử chỉ được giữ trong phiên này.",
          );
      });
    return () => {
      mounted = false;
      activeWorker.current?.terminate();
    };
  }, []);

  useEffect(() => {
    if (!notice) return;
    const timer = setTimeout(() => setNotice(""), 6500);
    return () => clearTimeout(timer);
  }, [notice]);

  useEffect(() => {
    window.scrollTo({ top: 0, behavior: "instant" });
  }, [page]);

  async function runWorker(request: WorkerRequest): Promise<WorkerResponse> {
    return new Promise((resolve, reject) => {
      const worker = new Worker(
        new URL("./lib/invoice.worker.ts", import.meta.url),
        { type: "module" },
      );
      activeWorker.current = worker;
      const finish = () => {
        worker.terminate();
        activeWorker.current = null;
      };
      worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
        if (event.data.type === "progress") setProgress(event.data);
        else {
          finish();
          event.data.type === "error"
            ? reject(new Error(event.data.message))
            : resolve(event.data);
        }
      };
      worker.onerror = () => {
        finish();
        reject(
          new Error(
            "Không thể khởi động bộ xử lý Excel. Hãy tải lại trang và thử lại.",
          ),
        );
      };
      worker.postMessage(request);
    });
  }

  async function importFile(file: File) {
    if (busy) return;
    setError("");
    if (!/\.xlsx$/i.test(file.name)) {
      setError(
        "Vui lòng chọn tệp Excel .xlsx. Ứng dụng chưa hỗ trợ định dạng .xls và CSV.",
      );
      return;
    }
    if (!file.size || file.size > 25 * 1024 * 1024) {
      setError(
        "Vui lòng chọn tệp Excel có dữ liệu và dung lượng không quá 25 MB.",
      );
      return;
    }
    setStatus("reading");
    setCompleted(null);
    try {
      const response = await runWorker({
        type: "parse",
        buffer: await file.arrayBuffer(),
      });
      if (response.type === "parsed") {
        setResult(response.result);
        setFilename(file.name);
        setSelected(
          new Set(response.result.invoices.map((invoice) => invoice.code)),
        );
        setPage("generate");
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Không thể đọc tệp Excel.");
    } finally {
      setStatus("idle");
    }
  }

  async function sample(useNow: boolean) {
    if (busy) return;
    setSampleLoading(true);
    setError("");
    try {
      const response = await runWorker({ type: "sample" });
      if (response.type === "sample") {
        if (useNow) {
          // The sample uses the exact same import and validation path as your file.
          setSampleLoading(false);
          await importFile(new File([response.buffer], "don-hang-mau.xlsx"));
        } else
          download(
            new Blob([response.buffer], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }),
            "don-hang-mau.xlsx",
          );
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Không thể tạo tệp mẫu.");
    } finally {
      setSampleLoading(false);
    }
  }

  async function generate() {
    if (!result || !selected.size || busy) return;
    setStatus("generating");
    setProgress({ done: 0, total: selected.size });
    setError("");
    try {
      const templateResponse = await fetch(
        `${import.meta.env.BASE_URL}templates/invoice.xlsx`,
      );
      if (!templateResponse.ok)
        throw new Error(
          "Không thể tải mẫu hóa đơn. Hãy kiểm tra kết nối mạng và thử lại.",
        );
      const invoices = result.invoices.filter((invoice) =>
        selected.has(invoice.code),
      );
      const response = await runWorker({
        type: "generate",
        template: await templateResponse.arrayBuffer(),
        invoices,
      });
      if (response.type !== "generated") return;
      const batch: Batch = {
        id: crypto.randomUUID(),
        name: filename,
        createdAt: new Date().toISOString(),
        invoiceCount: invoices.length,
        productCount: invoices.reduce(
          (sum, invoice) => sum + invoice.products.length,
          0,
        ),
        total: invoices.reduce((sum, invoice) => sum + invoice.total, 0),
        zip: new Blob([response.buffer], { type: "application/zip" }),
      };
      try {
        await saveBatch(batch);
      } catch {
        setStorageWarning(
          "Không thể lưu lần tải này vào bộ nhớ trình duyệt. Hãy tải tệp ZIP ngay; tệp chỉ còn khả dụng trong phiên này.",
        );
      }
      setBatches((previous) => [batch, ...previous]);
      setCompleted(batch);
      download(batch.zip, zipName(batch.name));
      setNotice(
        `Đã tạo ${batch.invoiceCount} hóa đơn. Đang tải tệp ZIP xuống.`,
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Không thể tạo hóa đơn.");
    } finally {
      setStatus("idle");
    }
  }

  function reset() {
    setResult(null);
    setCompleted(null);
    setFilename("");
    setError("");
    setSelected(new Set());
    setPage("generate");
    window.scrollTo({ top: 0, behavior: "instant" });
  }

  async function remove(batch: Batch) {
    try {
      await deleteBatch(batch.id);
    } catch {
      setStorageWarning(
        "Không thể cập nhật bộ nhớ trình duyệt. Mục này có thể xuất hiện lại sau khi tải lại trang.",
      );
    }
    setBatches((items) => items.filter((item) => item.id !== batch.id));
    setRemoved(batch);
    setNotice("Đã xóa bản tải xuống khỏi trình duyệt này.");
  }

  async function undoRemove() {
    if (!removed) return;
    try {
      await saveBatch(removed);
    } catch {
      setStorageWarning(
        "Không thể lưu bản tải xuống đã khôi phục. Bạn vẫn có thể tải tệp trong phiên này.",
      );
    }
    setBatches((items) =>
      [...items, removed].sort((a, b) =>
        b.createdAt.localeCompare(a.createdAt),
      ),
    );
    setRemoved(null);
    setNotice("Đã khôi phục bản tải xuống.");
  }

  const totalInvoices = batches.reduce(
    (sum, batch) => sum + batch.invoiceCount,
    0,
  );
  const selectedInvoices =
    result?.invoices.filter((invoice) => selected.has(invoice.code)) ?? [];
  const selectedTotal = selectedInvoices.reduce(
    (sum, invoice) => sum + invoice.total,
    0,
  );
  const filteredBatches = batches.filter((batch) =>
    batch.name.toLowerCase().includes(query.toLowerCase()),
  );
  const title =
    page === "generate"
      ? "Tạo hóa đơn"
      : page === "history"
        ? "Lịch sử tải xuống"
        : "Hướng dẫn và mẫu";

  function HistoryTable({ compact = false }: { compact?: boolean }) {
    const items = compact ? batches.slice(0, 4) : filteredBatches;
    return items.length ? (
      <div className="table-scroll">
        <table className="history-table">
          <thead>
            <tr>
              <th>Tệp nguồn</th>
              <th>Ngày tạo</th>
              <th>Hóa đơn</th>
              <th>Trạng thái</th>
              <th>
                <span className="sr-only">Thao tác</span>
              </th>
            </tr>
          </thead>
          <tbody>
            {items.map((batch) => (
              <tr key={batch.id}>
                <td>
                  <div className="file-cell">
                    <span className="file-icon">
                      <FileSpreadsheet size={20} />
                    </span>
                    <div>
                      <strong>{batch.name}</strong>
                      <small>
                        {batch.productCount} dòng sản phẩm ·{" "}
                        {(batch.zip.size / 1024).toFixed(0)} KB ZIP
                      </small>
                    </div>
                  </div>
                </td>
                <td className="date-cell">{formatDate(batch.createdAt)}</td>
                <td>{batch.invoiceCount}</td>
                <td>
                  <span className="status-pill">
                    <span />
                    Sẵn sàng
                  </span>
                </td>
                <td>
                  <div className="row-actions">
                    <button
                      className="icon-button"
                      aria-label={`Tải xuống ${batch.name}`}
                      title="Tải tệp ZIP"
                      onClick={() => download(batch.zip, zipName(batch.name))}
                    >
                      <ArrowDownToLine size={18} />
                    </button>
                    <button
                      className="icon-button delete"
                      aria-label={`Xóa ${batch.name}`}
                      title="Xóa khỏi lịch sử"
                      onClick={() => void remove(batch)}
                    >
                      <Trash2 size={17} />
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    ) : (
      <div className="history-empty">
        <span className="empty-icon">
          <FolderArchive size={24} strokeWidth={1.5} />
        </span>
        <div>
          <strong>
            {query && !compact
              ? "Không tìm thấy tệp phù hợp"
              : "Lịch sử tải xuống của bạn"}
          </strong>
          <p>
            {query && !compact
              ? "Hãy thử tìm bằng tên tệp khác."
              : "Các tệp ZIP đã tạo sẽ xuất hiện ở đây để bạn tải lại khi cần."}
          </p>
        </div>
        {!compact && !query && (
          <button className="button secondary small" onClick={reset}>
            Tạo hóa đơn <ArrowRight size={15} />
          </button>
        )}
      </div>
    );
  }

  return (
    <div className="app-shell">
      <a className="skip-link" href="#main">
        Chuyển đến nội dung
      </a>
      <aside className="sidebar">
        <a
          className="brand"
          href="#"
          onClick={(event) => {
            event.preventDefault();
            setPage("generate");
          }}
          aria-label="Trang chủ Chi – Tạo hóa đơn"
        >
          <span className="brand-mark">
            <Files size={23} strokeWidth={1.8} />
          </span>
          <span className="brand-name">
            chi<span className="brand-dot">.</span>
            <small>TẠO HÓA ĐƠN</small>
          </span>
        </a>
        <div className="nav-label">KHÔNG GIAN LÀM VIỆC</div>
        <nav aria-label="Điều hướng chính">
          {(
            [
              { id: "generate", label: "Tạo hóa đơn", icon: FileText },
              { id: "history", label: "Lịch sử tải xuống", icon: History },
              { id: "guide", label: "Hướng dẫn và mẫu", icon: CircleHelp },
            ] as { id: Page; label: string; icon: LucideIcon }[]
          ).map((item) => (
            <button
              key={item.id}
              aria-label={item.label}
              title={item.label}
              className={`nav-item ${page === item.id ? "active" : ""}`}
              onClick={() => {
                setPage(item.id);
                setQuery("");
              }}
              aria-current={page === item.id ? "page" : undefined}
            >
              <item.icon size={19} strokeWidth={1.7} />
              <span>{item.label}</span>
              {item.id === "history" && batches.length > 0 && (
                <span className="nav-count">{batches.length}</span>
              )}
            </button>
          ))}
        </nav>
        <div className="sidebar-tip">
          <span className="tip-icon">
            <Sparkles size={19} />
          </span>
          <strong>Bớt thao tác. Thêm thời gian.</strong>
          <p>
            Chỉ một tệp Excel.
            <br />
            Đầy đủ mọi hóa đơn.
          </p>
          <button
            onClick={() => {
              setPage("guide");
              setQuery("");
            }}
          >
            Xem hướng dẫn <ArrowRight size={14} />
          </button>
        </div>
        <div className="sidebar-bottom">
          <span className="local-dot" />
          <span>Nhẹ nhàng hơn mỗi ngày</span>
          <small>Chi · v1.0</small>
        </div>
      </aside>

      <div className="main-shell">
        <header className="topbar">
          <div className="breadcrumb">
            <span>Làm việc</span>
            <ChevronRight size={14} />
            <strong>{title}</strong>
          </div>
          <div className="local-badge">
            <LockKeyhole size={13} />
            <span>Xử lý trên thiết bị</span>
            <span className="local-dot" />
          </div>
        </header>
        <main id="main">
          <div className="page-heading">
            <div>
              <p className="eyebrow">GIẢM VIỆC GIẤY TỜ MỖI NGÀY</p>
              <h1>
                {title}
                <span>.</span>
              </h1>
              <p className="heading-description">
                {page === "generate"
                  ? "Tạo hóa đơn từ đơn hàng Excel, nhanh chóng và gọn gàng."
                  : page === "history"
                    ? "Các hóa đơn đã tạo, sẵn sàng để tải lại bất cứ lúc nào."
                    : "Vẫn là tệp Excel quen thuộc, với cách làm đơn giản hơn."}
              </p>
            </div>
            <span className="date-label">
              <Clock3 size={15} />
              {new Intl.DateTimeFormat("vi-VN", {
                day: "numeric",
                month: "short",
                year: "numeric",
              }).format(new Date())}
            </span>
          </div>
          {storageWarning && (
            <div className="alert warning" role="status">
              <Info size={18} />
              <span>{storageWarning}</span>
            </div>
          )}
          {error && (
            <div className="alert error" role="alert">
              <Info size={19} />
              <span>{error}</span>
              <button
                className="icon-button"
                aria-label="Đóng thông báo lỗi"
                onClick={() => setError("")}
              >
                <X size={16} />
              </button>
            </div>
          )}

          {page === "generate" && (
            <>
              <div className="stats-row">
                <Stat
                  label="HÓA ĐƠN ĐÃ TẠO"
                  value={String(totalInvoices).padStart(2, "0")}
                  detail="Trong các lần tải xuống đã lưu"
                  icon={FileCheck2}
                />
                <Stat
                  label="TỆP EXCEL ĐÃ XỬ LÝ"
                  value={String(batches.length).padStart(2, "0")}
                  detail="Tiết kiệm thời gian mỗi lần tạo"
                  icon={FolderArchive}
                />
                <div className="privacy-stat">
                  <span className="stat-icon">
                    <ShieldCheck size={22} strokeWidth={1.5} />
                  </span>
                  <div>
                    <strong>Dữ liệu luôn ở trên máy bạn</strong>
                    <p>
                      Xử lý trực tiếp trên thiết bị.
                      <br />
                      Không gửi tệp lên máy chủ.
                    </p>
                  </div>
                  <span className="privacy-seal">
                    <Check size={13} />
                  </span>
                </div>
              </div>

              <section
                className="workflow-card"
                aria-label="Quy trình tạo hóa đơn"
                aria-busy={busy}
              >
                <div className="workflow-top">
                  <div className="workflow-title">
                    <span className="section-icon">
                      <Plus size={19} />
                    </span>
                    <h2>
                      {result ? "Kiểm tra dữ liệu" : "Tạo loạt hóa đơn mới"}
                    </h2>
                  </div>
                  <div className="steps" aria-label="Tiến trình">
                    <span className="step active">
                      <b>{result ? <Check size={12} /> : "1"}</b>Chọn tệp
                    </span>
                    <i />
                    <span className={`step ${result ? "active" : ""}`}>
                      <b>{completed ? <Check size={12} /> : "2"}</b>Kiểm tra
                    </span>
                    <i />
                    <span className={`step ${completed ? "active" : ""}`}>
                      <b>{completed ? <Check size={12} /> : "3"}</b>Tải xuống
                    </span>
                  </div>
                </div>
                {!result ? (
                  <div className="import-layout">
                    <div className="import-main">
                      <div className="section-overline">
                        <span>01 / TỆP EXCEL NGUỒN</span>
                        <span className="filetype">.XLSX</span>
                      </div>
                      <div
                        className={`dropzone ${dragging ? "dragging" : ""} ${status === "reading" ? "loading" : ""}`}
                        onDragOver={(event) => {
                          event.preventDefault();
                          if (!busy) setDragging(true);
                        }}
                        onDragLeave={() => setDragging(false)}
                        onDrop={(event) => {
                          event.preventDefault();
                          setDragging(false);
                          if (event.dataTransfer.files.length !== 1)
                            setError("Vui lòng chọn từng tệp Excel một.");
                          else void importFile(event.dataTransfer.files[0]);
                        }}
                      >
                        <div className="upload-art" aria-hidden="true">
                          <div className="paper-back" />
                          <div className="paper-front">
                            <FileSpreadsheet size={31} strokeWidth={1.4} />
                          </div>
                          <span>
                            <Upload size={15} strokeWidth={2} />
                          </span>
                        </div>
                        <h3>
                          {status === "reading"
                            ? "Đang đọc tệp Excel…"
                            : dragging
                              ? "Thả tệp vào đây"
                              : "Bắt đầu với một tệp Excel"}
                        </h3>
                        <p>Kéo thả tệp Excel vào đây, hoặc</p>
                        <button
                          className="button primary"
                          onClick={() => fileInput.current?.click()}
                          disabled={busy}
                        >
                          {status === "reading" ? (
                            <LoaderCircle size={17} className="spin" />
                          ) : (
                            <Plus size={17} />
                          )}
                          Chọn tệp Excel
                        </button>
                        <small>
                          Excel .xlsx · Tối đa 25 MB · Chỉ đọc trang tính đầu
                          tiên
                        </small>
                      </div>
                      <div className="sample-line">
                        <span>
                          <CircleHelp size={15} />
                          Lần đầu sử dụng?
                        </span>
                        <button
                          className="text-button"
                          onClick={() => void sample(false)}
                          disabled={busy}
                        >
                          Tải tệp mẫu <ArrowDownToLine size={14} />
                        </button>
                        <span className="sample-or">hoặc</span>
                        <button
                          className="text-button"
                          onClick={() => void sample(true)}
                          disabled={busy}
                        >
                          {sampleLoading ? "Đang chuẩn bị…" : "Dùng thử"}
                          <ArrowRight size={14} />
                        </button>
                      </div>
                    </div>
                    <aside className="how-it-works">
                      <div className="how-heading">
                        <span className="tiny-spark">
                          <Sparkles size={14} />
                        </span>
                        TẠO HÓA ĐƠN THẬT ĐƠN GIẢN
                      </div>
                      <h3>
                        Từ bảng tính
                        <br />
                        <em>đến hóa đơn.</em>
                      </h3>
                      <p className="how-intro">
                        Giữ tệp Excel quen thuộc.
                        <br />
                        Để Chi lo phần còn lại.
                      </p>
                      <ol>
                        <li>
                          <span>1</span>
                          <div>
                            <strong>Chọn tệp đơn hàng</strong>
                            <p>Mở tệp Excel bạn đang sử dụng.</p>
                          </div>
                        </li>
                        <li>
                          <span>2</span>
                          <div>
                            <strong>Kiểm tra nhanh dữ liệu</strong>
                            <p>Các hóa đơn được nhóm theo mã.</p>
                          </div>
                        </li>
                        <li>
                          <span>3</span>
                          <div>
                            <strong>Tải xuống là xong</strong>
                            <p>Nhận hóa đơn Excel trong một tệp ZIP.</p>
                          </div>
                        </li>
                      </ol>
                      <div className="template-note">
                        <FileText size={19} strokeWidth={1.4} />
                        <span>
                          Sử dụng mẫu hóa đơn gốc.
                          <br />
                          <strong>Giữ nguyên nội dung và bố cục.</strong>
                        </span>
                      </div>
                    </aside>
                  </div>
                ) : (
                  <div className="review-body">
                    <div className="source-summary">
                      <div className="file-cell">
                        <span className="file-icon">
                          <FileSpreadsheet size={24} />
                        </span>
                        <div>
                          <strong>{filename}</strong>
                          <small>
                            {result.rowCount} dòng sản phẩm ·{" "}
                            {result.invoices.length} hóa đơn được tìm thấy
                          </small>
                        </div>
                      </div>
                      <button
                        className="text-button"
                        onClick={reset}
                        disabled={busy}
                      >
                        <X size={15} />
                        Đổi tệp
                      </button>
                    </div>
                    {result.warnings.length > 0 && (
                      <details className="validation-notes">
                        <summary>
                          <Info size={15} />
                          {result.warnings.length} lưu ý cần kiểm tra
                        </summary>
                        <ul>
                          {result.warnings.map((warning) => (
                            <li key={warning}>{warning}</li>
                          ))}
                        </ul>
                      </details>
                    )}
                    {completed && (
                      <div className="success-banner" role="status">
                        <CheckCheck size={23} />
                        <div>
                          <strong>Hóa đơn đã sẵn sàng.</strong>
                          <p>
                            Đã đóng gói {completed.invoiceCount} hóa đơn Excel
                            trong một tệp ZIP.
                          </p>
                        </div>
                        <button
                          className="button secondary small"
                          onClick={() =>
                            download(completed.zip, zipName(completed.name))
                          }
                        >
                          <ArrowDownToLine size={16} />
                          Tải lại
                        </button>
                      </div>
                    )}
                    <div className="review-toolbar">
                      <strong>
                        Hóa đơn{" "}
                        <span className="count-tag">
                          {result.invoices.length}
                        </span>
                      </strong>
                      <span>Chọn các hóa đơn bạn muốn tạo</span>
                    </div>
                    <div className="table-scroll review-scroll">
                      <table className="invoice-table">
                        <thead>
                          <tr>
                            <th>
                              <input
                                type="checkbox"
                                aria-label="Chọn tất cả hóa đơn"
                                checked={
                                  selected.size === result.invoices.length
                                }
                                disabled={busy}
                                onChange={(event) =>
                                  setSelected(
                                    event.target.checked
                                      ? new Set(
                                          result.invoices.map(
                                            (invoice) => invoice.code,
                                          ),
                                        )
                                      : new Set(),
                                  )
                                }
                              />
                            </th>
                            <th>Mã hóa đơn</th>
                            <th>Khách hàng</th>
                            <th>Ngày</th>
                            <th>Số mặt hàng</th>
                            <th className="numeric">Thành tiền (VND)</th>
                            <th>
                              <span className="sr-only">Xem trước</span>
                            </th>
                          </tr>
                        </thead>
                        <tbody>
                          {result.invoices.map((invoice) => (
                            <tr
                              key={invoice.code}
                              className={
                                selected.has(invoice.code) ? "" : "unselected"
                              }
                            >
                              <td>
                                <input
                                  type="checkbox"
                                  aria-label={`Chọn ${invoice.code}`}
                                  checked={selected.has(invoice.code)}
                                  disabled={busy}
                                  onChange={(event) =>
                                    setSelected((current) => {
                                      const next = new Set(current);
                                      event.target.checked
                                        ? next.add(invoice.code)
                                        : next.delete(invoice.code);
                                      return next;
                                    })
                                  }
                                />
                              </td>
                              <td>
                                <button
                                  className="invoice-code"
                                  onClick={() => setPreview(invoice)}
                                >
                                  {invoice.code}
                                </button>
                              </td>
                              <td>{invoice.customer || "—"}</td>
                              <td className="date-cell">
                                {invoice.date || "—"}
                              </td>
                              <td>{invoice.products.length}</td>
                              <td className="numeric money">
                                {money(invoice.total)}
                              </td>
                              <td>
                                <button
                                  className="icon-button"
                                  onClick={() => setPreview(invoice)}
                                  aria-label={`Xem trước ${invoice.code}`}
                                >
                                  <ChevronRight size={17} />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="generate-footer">
                      <div>
                        <span>
                          Đã chọn {selected.size} / {result.invoices.length} hóa
                          đơn
                        </span>
                        <strong>
                          {money(selectedTotal)} <small>Tổng cộng (VND)</small>
                        </strong>
                      </div>
                      <button
                        className="button primary"
                        onClick={() => void generate()}
                        disabled={busy || !selected.size}
                      >
                        {status === "generating" ? (
                          <>
                            <LoaderCircle size={17} className="spin" />
                            Đang tạo {progress.done} / {progress.total}…
                          </>
                        ) : (
                          <>
                            <FolderArchive size={17} />
                            {completed
                              ? "Tạo lại tệp ZIP"
                              : "Tạo và tải tệp ZIP"}
                            <ArrowRight size={17} />
                          </>
                        )}
                      </button>
                    </div>
                    {status === "generating" && (
                      <div
                        className="progress-track"
                        role="progressbar"
                        aria-label="Đang tạo hóa đơn"
                        aria-valuenow={progress.done}
                        aria-valuemin={0}
                        aria-valuemax={progress.total}
                      >
                        <span
                          style={{
                            width: `${progress.total ? (progress.done / progress.total) * 100 : 0}%`,
                          }}
                        />
                      </div>
                    )}
                  </div>
                )}
              </section>

              <section className="history-section">
                <div className="section-heading">
                  <div>
                    <h2>Tải xuống gần đây</h2>
                    <p>Lưu gọn các hóa đơn đã tạo để dễ dàng tải lại.</p>
                  </div>
                  <button
                    className="text-button"
                    onClick={() => setPage("history")}
                  >
                    Xem lịch sử <ArrowRight size={15} />
                  </button>
                </div>
                <div className="history-card">
                  <HistoryTable compact />
                </div>
              </section>
            </>
          )}

          {page === "history" && (
            <section className="full-history">
              <div className="history-page-toolbar">
                <span>
                  <strong>{batches.length}</strong> lần tạo đã lưu ·{" "}
                  <strong>{totalInvoices}</strong> hóa đơn
                </span>
                <label className="search-field">
                  <Search size={16} />
                  <input
                    value={query}
                    placeholder="Tìm theo tên tệp…"
                    aria-label="Tìm trong lịch sử tải xuống"
                    onChange={(event) => setQuery(event.target.value)}
                  />
                </label>
                <button className="button primary small" onClick={reset}>
                  <Plus size={16} />
                  Tạo mới
                </button>
              </div>
              <div className="history-card">
                <HistoryTable />
              </div>
              <p className="storage-caption">
                <HardDrive size={14} />
                Lịch sử chỉ được lưu trong trình duyệt này và sẽ mất khi xóa dữ
                liệu trang web. Hãy giữ một bản sao các tệp ZIP đã tải xuống.
              </p>
            </section>
          )}

          {page === "guide" && (
            <div className="guide-layout">
              <section className="guide-card">
                <span className="guide-eyebrow">01 / TỆP DỮ LIỆU CỦA BẠN</span>
                <h2>Tệp Excel quen thuộc, chỉ cần đúng định dạng.</h2>
                <p>
                  Chi đọc trang tính đầu tiên trong tệp Excel. Bạn có thể bỏ
                  trống cột A; hãy giữ các cột B–I theo đúng thứ tự dưới đây.
                </p>
                <div className="table-scroll">
                  <table className="guide-table">
                    <thead>
                      <tr>
                        <th>Cột</th>
                        <th>Nội dung</th>
                        <th>Ví dụ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[
                        ["B", "Tên hàng hóa", "Loa Bluetooth mini"],
                        ["C", "Đơn vị tính", "Cái"],
                        ["D", "Số lượng", "2"],
                        ["E", "Đơn giá", "350000"],
                        ["F", "Giảm giá", "-50000"],
                        ["G", "Thông tin khách hàng", "Nguyễn Minh Anh"],
                        ["H", "Số phiếu CKHO4", "CKHO4/26/11/001"],
                        ["I", "Ngày", "14/11/2026"],
                      ].map(([column, field, example]) => (
                        <tr key={column}>
                          <td>
                            <span className="column-letter">{column}</span>
                          </td>
                          <td>{field}</td>
                          <td>{example}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="guide-actions">
                  <button
                    className="button primary small"
                    onClick={() => void sample(false)}
                    disabled={busy}
                  >
                    <ArrowDownToLine size={16} />
                    Tải tệp Excel mẫu
                  </button>
                  <button
                    className="button secondary small"
                    onClick={() => void sample(true)}
                    disabled={busy}
                  >
                    Dùng thử tệp mẫu
                    <ArrowRight size={16} />
                  </button>
                </div>
              </section>
              <aside className="guide-aside">
                <section className="guide-card compact-guide">
                  <span className="stat-icon">
                    <Info size={21} />
                  </span>
                  <h3>Một vài lưu ý</h3>
                  <ul>
                    <li>
                      Các dòng có cùng mã được gộp thành một hóa đơn. Thông tin
                      khách hàng và ngày lập được lấy từ dòng đầu tiên.
                    </li>
                    <li>
                      Giá trị giảm giá được cộng vào thành tiền. Hãy nhập số âm
                      để giảm tổng tiền; để trống được tính là 0.
                    </li>
                    <li>
                      Số lượng × đơn giá được làm tròn đến đơn vị đồng theo quy
                      tắc của ứng dụng gốc.
                    </li>
                    <li>
                      Các ô có công thức cần lưu sẵn kết quả. Hãy tính toán lại
                      và lưu tệp trong Excel trước khi sử dụng.
                    </li>
                    <li>
                      Nếu có hơn 10 mặt hàng, mẫu hóa đơn sẽ tự thêm dòng.
                    </li>
                  </ul>
                </section>
                <section className="template-download">
                  <FileText size={27} strokeWidth={1.4} />
                  <h3>Mẫu hóa đơn gốc</h3>
                  <p>
                    Giữ nguyên bố cục Excel, logo, mã QR và thiết lập in của mẫu
                    hóa đơn gốc.
                  </p>
                  <a
                    className="text-button"
                    href={`${import.meta.env.BASE_URL}templates/invoice.xlsx`}
                    download="mau-hoa-don.xlsx"
                  >
                    Tải mẫu hóa đơn
                    <ArrowDownToLine size={15} />
                  </a>
                </section>
              </aside>
              <section className="guide-card privacy-guide">
                <ShieldCheck size={29} strokeWidth={1.5} />
                <div>
                  <h3>Dữ liệu và lịch sử lưu trên thiết bị.</h3>
                  <p>
                    Tệp Excel được xử lý ngay trên thiết bị của bạn, không gửi
                    lên máy chủ. Tệp ZIP được lưu trong trình duyệt để bạn tải
                    lại khi cần. Lịch sử gắn với trình duyệt và địa chỉ trang
                    web này, không đồng bộ giữa các thiết bị.
                  </p>
                </div>
              </section>
            </div>
          )}

          <footer className="page-footer">
            <span>
              <ShieldCheck size={14} />
              Tạo hóa đơn nhẹ nhàng hơn mỗi ngày.
            </span>
            <span>
              Đồng hành cùng công việc mỗi ngày.{" "}
              <span className="footer-brand">chi.</span>
            </span>
          </footer>
        </main>
      </div>
      <input
        ref={fileInput}
        className="sr-only"
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        aria-label="Chọn tệp Excel nguồn"
        onChange={(event) => {
          if (event.target.files?.[0]) void importFile(event.target.files[0]);
          event.target.value = "";
        }}
      />
      {notice && (
        <div className="toast" role="status">
          <Check size={19} />
          <span>{notice}</span>
          {removed && notice.includes("removed") && (
            <button onClick={() => void undoRemove()}>Hoàn tác</button>
          )}
          <button
            className="icon-button"
            aria-label="Đóng thông báo"
            onClick={() => setNotice("")}
          >
            <X size={16} />
          </button>
        </div>
      )}
      {preview && (
        <InvoicePreview invoice={preview} close={() => setPreview(null)} />
      )}
    </div>
  );
}

function Stat({
  label,
  value,
  detail,
  icon: Icon,
}: {
  label: string;
  value: string;
  detail: string;
  icon: LucideIcon;
}) {
  return (
    <div className="stat">
      <div>
        <span className="stat-label">{label}</span>
        <strong>{value}</strong>
        <p>{detail}</p>
      </div>
      <span className="stat-icon">
        <Icon size={22} strokeWidth={1.5} />
      </span>
    </div>
  );
}

function InvoicePreview({
  invoice,
  close,
}: {
  invoice: Invoice;
  close: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  useEffect(() => {
    dialog.current?.showModal();
  }, []);
  return (
    <dialog
      className="preview-dialog"
      ref={dialog}
      onCancel={close}
      onClick={(event) => {
        if (event.target === event.currentTarget) close();
      }}
      aria-labelledby="preview-title"
    >
      <div className="preview-top">
        <button className="text-button" onClick={close}>
          <ArrowLeft size={16} />
          Quay lại kiểm tra
        </button>
        <button
          className="icon-button"
          aria-label="Đóng bản xem trước"
          onClick={close}
        >
          <X size={21} />
        </button>
      </div>
      <div className="preview-document">
        <div className="preview-label">
          <FileText size={17} />
          <span>XEM TRƯỚC DỮ LIỆU HÓA ĐƠN</span>
        </div>
        <h2 id="preview-title">Hóa đơn bán lẻ</h2>
        <p className="preview-subtitle">Kiêm phiếu bảo hành</p>
        <div className="preview-meta">
          <div>
            <small>KHÁCH HÀNG</small>
            <strong>{invoice.customer || "—"}</strong>
          </div>
          <div>
            <small>NGÀY</small>
            <strong>{invoice.date || "—"}</strong>
          </div>
        </div>
        <div className="preview-code">
          <span>Số phiếu</span>
          <strong>{invoice.code}</strong>
        </div>
        <div className="table-scroll">
          <table className="preview-table">
            <thead>
              <tr>
                <th>Sản phẩm</th>
                <th>SL</th>
                <th className="numeric">Đơn giá</th>
                <th className="numeric">Giảm giá</th>
                <th className="numeric">Thành tiền</th>
              </tr>
            </thead>
            <tbody>
              {invoice.products.map((product, i) => (
                <tr key={i}>
                  <td>
                    <strong>{product.name}</strong>
                    <small>{product.unit}</small>
                  </td>
                  <td>{product.quantity}</td>
                  <td className="numeric">{money(product.price)}</td>
                  <td className="numeric">{money(product.adjustment)}</td>
                  <td className="numeric">{money(product.total)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="preview-total">
          <span>Thanh toán</span>
          <strong>
            {money(invoice.total)} <small>VND</small>
          </strong>
        </div>
        <p className="amount-words">{amountInWords(invoice.total)}</p>
        <div className="signature-row">
          <span>
            Khách hàng<small>(Ký, họ tên)</small>
          </span>
          <span>
            Người bán hàng<small>(Ký, họ tên)</small>
          </span>
        </div>
      </div>
      <p className="preview-footnote">
        <Info size={14} />
        Đây là bản xem trước dữ liệu. Tệp Excel xuất ra sử dụng mẫu hóa đơn gốc.
      </p>
    </dialog>
  );
}

export default App;
