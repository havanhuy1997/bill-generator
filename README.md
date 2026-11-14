# Chi · Invoice Studio

A React + TypeScript port of the Django app at `D:\it\chi-2-excel-formater`, with a Vietnamese interface, ready for static hosting on GitHub Pages. Excel processing runs entirely in the browser. There is no Python service, database server, account, API key, or paid hosting dependency.

## Run locally

Install Node.js 24+ and pnpm 11, then run:

```sh
pnpm install
pnpm dev
```

Open the local URL printed by Vite. Choose **Dùng thử** for a working example, or download the sample workbook to see the expected columns.

```sh
pnpm test     # Workbook import and actual XLSX export tests
pnpm build    # TypeScript checks + production build to dist/
pnpm preview  # Serve the production build locally
```

## Deploy free on GitHub Pages

1. Put this project in a **public GitHub repository**, with `main` as the default branch. Use this new project folder; the Django project is not required.
2. In the repository, open **Settings → Pages → Build and deployment → Source → GitHub Actions**.
3. Push to `main`, or run **Deploy to GitHub Pages** from the Actions tab. The included workflow installs the locked dependencies, runs the tests, builds the app, and deploys `dist/`.

The deployment URL appears in the workflow and in Settings → Pages, normally `https://<username>.github.io/<repository>/`. `base: './'` supports both repository subpaths and custom domains without editing configuration. Navigation is within one page, so refreshing a view does not require server-side routing.

GitHub Free includes Pages for public repositories. Pages serves the HTML, JavaScript, and template; the visitor's own browser reads their workbook and builds the ZIP. See [GitHub Pages documentation](https://docs.github.com/en/pages/getting-started-with-github-pages/what-is-github-pages) and [Vite deployment guidance](https://vite.dev/guide/static-deploy.html#github-pages).

This workspace has not been connected to or published to a GitHub repository.

### Repo công khai hay riêng tư?

- **GitHub Pages với GitHub Free:** cần repo công khai (public). Repo riêng tư (private) có thể dùng Pages khi tài khoản có gói trả phí phù hợp. Xem [điều kiện GitHub Pages](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site).
- **Repo riêng tư và hosting miễn phí:** có thể kết nối repo GitHub private với Cloudflare Pages. Cloudflare Pages hỗ trợ [repo public và private](https://developers.cloudflare.com/pages/configuration/git-integration/) và có [gói Free](https://developers.cloudflare.com/pages/platform/limits/). Khi triển khai, chọn lệnh build `pnpm build`, thư mục đầu ra `dist` và Node.js 24.

Quyền riêng tư của repo kiểm soát ai xem được mã nguồn trên GitHub; website đã triển khai có thể vẫn công khai. Tệp Excel người dùng chọn không được gửi lên máy chủ hoặc đưa vào repo.

## Features and compatibility

| Original Django behavior | Browser version |
| --- | --- |
| Upload `.xlsx` | Choose or drag in one `.xlsx` workbook, up to 25 MB |
| Use the first worksheet | Same |
| Group by column H, “Số phiếu CKHO4” | Same, including non-adjacent rows |
| Use first customer and date in a group | Same; inconsistent rows produce a warning |
| Fill `sample_files/invoice.xlsx` | Uses an unchanged copy at `public/templates/invoice.xlsx` |
| One invoice workbook per code | Same; select all or a subset before generation |
| Vietnamese amount in words | Included, with corrected zero-group wording |
| Signed discount added to rounded line value | Same; negative discounts reduce the total |
| Download all invoices as ZIP | Same, with cached formula results already populated |
| Upload list, redownload, delete | Browser-local ZIP history in IndexedDB, with delete/undo |

The original Django database and uploaded files are not copied. History is **per browser and website origin**, survives normal reloads, and does not sync between devices. Clearing site data or browser storage eviction removes it. Keep a copy of important downloads. When storage is blocked or full, generation still works and the app explains that history cannot be saved. Only completed ZIPs and their metadata are saved; the source workbook is not retained.

## Input layout

Keep columns in their original positions. The sample includes Vietnamese headers. Column A is ignored.

| Column | Meaning | Notes |
| --- | --- | --- |
| B | Product name | Required for a row with an invoice code |
| C | Unit | Text, e.g. `Cái` |
| D | Quantity | Numeric, zero or greater |
| E | Unit price | Numeric, zero or greater |
| F | Discount/adjustment | Whole VND, signed; blank is zero |
| G | Customer details | First row for each invoice is used |
| H | Invoice code | Grouping key; rows with no code are skipped and reported |
| I | Date | Excel date or text; first row for each invoice is used |

Rows whose column H reads `Số phiếu CKHO4` are treated as headers. The app accepts Excel numeric cells or simple numeric strings such as `350000`; it rejects ambiguous strings such as `350,000` and currency labels. Formula cells need cached results: open, recalculate, and save the workbook in Excel before importing. The browser does not execute arbitrary workbook formulas.

## Intentional fixes

- **More than ten products:** insert styled product rows, extend totals and merged cells, and move the footer, QR drawing, and print area. The original code overwrote the footer.
- **Rounding:** retain Python's ties-to-even rounding for quantity × price. The generated Excel formulas use that same rule, fixing the old mismatch between Python `round` and Excel `ROUND` on `.5` ties.
- **Readable Vietnamese totals:** retain `lẽ`, `mốt`, and `lăm`, but omit unnecessary trailing zero groups. For example, `800000` becomes `Tám trăm nghìn đồng.` rather than `Tám trăm nghìn không trăm đồng.` Negative totals and zero are supported. Values must be safe whole-VND integers.
- **Filename safety:** replace invalid filename characters, handle Windows reserved names, and add suffixes when different invoice codes map to the same filename. No invoice is silently overwritten.
- **Validation:** report malformed numeric cells and Excel errors before exporting. Positive discounts and missing dates/customer details produce review notes.

The export changes the worksheet values and relevant workbook metadata directly inside the original XLSX archive. It preserves the original images, QR code, styles, shared strings, comments, notes, and printer settings. The in-app preview shows invoice data; the downloaded workbook uses the original branded layout.

## Structure

- `src/App.tsx`: responsive import, review, selection, previews, history, and guide.
- `src/lib/engine.ts`: input parsing, template expansion, XLSX output, and ZIP packaging.
- `src/lib/invoice.worker.ts`: keeps workbook work off the UI thread.
- `src/lib/numbers.ts`: original rounding rule and Vietnamese wording.
- `src/lib/storage.ts`: IndexedDB history and local downloads.
- `tests/engine.test.ts`: real-workbook regression tests, including formatting preservation and expanded invoices.
- `.github/workflows/deploy.yml`: tested-build deployment to GitHub Pages.

The Django source directory remains unchanged. No Django credentials, SQLite database, or customer uploads are included in this project.
