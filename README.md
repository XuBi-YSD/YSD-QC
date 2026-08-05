# ecoism-qc · App lập hồ sơ QC (chạy tĩnh, miễn phí trên GitHub Pages)

Ứng dụng web đơn giản để nhập dữ liệu và xuất ra đúng file Excel/Word theo
template chính thức `v1_20260729`. **Chạy hoàn toàn trong trình duyệt** —
không có server, không có backend, dữ liệu không rời khỏi máy người dùng
(chỉ tải template tĩnh về, không gửi gì lên đâu cả).

## Vì sao chọn cách này (không cần server)

- GitHub Pages host file tĩnh (HTML/CSS/JS) **miễn phí vĩnh viễn**, không giới hạn thời gian dùng thử.
- Việc đọc/ghi Excel (SheetJS) và Word (docxtemplater) đều làm được ngay trong trình duyệt.
- Không cần quản lý server, không cần lo bảo mật backend, không tốn chi phí vận hành.

## Cấu trúc thư mục

```
app/
├── index.html          # Trang chính
├── style.css
├── app.js               # Toàn bộ logic (đọc field_map, render form, xuất file)
├── data/
│   ├── field_map.json    # Danh sách ô cần nhập của từng sheet Excel (tự sinh từ diff Template vs bản đã điền)
│   ├── docx_forms.json   # Định nghĩa field cho 3 form Word
│   ├── locations.json    # Danh sách vị trí (lấy từ ảnh XuBi gửi)
│   ├── personnel.json    # Danh sách nhân sự hay lặp lại
│   ├── materials.json    # Danh sách vật liệu hay lặp lại
│   └── pipe_names.json   # Danh sách tên cống mẫu (S5, S12...)
├── templates_xlsx/       # 4 file template Excel gốc (đã xoá dữ liệu mẫu)
└── templates_docx/       # 3 file Word đã gắn tag {{...}} tại đúng vị trí cần điền
```

## Cách chạy thử ở máy (trước khi deploy)

```bash
cd app
python3 -m http.server 8000
# rồi mở http://localhost:8000 trên trình duyệt
```

(Không mở trực tiếp file `index.html` bằng `file://` — trình duyệt sẽ chặn
`fetch()` đọc file JSON/template. Phải chạy qua 1 server, dù chỉ là local.)

## Deploy lên GitHub Pages (miễn phí, ~5 phút)

1. Tạo repo mới trên GitHub (Public hoặc Private đều được, Private cần GitHub Pro/Team để bật Pages).
2. Đẩy toàn bộ nội dung thư mục `app/` lên nhánh `main` (hoặc nhánh bất kỳ):
   ```bash
   git init
   git add .
   git commit -m "Initial commit - ecoism-qc form app"
   git branch -M main
   git remote add origin https://github.com/<ten-to-chuc>/<ten-repo>.git
   git push -u origin main
   ```
3. Vào repo trên GitHub → **Settings → Pages**.
4. Ở mục "Build and deployment" → **Source**: chọn **Deploy from a branch**.
5. **Branch**: chọn `main`, thư mục `/ (root)` → **Save**.
6. Đợi 1-2 phút, GitHub sẽ cấp link dạng:
   `https://<ten-to-chuc>.github.io/<ten-repo>/`
7. Mở link đó — xong, dùng được ngay, chia sẻ link này cho cả nhóm QC.

## Cách cập nhật khi template đổi

Chỉ cần thay file trong `templates_xlsx/` hoặc `templates_docx/`, cập nhật lại
`data/field_map.json` (hoặc `docx_forms.json` nếu thêm/bớt field), rồi
`git push` — GitHub Pages tự cập nhật sau ~1 phút, không cần làm gì thêm.

### Cách tự sinh lại `field_map.json` khi có template mới (cần Python)

```bash
python3 scripts/build_field_map.py \
  --template path/to/Template.xlsx \
  --filled path/to/DaDien.xlsx \
  --label TenForm
```//(script mẫu, xem scripts/build_field_map.py)

## Quy tắc quan trọng đã áp dụng trong app

- **Ô không nhập → xuất ra để TRỐNG**, không dùng giá trị mẫu của template
  (đúng theo yêu cầu đã xác nhận với dev team, xem `Xac_nhan_Q1-Q9`).
- **Dropdown tự động gợi ý** cho các trường hay lặp: vị trí (Location), tên
  người tham gia, tên vật liệu, tên cống mẫu — người dùng vẫn có thể chọn
  "Khác (nhập tay)" nếu giá trị không có trong danh sách.
- **Tên file xuất ra** tự động gắn ngày dạng `YYYYMMDD-...` (lưu ý: mục 1.5(b)
  trước đó XuBi đã xác nhận KHÔNG áp dụng quy tắc đặt tên này cho phía APP —
  quy tắc này chỉ áp dụng cho file xuất từ chính app thử nghiệm này, có thể
  đổi lại dễ dàng trong `app.js`, hàm `todayStr()`/`exportXlsx()`/`exportDocx()`).

## Giới hạn hiện tại (cần biết trước khi dùng thật)

- Đây là bản MVP (minimum viable product) tập trung vào cơ chế điền dữ liệu +
  xuất file đúng template — CHƯA có: đăng nhập/phân quyền, lưu trữ tập trung
  (mỗi lần chỉ xuất 1 file về máy người dùng), lịch sử chỉnh sửa, đồng bộ
  nhiều người dùng cùng lúc.
- Form Excel dựa hoàn toàn vào `field_map.json` (tự sinh từ diff Template vs
  bản nháp) — nếu 1 ô nào đó KHÔNG thay đổi giữa Template và bản nháp nhưng
  thực ra vẫn cần nhập (VD: người điền để nguyên placeholder), ô đó sẽ KHÔNG
  xuất hiện trong form. Cần rà lại thủ công nếu phát hiện thiếu trường.
- 3 form Word dùng cơ chế tag `{{...}}` do Claude tự chèn dựa trên vị trí dấu
  chấm placeholder (`……`) trong bản Template — nên rà lại bản
  `templates_docx/*_Tagged.docx` để xác nhận đúng vị trí trước khi dùng thật
  (mở file bằng Word, tìm `{{` để xem toàn bộ vị trí đã đánh dấu).
- Thư viện SheetJS/docxtemplater tải từ CDN (cdnjs.cloudflare.com,
  unpkg.com) — cần máy có internet khi dùng (không hoạt động hoàn toàn
  offline trừ khi tự vendor thư viện vào repo).
