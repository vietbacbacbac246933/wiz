# Word → HTML Quiz

Công cụ nội bộ tối giản: tải file Word (`.docx`) lên, xem lại câu hỏi, tải xuống **một gói ZIP** chứa file HTML chạy được khi không có mạng.

Gói ZIP gồm: `<slug>.html`, `manifest.webmanifest`, `sw.js`, `icon-192.png`, `icon-512.png` — thả thẳng vào thư mục gốc của repo thư viện để importer xử lý.

Không có tài khoản, không có cơ sở dữ liệu, không lưu file gốc hay ảnh trên đĩa. File Word và ảnh nhúng chỉ nằm trong RAM và tự hết hạn sau 30 phút.

## Cài đặt

```bash
npm install
npm start
```

Mở http://localhost:3456 (đổi cổng bằng biến môi trường `PORT`).

## Ba bước sử dụng

1. **Tải file Word** — kéo thả hoặc chọn một file `.docx` (tối đa **15 MB**).
2. **Xem lại câu hỏi** — kiểm tra câu hỏi, đáp án, điểm; sửa trực tiếp nếu cần. Câu chưa xác định được đáp án sẽ được đánh dấu để giáo viên tự kiểm tra.
3. **Tải file ZIP** — đặt tiêu đề, chọn chế độ làm bài rồi tải xuống. File HTML bên trong mở được trên máy không có mạng: mọi CSS, JavaScript, công thức và ảnh đã nhúng sẵn; manifest và service worker đi kèm để dùng ngoại tuyến.

## Định dạng Word được hỗ trợ

- Câu hỏi bắt đầu bằng `Câu 1.`, `Câu 2:` (hỗ trợ cả `Cau 1.` không dấu).
- Phương án `A.`, `B.`, `C.`, `D.`; đáp án đúng đánh dấu bằng `*`, gạch chân hoặc in đậm.
- Câu đúng/sai với các ý `a)`, `b)`, `c)`, `d)` và đáp án `Đ`/`S`.
- Câu trả lời ngắn: đáp án là một số hoặc chuỗi ngắn.
- Bảng đáp án đặt ở cuối đề (`ĐÁP ÁN PHẦN I`, ...) theo hàng ngang hoặc hàng dọc.
- Công thức toán (OMML) được chuyển thành LaTeX; ảnh nhúng trong file Word được giữ lại.
- Nếu thiếu hoặc mâu thuẫn đáp án, hệ thống **không tự đoán**: câu đó được đánh dấu "cần kiểm tra" và chặn xuất file cho tới khi giáo viên sửa.

## Kiểm thử

```bash
npm test
```

Các file mẫu nằm trong `data/`. Test tự khởi động server ở cổng 3999.
