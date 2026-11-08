# Quy tắc soạn file Word để hệ thống đọc đúng

Tài liệu này dựa trên cách `lib/question-parser.js`, `lib/answer-tables.js` và
`lib/docx-import.js` thực sự hoạt động. Làm đúng các quy tắc dưới đây thì hệ
thống nhận đúng câu hỏi, loại câu và đáp án; làm khác đi thì câu có thể bị đánh
dấu **Cần kiểm tra** chứ hệ thống **không tự đoán**.

---

## 0. Trước khi tải file lên

- Chỉ nhận file **`.docx`**, dung lượng tối đa **15 MB**.
- Vào **Review → Accept All Changes** để chấp nhận hết tracked changes; xoá
  comment. Nội dung còn ở dạng sửa đổi có thể bị đọc sai.
- **Gõ số câu bằng tay** (`Câu 1.`), không dùng chức năng đánh số tự động
  (AutoNumber/Bulleted List) của Word.
- Nên đặt nội dung trong **đoạn văn thường**. Tránh Text Box, WordArt, SmartArt
  và bảng trang trí cho phần câu hỏi.
- Công thức phải là **Equation của Word** (OMML). Ảnh chụp công thức chỉ được
  giữ làm ảnh, hệ thống **không OCR** để đọc chữ trong ảnh.
- Ký hiệu vector nên dùng Equation (`\vec{AB}`) hoặc gõ chữ kèm mũi tên tổ hợp
  (`AB⃗`). Hệ thống tự đổi `AB⃗` sang công thức để file xuất ra không bị mất dấu.
- Ảnh nhúng trong file được giữ nguyên, nhưng hệ thống **không đọc nội dung chữ
  trong ảnh**. Đừng đưa đề bài hoặc đáp án dưới dạng ảnh chụp.

---

## 1. Bố cục chung của đề

### 1.1. Tiêu đề phần

Mỗi phần đặt trên **một đoạn riêng**, bắt đầu bằng `Phần` (hoặc `PHẦN`) và số
La Mã/Ả Rập:

```text
PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN
PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI
PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN
```

Hệ thống nhận loại câu theo tiêu đề phần khi tiêu đề có các cụm:

| Loại câu | Cụm từ cần có trong tiêu đề |
|---|---|
| Trắc nghiệm 1 đáp án | `trắc nghiệm nhiều phương án` |
| Đúng/Sai | `trắc nghiệm đúng sai`, `PHẦN ... ĐÚNG SAI`, hoặc `chọn đúng hoặc sai` |
| Trả lời ngắn | `trả lời ngắn` |

Nếu tiêu đề không ghi loại, hệ thống suy ra: phần không có phương án nào → trả
lời ngắn; còn lại → trắc nghiệm 1 đáp án.

### 1.2. Câu hỏi

Câu hỏi phải bắt đầu bằng một trong hai dạng:

```text
Câu 1. Nội dung câu hỏi...
Câu 1: Nội dung câu hỏi...
1. Nội dung câu hỏi...
1: Nội dung câu hỏi...
```

**Không dùng** `Câu 1)` hoặc `Câu 1` không có dấu `.`/`:`. Không đặt câu hỏi
trong bảng.

Toàn bộ nội dung dẫn dắt của câu phải nằm **trước** các phương án. Đoạn văn nằm
sau các phương án sẽ không được hiểu là nội dung câu hỏi.

Bảng nội dung như **bảng biến thiên** hoặc **bảng số liệu** đặt giữa câu hỏi và
phần phương án được giữ nguyên cấu trúc khi xuất file. Không dùng bảng để chứa
toàn bộ câu hỏi/phương án.

### 1.3. Phương án

Mỗi phương án nên đặt trên **một đoạn riêng**:

```text
A. Phương án A
B. Phương án B
C. Phương án C
D. Phương án D
```

Chấp nhận `A.` hoặc `A)`. Có thể viết nhiều phương án trên cùng một dòng
(`A. ... B. ... C. ... D. ...`) nhưng **một dòng một phương án là an toàn nhất**.
Mỗi câu trắc nghiệm cần **ít nhất 2 phương án**.

Số thứ tự câu có thể đánh lại từ 1 ở mỗi phần. Trong cùng một phần nên đánh số
liên tục, không trùng, không nhảy cách.

### 1.4. Kết thúc đề

Dòng `HẾT` (hoặc `--- HẾT ---`) được bỏ qua, không thành câu hỏi.

---

## 2. Trắc nghiệm một đáp án

Đánh dấu **đúng một** phương án bằng một trong các cách:

- dấu `*` ở cuối phương án;
- chữ `(đúng)`;
- **gạch chân** chữ cái hoặc nội dung phương án.

Ví dụ:

```text
Câu 1. Thủ đô của Việt Nam là thành phố nào?
A. TP. Hồ Chí Minh
B. Hà Nội *
C. Đà Nẵng
D. Cần Thơ
```

Lưu ý:

- **In đậm không phải là dấu hiệu đáp án đúng.** Chỉ `*`, `(đúng)` và gạch chân
  được tính.
- Chỉ đánh dấu **một** phương án. Nếu đánh dấu nhiều phương án mà không khai báo
  câu nhiều đáp án, hệ thống chỉ giữ phương án đánh dấu cuối cùng.
- Có thể đánh dấu đáp án trong **bảng đáp án** ở cuối đề thay vì trong câu hỏi.

---

## 3. Trắc nghiệm nhiều đáp án

Phải khai báo rõ bằng thẻ `[MULTI]` ngay trong dòng câu hỏi:

```text
Câu 5. [MULTI] Chọn các số chẵn trong các số sau:
A. 1
B. 2 *
C. 4 *
D. 5
```

- Đánh dấu **tất cả** phương án đúng bằng `*`, `(đúng)` hoặc gạch chân.
- Không có `[MULTI]` thì câu được hiểu là trắc nghiệm một đáp án.
- Trong bảng đáp án, câu nhiều đáp án ghi nhiều chữ cái, ví dụ `A,C` hoặc
  `A; C`.

---

## 4. Câu trắc nghiệm Đúng/Sai

### 4.1. Cấu trúc

```text
PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI
Câu 1. Xét các phát biểu sau:
a) Nước có công thức hoá học là H2O.
b) Nước không dẫn điện.
c) Nước đóng băng ở 0 °C.
d) Nước có vị ngọt.
```

Các ý dùng `a)`, `b)`, `c)`, `d)` (chấp nhận cả `a.`). Chuẩn là **4 ý**.

### 4.2. Đánh dấu đáp án

- **Gạch chân** một ý → ý đó **Đúng**.
- Các ý **còn lại tự hiểu là Sai**.
- Có thể dùng `(Đ)`/`(đúng)` cho ý đúng và `(S)`/`(sai)` cho ý sai. Tuy nhiên
  phải có **ít nhất một dấu hiệu Đúng** (gạch chân hoặc `(Đ)`), nếu không hệ
  thống không dám kết luận và đánh dấu **Cần kiểm tra**.
- Nếu không đánh dấu ý nào, cả câu vào trạng thái **Cần kiểm tra**, không tự
  gán Sai.

Ví dụ: gạch chân ý a) và c) → kết quả `a = Đúng, b = Sai, c = Đúng, d = Sai`.

Có thể đánh dấu đáp án trong bảng đáp án Đúng/Sai ở cuối đề.

---

## 5. Câu trả lời ngắn

### 5.1. Cấu trúc

```text
PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN
Câu 1. Tính 2 + 3.
Câu 2. Nhiệt độ là bao nhiêu?
```

Câu trả lời ngắn **không có phương án A/B/C/D**.

### 5.2. Đáp án

Đáp án phải là **một số**:

- số nguyên: `5`, `-3`, `0`;
- số thập phân: `2.5` hoặc `2,5`;
- có thể có dấu `+`/`-`.

**Không dùng** phân số (`5/2`), đơn vị (`5 cm`), chữ (`năm`) hay biểu thức
(`2+3`) làm đáp án. Hãy ghi kết quả cuối cùng.

Ba cách ghi được hỗ trợ:

```text
Câu 1. Tính 2 + 3. Đáp án: 5

Câu 2. Tính 4 + 5.
Đáp án: 9

Câu 3. Tính 6 + 7.
13
```

Nhãn được nhận: `Đáp án`, `Đáp án đúng`, `ĐA`, `Kết quả`, `Answer`, theo sau là
dấu `:` hoặc `：`.

### 5.3. Khi thiếu đáp án

Câu trả lời ngắn không có đáp án vẫn được nạp để giáo viên thấy, nhưng bị đánh
dấu **Cần kiểm tra** và **chặn xuất file** cho tới khi nhập đáp án trên giao
diện.

---

## 6. Bảng đáp án ở cuối đề

### 6.1. Quy tắc chung

- Phải là **bảng Word thật** (Insert → Table), không phải ảnh chụp hay chữ căn
  bằng khoảng trắng.
- Dòng tiêu đề bảng đặt **ngay sát bảng**, không có đoạn văn hay dòng trống xen
  giữa:

```text
ĐÁP ÁN TRẮC NGHIỆM
<bảng ngay bên dưới>
```

- Tiêu đề phải bắt đầu bằng `ĐÁP ÁN`, `BẢNG ĐÁP ÁN` hoặc `PHẦN ... ĐÁP ÁN`.
  Nếu tiêu đề không đúng dạng này, hệ thống **không tự ghép** và sẽ cảnh báo.
- Trong tiêu đề nên ghi loại đáp án: `TRẮC NGHIỆM`, `ĐÚNG SAI`, `TRẢ LỜI NGẮN`.
- Nếu đề có **nhiều phần cùng loại**, tiêu đề bảng phải ghi rõ phần, ví dụ
  `ĐÁP ÁN PHẦN II`. Không ghi phần thì hệ thống từ chối ghép để tránh nhầm.
- Không đặt tiêu đề `ĐÁP ÁN` cho một bảng nội dung bình thường.

Các tên cột được nhận:

| Ý nghĩa | Tên cột hợp lệ |
|---|---|
| Số câu | `Câu`, `Câu hỏi`, `STT`, `Số thứ tự`, `Số câu` |
| Đáp án | `Đáp án`, `Answer` |
| Ý / mệnh đề | `Lệnh hỏi`, `Ý`, `Ý kiến`, `Mệnh đề`, `Phát biểu`, `Nội dung`, hoặc `a`, `b`, `c`, `d` |

Giá trị đáp án được nhận:

- trắc nghiệm: `A`, `B`, `C`, `D`; nhiều đáp án: `A,C`, `A; C`, `A/C`;
- đúng/sai: `Đ`, `S`, `Đúng`, `Sai`, `D`, `S`, `true`, `false`, `1`, `0`;
- trả lời ngắn: số nguyên/thập phân như mục 5.

Ô gộp **rowspan/colspan** được hỗ trợ. Bảng có hai cụm cột
`Câu | Đáp án | Câu | Đáp án` cũng được hỗ trợ.

### 6.2. Bảng có cột phụ `Câu 2`

```text
Câu | Lệnh hỏi | Đáp án (Đ/S) | Câu 2
1   | a         | Đ            | Đ
1   | b         | Đ            | Đ
```

Cột `Câu 2` chứa đáp án cho **câu kế tiếp** (số câu + 1). Quy tắc này chỉ đúng
khi câu phụ là câu liền sau; không dùng cho câu nhảy cách.

### 6.3. Bảng ma trận theo mã đề

```text
Câu hỏi | Mã đề thi
        | 111 | 112 | 113 | ...
1       | A   | B   | C   | ...
...
```

Quy tắc:

- Cột đầu là `Câu hỏi` (hoặc `Câu`); các cột sau là mã đề.
- Dòng mã đề có **ít nhất 2 mã**, mỗi mã là số 2–4 chữ số.
- Mỗi dòng dữ liệu là **một câu**, bắt đầu bằng số câu.
- Loại đáp án lấy theo **dạng giá trị**:
  - câu trắc nghiệm: **một chữ cái** `A`/`B`/`C`/`D`;
  - câu đúng/sai: **đúng 4 ký tự** `Đ`/`S` liền nhau, ví dụ `ĐĐSS`;
  - câu trả lời ngắn: **một số**.
- Với mẫu 22 câu thường gặp: câu 1–12 ghi chữ cái, câu 13–16 ghi `ĐĐSS`, câu
  17–22 ghi số.
- Hệ thống chỉ đọc **mã đề đầu tiên** và hiện cảnh báo rằng còn các mã khác.
- **Thứ tự câu trong file phải khớp đúng thứ tự dòng của bảng.** Bảng ma trận
  ghép theo vị trí câu trong đề, không theo số in ở từng phần. Thiếu/thừa một
  câu sẽ làm lệch toàn bộ các câu sau.
- Không ghi `PHẦN ...` vào tiêu đề bảng ma trận, vì bảng ma trận áp dụng cho
  toàn bộ đề.

---

## 7. Ảnh, công thức và khối mã

- **Công thức**: dùng Equation của Word. Hệ thống chuyển sang LaTeX khi xem
  trước và nhúng vào file HTML xuất ra.
- **Ảnh**: ảnh nhúng được giữ nguyên trong câu hỏi/phương án. Đặt ảnh ngay dưới
  nội dung liên quan. Ảnh nằm riêng sau các phương án có thể bị đánh dấu **Cần
  kiểm tra** để giáo viên xác nhận thuộc câu hỏi hay đáp án.
- **Khối mã**: đặt giữa `[CODE]` và `[/CODE]`, mỗi dòng một dòng riêng:

```text
[CODE]
print("hello")
[/CODE]
```

- **Thẻ loại câu** có thể đặt ngay trong dòng câu hỏi:
  `[MULTI]`, `[TRẢ LỜI NGẮN]`, `[SHORT]`, `[ĐÚNG/SAI]`.

---

## 8. Sau khi tải file lên

1. Xem số câu và số câu **Cần kiểm tra**.
2. Dùng bộ lọc **Tất cả / Có đáp án / Cần kiểm tra**.
3. Đọc cảnh báo bảng đáp án (nếu có).
4. Sửa trực tiếp câu hỏi, đáp án, điểm trên giao diện.
5. Chỉ xuất file HTML khi mọi câu đã có đáp án hợp lệ. Câu chưa có đáp án (thường
   nằm trong nhóm **Cần kiểm tra**) sẽ bị hệ thống chặn xuất và báo lỗi; hãy sửa
   hết trước khi tải.

---

## 9. Checklist nhanh

- [ ] File `.docx`, dưới 15 MB, đã Accept All Changes.
- [ ] Số câu gõ tay, không dùng auto-numbering.
- [ ] Tiêu đề phần đúng cụm từ và nằm trên đoạn riêng.
- [ ] Câu bắt đầu bằng `Câu N.` hoặc `N.`.
- [ ] Phương án `A.`/`B.`/`C.`/`D.`, mỗi phương án một đoạn.
- [ ] Đáp án đúng đánh dấu bằng `*`, `(đúng)` hoặc gạch chân (không dùng in đậm).
- [ ] Câu nhiều đáp án có `[MULTI]`.
- [ ] Đúng/Sai: gạch chân hoặc `(Đ)` ít nhất một ý đúng.
- [ ] Trả lời ngắn: đáp án là một số, không kèm đơn vị/phân số.
- [ ] Bảng đáp án là bảng Word thật, tiêu đề `ĐÁP ÁN ...` sát ngay trên bảng.
- [ ] Bảng đáp án ghi rõ `PHẦN ...` nếu đề có nhiều phần cùng loại.
- [ ] Bảng ma trận: đủ số dòng, đúng thứ tự câu, chỉ dùng mã đề đầu tiên.
- [ ] Công thức là Equation, không phải ảnh chụp.
