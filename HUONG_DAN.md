# Hướng dẫn chạy & test

## Cài nhanh trên máy Windows mới (khuyên dùng)

Máy đích **chưa cần có gì** (Python, Git… script tự lo). Chọn 1 trong 2:

**Cách A — dán 1 dòng:** mở **PowerShell**, dán rồi Enter:
```
iex (iwr -useb 'https://raw.githubusercontent.com/hungvumoh/trinh-van-ban-voffice/main/setup.ps1').Content
```

**Cách B — double-click:** tải file [`install.bat`](https://raw.githubusercontent.com/hungvumoh/trinh-van-ban-voffice/main/install.bat) về, nhấp đúp.

Script sẽ: cài Python (nếu thiếu) → tải chương trình về `%LOCALAPPDATA%\TroLyTrinhVanBan` →
cài thư viện (`requirements.txt`) → tạo lối tắt Desktop + Start Menu → mở chương trình.
**Chạy lại lệnh trên = cập nhật** lên bản mới nhất (giữ lại `noi_nhan.json` bạn sửa tay thành
`noi_nhan.json.bak`).

Cần: Windows 10+ 64-bit, có mạng. Không cần quyền admin.

---

## Cài thủ công (cho máy dev / khi script không chạy được)
1. Cài Python 3.9+ (tick "Add to PATH" khi cài trên Windows).
2. Mở Terminal/CMD, chạy: `pip install -r requirements.txt`
3. Để **các file dữ liệu cùng một thư mục** với `trinh_van_ban.py` (`du_lieu.json`, `noi_nhan.json`, `cay_don_vi.json`, `luong_trinh.json`).

## Chạy
```
python trinh_van_ban.py
```

## Đăng nhập
Nhập **Tên đăng nhập** + **Mật khẩu** (như khi vào web). Chương trình tự đăng nhập, lấy phiên. Không cần copy cookie nữa.
- Nếu hệ thống đòi **captcha**: chương trình mở ảnh captcha, bạn gõ mã vào hộp thoại hiện ra.
- Mật khẩu không lưu ra file; nhập lại mỗi lần mở chương trình.

## LẦN ĐẦU — chạy ở chế độ an toàn
1. Nhập tên đăng nhập + mật khẩu.
2. Bấm **Chọn…** ở "File phiếu trình", chọn file PDF thử nghiệm.
   (Để trống "File văn bản" là được — nó dùng chung file để thử.)
3. **Giữ nguyên dấu tick "Chỉ kiểm tra (không ghi gì)"**.
4. Bấm **CHẠY**.

→ Nó sẽ: mở form → upload file → xin token, rồi **DỪNG, không tạo gì cả**.
Nếu khung log hiện `✔ Kiểm tra OK` → phiên, upload, token đều chạy tốt.

**Nếu lỗi ở bước token hoặc upload:** copy nguyên khung log gửi mình. Đó là 2 mảnh mình chưa có mẫu (định dạng token / phản hồi upload) — chỉ cần chỉnh 1–2 dòng.

## LẦN THẬT — lưu nháp
1. **Bỏ tick "Chỉ kiểm tra"**.
2. Nhóm **VĂN BẢN**: mỗi văn bản dự thảo là **1 khối riêng** (file + Loại VB + Số/ký hiệu +
   Trích yếu). Bấm **"+ Thêm văn bản"** nếu 1 phiếu trình cần gửi nhiều văn bản cùng lúc — mỗi
   khối có thể khác loại/khác số hoàn toàn (chương trình lưu từng văn bản riêng rồi mới gộp vào
   chung 1 phiếu trình, đúng như cách trang web tự làm khi bạn bấm "Thêm văn bản" nhiều lần).
   Bấm dấu **✕** ở 1 khối để bỏ bớt (luôn phải còn ít nhất 1 khối).
3. Trong mỗi khối, chọn **file Dự thảo văn bản** (PDF, hoặc **.docx** — chương trình tự
   chuyển sang PDF trước bằng Word cài sẵn trên máy, mất vài giây, xong tự đọc tiếp như PDF.
   File **.doc** (định dạng cũ) chưa hỗ trợ — tự "Save As" sang .docx hoặc .pdf trong Word trước)
   thường) trước — chương trình tự đọc trang 1 và
   điền sẵn **Loại VB**, **Số/ký hiệu**, **Trích yếu** CHO RIÊNG khối đó, và tự chọn **Luồng
   trình** (dùng chung cho cả phiếu trình) theo ký hiệu (chứa "QLD" → Luồng Cục, "CL" → Luồng
   Phòng, "BYT" → Luồng Bộ).
   - Mẫu **Công văn** (có dòng "Số: .../..." và "V/v ...", không có dòng tên loại): điền như trước.
   - Mẫu **có tên loại** viết HOA ngay dưới dòng "Số:" (Quyết định, Kế hoạch, Giấy mời,
     Giấy chứng nhận, và mọi loại khác trong danh sách "Loại văn bản"): chương trình nhận
     ra tên loại đó và tự chọn đúng combobox **Loại VB**, rồi lấy trích yếu từ đoạn ngay dưới
     dòng tên loại (dừng khi gặp "Căn cứ/Thực hiện/Nhằm..." hoặc dòng chức danh viết HOA khác).
     Riêng **Giấy chứng nhận** (không có đoạn trích yếu rõ ràng) sẽ được điền sẵn đúng cụm từ
     "Giấy chứng nhận đủ điều kiện kinh doanh dược" và **luôn** chọn **Luồng Bộ** (loại này luôn
     do Bộ Y tế cấp, không xét ký hiệu).
   - Nếu không nhận ra được gì cả (mẫu lạ, PDF quét ảnh không có lớp chữ...): báo trong khung
     log, tự điền tay là được.
   Luôn kiểm tra lại các trường trước khi bấm CHẠY — tự điền chỉ để đỡ gõ tay, không phải lúc nào cũng đúng 100%.
4. "Nội dung phiếu" (dùng chung cho cả phiếu trình) tự điền theo văn bản ĐẦU TIÊN nếu còn trống —
   nếu có nhiều văn bản, kiểm tra/sửa lại cho khớp nội dung chung.
5. Nơi nhận: gõ **đúng tên** đơn vị, cách nhau bằng dấu phẩy (vd: `Lãnh đạo Bộ, Cục Quản lý Dược`).
   Tên phải khớp với `noi_nhan.json`; sai tên nó sẽ báo ngay tên nào không thấy.
6. Bấm **CHẠY** — vì "Chỉ kiểm tra" đang tắt, chương trình mở khung **Xem trước & Gửi** thay vì gửi ngay.

## Khung XEM TRƯỚC & GỬI
- Bên trái: thông tin đã điền (nội dung, luồng trình, nơi nhận, và tóm tắt từng văn bản) + danh
  sách file, chia nhóm PHIẾU TRÌNH và **1 nhóm VĂN BẢN cho mỗi văn bản đã thêm** (VĂN BẢN 1, VĂN
  BẢN 2, ...). Kéo thanh ngăn giữa phần "Thông tin đã lưu" và phần "File" nếu 1 bên cần nhiều chỗ hơn.
- Bấm vào 1 file PDF (Phiếu trình hoặc file chính của BẤT KỲ văn bản nào) ở bên trái → bên phải
  hiện trang PDF đó. Nếu tick "Tự đánh số chữ ký" đang bật, chương trình tự quét toàn bộ file (mọi
  trang) CỦA VĂN BẢN ĐÓ, tìm các dòng CHUYÊN VIÊN/TRƯỞNG PHÒNG/CỤC TRƯỞNG/THỨ TRƯỞNG viết HOA, và
  hiện thành các **dấu tròn số** ngay tại vị trí định đóng dấu theo Luồng trình đã chọn — **kéo**
  dấu để di chuyển, **nhấp đúp** để sửa số/chức danh hoặc xoá, bấm **"➕ Thêm dấu"** rồi nhấp vào
  trang để thêm dấu mới (chọn chức danh, số tự điền theo luồng, vẫn sửa được tay). Mỗi văn bản có
  dấu riêng, sửa dấu ở văn bản này không ảnh hưởng văn bản khác.
- Bấm vào file PDF phụ (tài liệu thêm) → chỉ xem, không có dấu để sửa.
- Bấm vào file không phải PDF (.doc, .docx...) → mở thư mục chứa file đó trong Explorer, không đổi khung xem.
- Kiểm tra xong, có 2 lựa chọn ở trên cùng:
  - **"Lưu dự thảo"** → chương trình gửi thật (upload + lưu nháp CHO TỪNG VĂN BẢN, rồi gộp tất
    cả vào 1 phiếu trình, `sign=0`) nhưng CHƯA đi vào luồng ký duyệt — bạn tự mở web, thùng nháp
    phiếu trình, kiểm tra rồi tự bấm Trình sau.
  - **"Trình văn bản"** → làm y hệt "Lưu dự thảo" nhưng gửi kèm `sign=1` — phiếu trình đi THẲNG
    vào luồng ký duyệt ngay, không cần vào web bấm Trình nữa.
  - Cả 2 đều dùng đúng vị trí các dấu đang hiện trên màn hình (kể cả bạn vừa sửa tay). Đóng cửa
    sổ này mà không bấm nút nào thì không có gì được gửi/ghi.

→ Log (trong khung Xem trước) báo `✔ XONG. Phiếu trình đã lưu NHÁP.` hoặc `...đã được TRÌNH.`
tuỳ nút đã bấm.

## Lưu ý
- "Trình văn bản" gửi thẳng vào luồng ký duyệt, không hỏi xác nhận lại — nếu lỡ trình nhầm vẫn
  thu hồi được trên hệ thống như bình thường.
- Mỗi lần mở lại khung Xem trước, các dấu được quét lại từ đầu — sửa tay lần trước **không được nhớ** lại.
- Danh sách "Luồng trình" giờ lấy TỰ ĐỘNG từ web theo đúng tài khoản đang đăng nhập — không cần
  bắt file `.har` tìm `flowId` nữa. Combobox tự hiện tối đa 5 luồng "quen" ở đầu (theo bạn đã
  ghim/hay dùng nhất trên máy này), phần còn lại nằm dưới 1 dòng phân cách. Chọn 1 luồng chưa có
  sẵn người ký (vd luồng dùng chung của cả đơn vị) sẽ tự hiện thêm khung "Chọn người ký" ngay dưới.
  Dữ liệu ghim/tần suất + quy tắc tự nhận luồng theo ký hiệu văn bản lưu trong `luong_trinh.json`
  (sổ RIÊNG của máy đang chạy — máy khác cài chương trình này sẽ có sổ trống, tự học lại từ đầu).
- Muốn thêm nơi nhận chưa có trong sổ: mở `noi_nhan.json`, thêm dòng `"Tên đơn vị": "ID"`.

## Chạy ngầm khi bật máy (chỉ Windows)

Mục đích: bật máy lên là chương trình tự đăng nhập sẵn, nằm im dưới **khay hệ thống** (góc phải
thanh taskbar, cạnh đồng hồ — Windows tự dồn vào nút mũi tên `^`). Bấm icon là mở ra ngay,
không phải chờ đăng nhập lại.

**Cài lần đầu:**
1. Mở chương trình bình thường, đăng nhập, **tick "Nhớ đăng nhập"** (mật khẩu lưu trong
   Windows Credential Manager, không ra file).
2. Trên thanh trên cùng, tick **"Khởi động cùng Windows"**.

Từ lần bật máy sau: chương trình tự chạy ẩn (`--tray`), tự đăng nhập, hiện balloon "Đã sẵn sàng".

**Dùng:**
- **Bấm (đúp) icon khay** hoặc chuột phải → *Mở*: hiện cửa sổ chính.
- Bấm **✕** trên cửa sổ: chỉ thu xuống khay, chương trình **vẫn chạy**. (Đúng trên Windows
  bất cứ khi nào có icon khay — kể cả mở chương trình bình thường, không riêng chế độ `--tray`.
  Thiếu `pystray`/`pillow` → không có icon khay và ✕ = thoát hẳn.)
- Chuột phải icon khay: *Mở* · *Đăng nhập lại* · *Khởi động cùng Windows* (bật/tắt) · **Thoát**
  (thoát hẳn).
- Mở lần thứ 2 (tự bấm shortcut trong khi đã có bản chạy): không mở thêm bản mới, chỉ bật cửa
  sổ bản đang chạy lên.

**Giữ phiên:** chạy nhiều ngày liền thì phiên VOffice sẽ hết hạn — chương trình cứ ~25 phút tự
kiểm tra, rớt thì tự đăng nhập lại bằng mật khẩu đã lưu.

**Tắt hẳn tính năng:** bỏ tick "Khởi động cùng Windows" (hoặc menu khay → bỏ tick).

**Cần thư viện:** `pip install pystray pillow` (bản đóng gói .exe đã kèm sẵn). Thiếu 2 thư viện
này thì `--tray` tự chuyển về mở cửa sổ bình thường.

> macOS: chưa có. Cờ `--tray` trên mac chỉ mở cửa sổ như thường.

## Khi gặp lỗi / muốn góp ý cải tiến
Chương trình tự ghi log ra file **`app_log.txt`** (cùng thư mục với `settings.json`/`nguoi_dung.json`) —
kể cả lỗi không lường trước được (crash) cũng được ghi lại đầy đủ vào đây. File này CHỈ nằm trên máy
bạn, chương trình không tự gửi đi đâu cả. Khi gặp lỗi/muốn báo để cải tiến, gửi lại file này (hoặc
đoạn cuối gần thời điểm xảy ra lỗi) — có sẵn dòng "=== Khởi động chương trình — bản ... ===" ở đầu mỗi
lần mở app để biết chính xác đang chạy bản nào. Giữ tối đa 14 ngày gần nhất, tự xoá bớt file cũ hơn.
