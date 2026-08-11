# Bản đồ cấu trúc trang "Soạn thảo / Trình văn bản" — emoh.moh.gov.vn

> Đây là hiểu biết hiện tại của mình, tổng hợp từ 4 file HAR bạn gửi.
> Mục **?** hoặc **[CẦN XÁC NHẬN]** là chỗ mình chưa chắc — nhờ bạn soát.

---

## 0. Bối cảnh kỹ thuật

- **Nền tảng:** ứng dụng Java **Struts** (URL kiểu `Action!method.do`) + giao diện **dojo/vt widgets**.
- **Xác thực:** bằng **cookie phiên** (đăng nhập sẵn trên trình duyệt là dùng lại được).
- **Cấu trúc HAI TẦNG (đã xác nhận từ HAR đầy-cuối):** cái bạn tạo là một **Phiếu trình** (`voReport` / `reportForm`) — bên trong *chứa* **văn bản dự thảo** (`voPublishDocument`). Văn bản là một dòng trong lưới `draftDocumentGridForm[0]` của phiếu trình.
- **Đăng nhập:** POST username/mật khẩu tới `/passportv3/login` (SSO passport, không thấy captcha). Ta vẫn dùng phiên sẵn có, nhưng tự đăng nhập cũng khả thi.
- **Token: KHÔNG bắt buộc (đã xác nhận).** Form có `struts.token.name=token`, hệ thống có gọi `token!reloadToken` trước khi lưu/trình, NHƯNG cả `onInsertDraft` lẫn `onUpdate` thực tế **đều không gửi token nào** và vẫn thành công → script bỏ qua token.

---

## 1. Các "lớp"/mục của form soạn thảo

Trang soạn thảo gồm các khối (theo nhãn đọc được):

1. **Thông tin đơn vị soạn thảo** — người soạn, đơn vị, nhóm soạn thảo (cố định theo bạn).
2. **Thông tin văn bản** — loại VB, số/ký hiệu, trích yếu, cấp VB, độ mật, độ khẩn, ngày soạn/ngày ký, số trang, hạn trả lời...
3. **Người duyệt và ký trình** — người ký (signId), người duyệt (approveId), người ký trình/nháp (initialId).
4. **Nơi nhận** (nhiều loại — xem mục 4).
5. **File đính kèm** — Văn bản dự thảo, Ban hành kèm bản gốc (.docx/.xlsx), "Là bản ký".
6. **Phiếu trình** — khối liên quan luồng trình (bước Trình, ta KHÔNG đụng tới).
7. **Thu hồi / Thay thế** — lý do, thông tin VB cần thu hồi/thay thế (thường bỏ trống).
8. **Công bố** — "Trên Cổng thông tin điện ttử" (notPublishOnPortal).

---

## 2. Các trường ENUM (danh sách xổ) — ĐÃ CÓ ĐỦ ID

### Loại văn bản (`documentTypeId`) — 36 loại
```
66728 Báo cáo        66729 Biên bản       54465 Chỉ thị        54466 Chương trình
66701 Công điện      54468 Công hàm       54469 Công văn       66887 Đề án
54471 Đơn thư        66745 Giấy chứng nhận 66744 Giấy giới thiệu 54473 Giấy mời
66522 Hồ sơ          66768 Hợp đồng       66767 Hướng dẫn      54476 Kế hoạch
66741 Kết luận       54479 Nghị định      54480 Nghị quyết     54481 Pháp lệnh
65395 Phiếu báo      65394 Phiếu chuyển   54487 Phiếu trình    66743 Phúc đáp
66747 Quy trình      54482 Quyết định     66827 Sao y bản chính 65405 Thẩm định VBQPPL
54483 Thông báo      54484 Thông tư       54485 Thông tư liên tịch 66742 Tờ trình
65025 Trả lời, góp ý kiến  66746 Ủy quyền  65415 Uỷ thác tư pháp
```

### Độ khẩn (`priorityId`) — 5
```
54724 Hỏa tốc   54613 Thượng khẩn   66731 Bình thường   54764 Khẩn
```

### Độ mật (`securityTypeId`) — 4
```
66726 Bình thường   54459 Mật   54804 Tối mật   54708 Tuyệt mật
```

### Người ký (`signId`) — 9 (Lãnh đạo Bộ)
```
500008500 Đào Hồng Lan   500015082 Vũ Mạnh Hà   500001860 Đỗ Xuân Tuyên
500002780 Trần Văn Thuấn 500008080 Nguyễn Thị Liên Hương  500009840 Lê Đức Luận
500012803 Nguyễn Tri Thức  (+1 nữa)
```
> `initialId` (16 người) và `approveId` (4 người) cũng là danh sách người, đã bóc được — [CẦN XÁC NHẬN vai trò chính xác: ai là "ký trình", ai là "duyệt"].

---

## 3. Trường CỐ ĐỊNH theo hồ sơ của bạn (không đổi giữa các VB)

```
editorId              = 500013302   (Nguyễn Vũ Hùng)
editorName            = Nguyễn Vũ Hùng
publishOfficeId       = 273         (Phòng Quản lý chất lượng thuốc)
publishOfficeName     = Phòng Quản lý chất lượng thuốc
documentLevel         = 2           [CẦN XÁC NHẬN: có luôn = 2?]
publishDocType        = 0
typeOfDoc             = 0
isDigitSignDoc        = 1           (bật ký số)
notPublishOnPortal    = 1           [CẦN XÁC NHẬN: mặc định không công bố Cổng?]
```

---

## 4. NƠI NHẬN — cơ chế & 6 nhóm trường

Mỗi nhóm có **cặp** trường: chuỗi tên + chuỗi ID, nối bằng dấu `;`.
Bấm cây phòng ban trên UI chỉ để sinh ra mấy chuỗi ID này.

| Nhóm | Trường tên | Trường ID | Ví dụ đã thấy |
|------|-----------|-----------|----------------|
| Nhận nội bộ | `receiveInside` | `receiveInsideId` | Lãnh đạo Bộ = 206 |
| Liên thông văn bản | `receiveEdoc` | `receiveEdocId` | Cơ quan thuộc Chính phủ = 500001385 |
| Nơi lưu | `receiveSaveDepartment` | `receiveSaveDepartmentId` | Cục QL Dược;Bộ Y Tế = 210;52 |
| Để biết | `receiveToKnow` | `receiveToKnowId` | (chưa dùng) |
| Để báo cáo | `receiveReport` | `receiveReportId` | (chưa dùng) |
| Ngoài hệ thống | `receiveOutside` | `receiveOutsideId` | (chưa dùng) |

**Nguồn ID nơi nhận** (không cần dò từng cái, server trả theo lô):
- `departmentAction!getRootTree` / `getChildrenNode` → cây nội bộ Bộ Y Tế (đã thu **2451** đơn vị).
- `departmentAction!getTreeLinkDocument` → đầu mối liên thông: Sở Y tế các tỉnh, BHXH, các Bộ... (đã thu **96**).

---

## 5. FILE ĐÍNH KÈM

| Trường | Ý nghĩa |
|--------|---------|
| `attachDraftId` | ID file văn bản dự thảo (từ upload) |
| `signRequere` | ID file cần ký (thường = attachDraftId) |
| `attachReportId` | file phiếu trình (nếu có) |
| `attachSaveOriginalId` | bản gốc .docx/.xlsx ban hành kèm |

**Cách upload lấy ID:** POST file (multipart) tới `/vou/file/upload` → nhận HTML chứa
`postMessage(JSON.stringify([tên_file, "<ID>", "uploadDraftFile", token]))` → **lấy phần tử thứ 2 = ID file**.
Đúng lúc upload, hệ thống **tự gắn chữ ký chuyên viên (số 0)** ở server.

---

## 6. Trường ĐIỀU KHIỂN trạng thái

```
status              = 0     → LƯU NHÁP (điểm dừng an toàn; ta luôn để 0)
publishDocumentId   = trống → tạo mới  |  có giá trị → cập nhật dự thảo cũ
code                        → số/ký hiệu văn bản
documentAbstract            → trích yếu (là <textarea>)
createDatePublish   = hôm nay
dojo.preventCache           → mốc thời gian (chống cache)
```

---

## 7. Các ENDPOINT liên quan

Trình tự đầy đủ (từ HAR đầu-cuối), bỏ nhiễu:

| # | Endpoint | Vai trò | Script dùng? |
|---|----------|---------|--------------|
| 1 | `/passportv3/login` | đăng nhập | (dùng phiên sẵn) |
| 2 | `voReport!prepareInsert.do` | mở **Phiếu trình** mới | Có |
| 3 | `/vou/file/upload` | upload file phiếu trình → ID | Có |
| 4 | `voPublishDocument!prepareCreateDraft.do` | mở form văn bản (chứa mọi option) | Có — bóc bảng tra 1 lần |
| 5 | `/vou/file/upload` | upload file văn bản → ID | Có |
| 6 | `departmentAction!getChildrenNode/...` | cây nơi nhận | Có — thu ID nơi nhận |
| 7 | `voPublishDocument!onInsertDraft.do` | **LƯU văn bản dự thảo** → publishDocumentId | Có — bước chính |
| 8 | `voReport!searchNodeInFlow.do` | lấy danh sách luồng trình | (chỉ cần nếu tự chọn luồng) |
| 9 | `voReport!onUpdate.do` | **TRÌNH BAN HÀNH** (móc VB vào phiếu + chọn luồng `profileFlowAsignId`) | **KHÔNG** — bạn tự bấm |

**Phiếu trình (`reportForm`) — các trường chính ở bước Trình (mục 9):**
```
reportForm.reportId          = trống → tạo mới
reportForm.reportType        = 1
reportForm.officeId/creatorId= 210 / 500013302 (cố định)
reportForm.stateId           = 66731 (Bình thường)  — độ khẩn của phiếu
reportForm.content           = nội dung phiếu trình
reportForm.attachId          = ID file phiếu trình
reportForm.profileFlowAsignId= 500092130  (= luồng "TRÌNH BAN HÀNH")
draftDocumentGridForm[0].publishDocumentId = <id văn bản vừa lưu>
```

---

## 8. TÌNH TRẠNG CÁC ĐIỂM CHƯA CHẮC (đã giải quyết phần lớn)

1. ~~Quy trình con khi chọn nơi nhận~~ → **Rõ:** chỉ cần điền chuỗi ID, không cần lời gọi xác thực.
2. ~~Cấu trúc Phiếu trình~~ → **Đã có** (xem mục 7). Hai tầng: Phiếu trình chứa văn bản.
3. ~~Token~~ → **Rõ:** không bắt buộc.
4. **CÒN LẠI:** `initialId` (16 người) vs `approveId` (4 người) vs `signId` (9 người) — vai trò chính xác từng danh sách trong luồng của bạn? (signId = người ký ban hành; hai cái kia = duyệt/ký trình nội bộ — cần bạn xác nhận.)
5. **CÒN LẠI:** các trường có luôn cố định không (documentLevel, notPublishOnPortal, publishDocType, reportType=1, stateId phiếu)?
6. **CÒN LẠI:** có luôn đúng **2 file** (1 phiếu trình + 1 văn bản) không, hay số lượng thay đổi?

## 9. ĐIỂM DỪNG AN TOÀN CHO TỰ ĐỘNG HÓA

Script chạy tới **mục 7 (onInsertDraft — lưu văn bản dự thảo)** thì dừng, KHÔNG gọi mục 9 (onUpdate/Trình).
Bạn mở Phiếu trình trên trình duyệt, kiểm tra, chọn luồng và tự bấm **Trình**.
(Tùy chọn nâng cao sau này: script làm tới sát mục 9 nhưng để `profileFlowAsignId` trống = chỉ "Lưu" phiếu, chưa trình.)
