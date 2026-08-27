# -*- coding: utf-8 -*-
"""
Test TẦNG 1: các hàm THUẦN (đưa gì vào ra đúng cái đó, không đụng mạng/server/GUI).
Chạy:  python3 -m unittest test_trinh_van_ban -v
Test nào tham chiếu file cá nhân của tác giả (PDF mẫu, HAR mẫu) mà máy đang chạy không có sẽ
tự bỏ qua (skip), không báo lỗi giả.
"""
import json
import os
import sys
import tempfile
import unittest
import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("trinh_van_ban", os.path.join(HERE, "trinh_van_ban.py"))
tvb = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tvb)

# ---------- File mẫu cá nhân (bỏ qua nếu máy đang chạy không có) ----------
PDF_GIAY_MOI = "/Users/hnguyen/Downloads/Làm việc ngày 6.7.2026/Trình họp Thông tư/Giấy mời hop - 08.07.26.pdf"
PDF_PT_THU_TRUONG = "/Users/hnguyen/Downloads/Làm việc tối 30.6.2026/Gửi in/2. Phiếu trình Thứ trưởng.pdf"
PDF_PT_CUC_CU = "/Users/hnguyen/Library/CloudStorage/Dropbox/A Politician/2. LƯU/Cũ/2.1. PHIẾU TRÌNH Lãnh đạo Cục.pdf"
PDF_PT_CUC_AC = "/Users/hnguyen/Library/CloudStorage/Dropbox/A Politician/1. SỰ VỤ/25.01.02. Hủy thuốc HIV AIDS/2.1. PHIẾU TRÌNH Lãnh đạo Cục - AC.pdf"
PDF_PT_CUC_ZPV = "/Users/hnguyen/Library/CloudStorage/Dropbox/A Politician/1. SỰ VỤ/25.5.30. ZPV HPC/2.1. PHIẾU TRÌNH Lãnh đạo Cục.pdf"
HAR_PREPARE_INSERT = "/Users/hnguyen/Downloads/000. Xem xoá/har 3 mở phiếu trình mới và chọn các luồng trình.har"
PDF_WATERMARKED = ("/Users/hnguyen/Downloads/000. Xem xoá/Van ban di (1)/"
                    "So_2.4._Phu_luc_X_-_Bieu_mau_-_17.06.26.pdf")

def _skip_if_missing(path):
    return unittest.skipUnless(os.path.exists(path), f"thiếu file mẫu: {path}")


# ==================== extract_draft_fields (đọc PDF dự thảo) ====================
class TestExtractDraftFields(unittest.TestCase):
    @_skip_if_missing(PDF_GIAY_MOI)
    def test_giay_moi(self):
        doc_type, code, abstract = tvb.extract_draft_fields(PDF_GIAY_MOI)
        self.assertEqual(doc_type, "Giấy mời")
        self.assertIn("GM-QLD", code or "")


# ==================== extract_phieu_trinh_content (bold-boundary) ====================
class TestExtractPhieuTrinhContent(unittest.TestCase):
    @_skip_if_missing(PDF_PT_THU_TRUONG)
    def test_thu_truong_noi_dung_xin_y_kien(self):
        content = tvb.extract_phieu_trinh_content(PDF_PT_THU_TRUONG)
        self.assertIsNotNone(content)
        self.assertTrue(content.startswith("Về việc báo cáo, xin ý kiến"))

    @_skip_if_missing(PDF_PT_CUC_CU)
    def test_cuc_ten_van_ban_trinh_uu_tien_truoc(self):
        # File này có CẢ "Tên văn bản trình:" lẫn "Nội dung xin ý kiến:" — phải lấy cái xuất
        # hiện TRƯỚC ("Tên văn bản trình"), không phải cái sau.
        content = tvb.extract_phieu_trinh_content(PDF_PT_CUC_CU)
        self.assertIsNotNone(content)
        self.assertTrue(content.startswith("Thông báo ý kiến kết luận của Cục trưởng Vũ Tuấn Cường"))

    @_skip_if_missing(PDF_PT_CUC_AC)
    def test_cuc_ac_noi_dung_trinh_ngan(self):
        content = tvb.extract_phieu_trinh_content(PDF_PT_CUC_AC)
        self.assertEqual(content, "Dự thảo văn bản đính kèm")

    @_skip_if_missing(PDF_PT_CUC_ZPV)
    def test_cuc_zpv_ten_van_ban_trinh(self):
        content = tvb.extract_phieu_trinh_content(PDF_PT_CUC_ZPV)
        self.assertIsNotNone(content)
        self.assertTrue(content.startswith("Kiểm tra, xác minh Lô sản phẩm vắc xin"))


# ==================== strip_view_watermark (xoá watermark "đã xem/tải") ====================
class TestStripViewWatermark(unittest.TestCase):
    @_skip_if_missing(PDF_WATERMARKED)
    def test_xoa_sach_watermark_giu_nguyen_noi_dung_that(self):
        import fitz, shutil
        tmp = tempfile.mktemp(suffix=".pdf")
        shutil.copyfile(PDF_WATERMARKED, tmp)
        try:
            before = fitz.open(tmp)
            n_pages = len(before)
            text_before = [before[i].get_text() for i in range(n_pages)]
            before.close()

            removed = tvb.strip_view_watermark(tmp, "hungnv1.qld", lambda *a: None)
            self.assertEqual(removed, n_pages)   # 1 watermark/trang

            after = fitz.open(tmp)
            self.assertEqual(len(after), n_pages)
            for i in range(n_pages):
                text_after = after[i].get_text()
                self.assertNotIn("hungnv1.qld", text_after)
                # nội dung thật (trừ đúng dòng watermark) phải giữ nguyên
                only_wm_removed = "".join(
                    ln for ln in text_before[i].splitlines(keepends=True)
                    if "hungnv1.qld" not in ln)
                self.assertEqual(only_wm_removed, text_after)
            after.close()

            # gọi lại lần 2 trên file ĐÃ SẠCH — không tìm thấy gì để xoá nữa (idempotent)
            removed2 = tvb.strip_view_watermark(tmp, "hungnv1.qld", lambda *a: None)
            self.assertEqual(removed2, 0)
        finally:
            os.remove(tmp)

    @_skip_if_missing(PDF_WATERMARKED)
    def test_khong_xoa_neu_sai_ten_dang_nhap(self):
        import shutil
        tmp = tempfile.mktemp(suffix=".pdf")
        shutil.copyfile(PDF_WATERMARKED, tmp)
        try:
            removed = tvb.strip_view_watermark(tmp, "nguoi_khac.qld", lambda *a: None)
            self.assertEqual(removed, 0)
        finally:
            os.remove(tmp)


# ==================== flow_keyword_from_code (bỏ số, lấy hậu tố) ====================
class TestFlowKeywordFromCode(unittest.TestCase):
    def test_cases(self):
        cases = [
            ("936/CL", "CL"),
            ("2205/QĐ-BYT", "BYT"),
            ("/GM-QLD", "QLD"),
            ("123", None),
            ("", None),
            (None, None),
            ("456/ATTP", "ATTP"),
        ]
        for code, expected in cases:
            with self.subTest(code=code):
                self.assertEqual(tvb.flow_keyword_from_code(code), expected)


# ==================== flow_id_for_code / flow_id_for_doc ====================
class TestFlowIdForCode(unittest.TestCase):
    def setUp(self):
        self.store = {
            "pinned": ["24702", "24682", "24703"],
            "freq": {},
            "rules": [
                {"keyword": "byt", "flowId": "24702"},
                {"keyword": "qld", "flowId": "24682"},
                {"keyword": "cl", "flowId": "24703"},
            ],
            "doc_type_rules": {"Giấy chứng nhận": "24702"},
            "default_flow_id": "24703",
        }

    def test_byt_uu_tien_cao_nhat(self):
        self.assertEqual(tvb.flow_id_for_code("123/QĐ-BYT", self.store), "24702")

    def test_qld(self):
        self.assertEqual(tvb.flow_id_for_code("456/QLD-CL", self.store), "24682")

    def test_cl(self):
        self.assertEqual(tvb.flow_id_for_code("789/CL-TB", self.store), "24703")

    def test_khong_khop_roi_ve_mac_dinh(self):
        self.assertEqual(tvb.flow_id_for_code("xyz", self.store), "24703")

    def test_doc_type_forced(self):
        self.assertEqual(tvb.flow_id_for_doc("Giấy chứng nhận", "bất kỳ", self.store), "24702")

    def test_doc_type_khong_forced_roi_xet_ky_hieu(self):
        self.assertEqual(tvb.flow_id_for_doc("Công văn", "123/QĐ-BYT", self.store), "24702")


# ==================== _migrate_old_flow_list ====================
class TestMigrateOldFlowList(unittest.TestCase):
    def test_migrate_giu_dung_hanh_vi_cu(self):
        old = {"flows": [
            {"name": "Luồng Cục", "flowId": "24682"},
            {"name": "Luồng Phòng", "flowId": "24703"},
            {"name": "Luồng Bộ", "flowId": "24702"},
        ]}
        new = tvb._migrate_old_flow_list(old)
        self.assertEqual(set(new["pinned"]), {"24682", "24703", "24702"})
        rules_by_kw = {r["keyword"]: r["flowId"] for r in new["rules"]}
        self.assertEqual(rules_by_kw.get("byt"), "24702")
        self.assertEqual(rules_by_kw.get("qld"), "24682")
        self.assertEqual(rules_by_kw.get("cl"), "24703")
        self.assertEqual(new["default_flow_id"], "24703")
        self.assertEqual(new["doc_type_rules"].get("Giấy chứng nhận"), "24702")


# ==================== build_role_number_map (order = số đóng dấu) ====================
class TestBuildRoleNumberMap(unittest.TestCase):
    """Dữ liệu lấy nguyên từ HAR thật ('chọn 3 luồng sẵn.har') — đã xác nhận khớp 100% với
    bảng số cũ từng hardcode (sig_role_number)."""

    def test_luong_phong(self):
        items = [
            {"order": 1, "actionType": 4, "roleName": "Trưởng phòng"},
            {"order": 2, "actionType": 3, "roleName": "Văn thư phòng"},
        ]
        m = tvb.build_role_number_map(items)
        self.assertEqual(m.get("TRƯỞNG PHÒNG"), 1)
        self.assertEqual(m.get("CHUYÊN VIÊN"), 0)
        self.assertNotIn("VĂN THƯ PHÒNG", m)   # actionType=3 (ban hành) không phải bước ký

    def test_luong_cuc(self):
        items = [
            {"order": 1, "actionType": 1, "roleName": "Trưởng phòng"},
            {"order": 2, "actionType": 2, "roleName": "Văn thư đơn vị"},
            {"order": 3, "actionType": 4, "roleName": "Phó Cục trưởng"},
            {"order": 4, "actionType": 3, "roleName": "Văn thư đơn vị"},
        ]
        m = tvb.build_role_number_map(items)
        self.assertEqual(m.get("TRƯỞNG PHÒNG"), 1)
        self.assertEqual(m.get("PHÓ CỤC TRƯỞNG"), 3)

    def test_luong_bo(self):
        items = [
            {"order": 1, "actionType": 1, "roleName": "Trưởng phòng"},
            {"order": 2, "actionType": 1, "roleName": "Phó Cục trưởng"},
            {"order": 3, "actionType": 2, "roleName": "Văn thư đơn vị"},
            {"order": 4, "actionType": 2, "roleName": "Văn thư Bộ"},
            {"order": 5, "actionType": 5, "roleName": "Thứ trưởng"},
            {"order": 6, "actionType": 3, "roleName": "Văn thư Bộ"},
        ]
        m = tvb.build_role_number_map(items)
        self.assertEqual(m.get("TRƯỞNG PHÒNG"), 1)
        self.assertEqual(m.get("PHÓ CỤC TRƯỞNG"), 2)
        self.assertEqual(m.get("THỨ TRƯỞNG"), 5)


# ==================== _sig_find_hits (so khớp chức danh 2 chiều) ====================
class TestSigFindHits(unittest.TestCase):
    def setUp(self):
        self.role_map = {"CHUYÊN VIÊN": 0, "TRƯỞNG PHÒNG": 1, "PHÓ CỤC TRƯỞNG": 3, "THỨ TRƯỞNG": 5}

    def _mk(self, text):
        return {"text": text, "x0": 0, "y0": 0, "x1": 100, "y1": 10}

    def test_khop_dung_nguyen_van(self):
        hits = tvb._sig_find_hits([self._mk("PHÓ CỤC TRƯỞNG")], self.role_map)
        self.assertEqual(hits, [(0, "PHÓ CỤC TRƯỞNG")])

    def test_dong_in_ngan_hon_role(self):
        # PDF chỉ in "CỤC TRƯỞNG" (không có "PHÓ") — vẫn phải khớp được role dài hơn
        hits = tvb._sig_find_hits([self._mk("CỤC TRƯỞNG")], self.role_map)
        self.assertEqual(hits, [(0, "PHÓ CỤC TRƯỞNG")])

    def test_khong_khop_gi_ca(self):
        hits = tvb._sig_find_hits([self._mk("GIÁM ĐỐC SỞ Y TẾ")], self.role_map)
        self.assertEqual(hits, [])

    def test_muc_so_la_ma_khong_bi_coi_la_chuc_danh(self):
        # Ca lỗi thực tế: mục lớn "I." và "V." đứng riêng 1 dòng trong văn bản — CHỮ HOA hết
        # nên qua được _is_upper_title(), nhưng "I"/"V" tình cờ là 1 ký tự con nằm trong
        # "CHUYÊN VIÊN" → từng bị khớp nhầm thành chức danh, cả 2 cùng gán số 0 (báo TRÙNG số
        # giả). Giờ phải bị loại hết, không phát sinh hit nào.
        hits = tvb._sig_find_hits([self._mk("I"), self._mk("V"), self._mk("X"), self._mk("II.")],
                                   self.role_map)
        self.assertEqual(hits, [])

    def test_dong_qua_ngan_khong_khop_kieu_nam_trong_nhau(self):
        # Dòng ngắn (< SIG_MIN_TITLE_LEN) không phải số La Mã nhưng cũng không đủ dài để so
        # khớp kiểu "nằm trong nhau" — chỉ còn khớp NGUYÊN VĂN mới được tính.
        hits = tvb._sig_find_hits([self._mk("ABC")], self.role_map)
        self.assertEqual(hits, [])


# ==================== preferred_signer / remember_signer_pick ====================
class TestSignerPreference(unittest.TestCase):
    """remember_signer_pick() ghi ra đĩa (save_flow_store) — dùng `path=` trỏ sang 1 file TẠM
    (xoá sau mỗi test), không bao giờ chạm tới luong_trinh.json thật của máy đang chạy. Đây
    chính là tham số đã thêm để sửa lỗi phát hiện được ở lần chạy test đầu tiên (từng ghi đè
    mất sổ thật do hàm cũ không nhận đường dẫn khác)."""
    def setUp(self):
        self.store = {"pinned": [], "freq": {}, "rules": [], "doc_type_rules": {}, "default_flow_id": None}
        self.cands = [{"userId": 501, "fullName": "A"}, {"userId": 502, "fullName": "B"},
                      {"userId": 503, "fullName": "C"}]
        fd, self.tmp_path = tempfile.mkstemp(suffix=".json")
        os.close(fd)

    def tearDown(self):
        os.remove(self.tmp_path)

    def test_chua_co_lich_su(self):
        self.assertIsNone(tvb.preferred_signer(self.store, "21323", 500028347, self.cands))

    def test_uu_tien_lan_gan_nhat(self):
        tvb.remember_signer_pick(self.store, "21323", 500028347, 502, path=self.tmp_path)
        tvb.remember_signer_pick(self.store, "21323", 500028347, 502, path=self.tmp_path)
        tvb.remember_signer_pick(self.store, "21323", 500028347, 503, path=self.tmp_path)
        self.assertEqual(tvb.preferred_signer(self.store, "21323", 500028347, self.cands), 503)
        # xác nhận đã ghi ra ĐÚNG file tạm — không phải file thật
        with open(self.tmp_path, encoding="utf-8") as f:
            saved = json.load(f)
        self.assertEqual(saved["signer_pref"]["21323:500028347"]["last"], 503)

    def test_roi_ve_tan_suat_khi_lan_gan_nhat_khong_con(self):
        tvb.remember_signer_pick(self.store, "21323", 500028347, 502, path=self.tmp_path)
        tvb.remember_signer_pick(self.store, "21323", 500028347, 502, path=self.tmp_path)
        tvb.remember_signer_pick(self.store, "21323", 500028347, 503, path=self.tmp_path)
        cands2 = [{"userId": 501, "fullName": "A"}, {"userId": 502, "fullName": "B"}]
        self.assertEqual(tvb.preferred_signer(self.store, "21323", 500028347, cands2), 502)


# ==================== _parse_select_options (đọc HTML prepareInsert.do) ====================
class TestParseSelectOptions(unittest.TestCase):
    @_skip_if_missing(HAR_PREPARE_INSERT)
    def test_flows_va_profiles(self):
        har = json.load(open(HAR_PREPARE_INSERT, encoding="utf-8"))
        resp_text = None
        for e in har["log"]["entries"]:
            if "prepareInsert.do" in e["request"]["url"]:
                resp_text = e["response"]["content"].get("text")
                break
        self.assertIsNotNone(resp_text, "HAR mẫu không có response text cho prepareInsert.do")
        flows = tvb._parse_select_options(resp_text, "flowAsignId")
        profiles = tvb._parse_select_options(resp_text, "profileFlowAsignId")
        self.assertEqual(len(flows), 34)
        self.assertEqual(len(profiles), 17)
        self.assertNotIn("-1", [f["id"] for f in flows])   # dòng "---Chọn---" phải bị loại


# ==================== _sig_tagged_path (đường dẫn file đã đánh số — thư mục tạm, giữ nguyên tên) ====================
class TestSigTaggedPath(unittest.TestCase):
    def test_giu_nguyen_ten_trong_thu_muc_tam_rieng(self):
        out = tvb._sig_tagged_path("/tmp/CV gửi Vụ PC.pdf")
        self.assertEqual(os.path.basename(out), "CV gửi Vụ PC.pdf")   # giữ nguyên tên gốc
        self.assertNotEqual(os.path.dirname(out), "/tmp")   # không còn nằm cạnh file gốc
        self.assertIn("voffice_stamp_", os.path.dirname(out))   # thư mục tạm riêng do chương trình tạo
        self.assertTrue(os.path.isdir(os.path.dirname(out)))


# ==================== _match_report (khớp phiếu trình vừa lưu — xác minh "đã trình") ====================
class TestMatchReport(unittest.TestCase):
    def setUp(self):
        from datetime import datetime
        self.since = datetime(2026, 7, 31, 15, 25, 30)

    def test_khop_dung_content_creator_va_thoi_gian(self):
        items = [
            {"reportId": 500594708, "content": "thử nghiệm 1", "creatorId": 500013302,
             "createdDate": "2026-07-31T15:25:40"},
            {"reportId": 500591063, "content": "rà soát Hồ sơ", "creatorId": 500013302,
             "createdDate": "2026-07-23T12:16:19"},
        ]
        item = tvb._match_report(items, "thử nghiệm 1", 500013302, self.since)
        self.assertIsNotNone(item)
        self.assertEqual(item["reportId"], 500594708)

    def test_khong_khop_neu_khac_creator(self):
        items = [{"reportId": 1, "content": "thử nghiệm 1", "creatorId": 999,
                  "createdDate": "2026-07-31T15:25:40"}]
        self.assertIsNone(tvb._match_report(items, "thử nghiệm 1", 500013302, self.since))

    def test_khop_du_thieu_creator_id_neu_khong_lay_duoc_danh_tinh(self):
        # creator_id=None (không lấy được danh tính tài khoản) — vẫn phải khớp được theo
        # content + thời gian, không bỏ cuộc hoàn toàn chỉ vì thiếu 1 lớp so khớp phụ.
        items = [{"reportId": 1, "content": "thử nghiệm 1", "creatorId": 500013302,
                  "createdDate": "2026-07-31T15:25:40"}]
        item = tvb._match_report(items, "thử nghiệm 1", None, self.since)
        self.assertIsNotNone(item)
        self.assertEqual(item["reportId"], 1)

    def test_bo_qua_phieu_cu_trung_noi_dung_tao_truoc_since(self):
        # Cùng content, cùng creatorId, nhưng tạo TRƯỚC lúc mình gọi lưu — không phải bản vừa
        # tạo (vd văn bản định kỳ dùng lại đúng 1 câu nội dung mỗi lần).
        items = [{"reportId": 42, "content": "báo cáo tuần", "creatorId": 500013302,
                  "createdDate": "2026-07-20T09:00:00"}]
        self.assertIsNone(tvb._match_report(items, "báo cáo tuần", 500013302, self.since))

    def test_khong_khop_neu_khac_noi_dung(self):
        items = [{"reportId": 1, "content": "khác hẳn", "creatorId": 500013302,
                  "createdDate": "2026-07-31T15:25:40"}]
        self.assertIsNone(tvb._match_report(items, "thử nghiệm 1", 500013302, self.since))

    def test_khop_theo_report_id_du_createdDate_cu_hon_since(self):
        # Sửa phiếu trình cũ: createdDate là ngày TẠO GỐC (trước since_dt rất xa) — nếu còn dùng
        # bộ lọc thời gian thì sẽ bị loại nhầm; có expect_report_id thì phải khớp thẳng theo ID,
        # bỏ qua hẳn since_dt.
        items = [{"reportId": 42, "content": "báo cáo tuần", "creatorId": 500013302,
                  "createdDate": "2026-07-20T09:00:00"}]
        item = tvb._match_report(items, "báo cáo tuần", 500013302, self.since, expect_report_id=42)
        self.assertIsNotNone(item)
        self.assertEqual(item["reportId"], 42)

    def test_khong_khop_report_id_khac_du_noi_dung_va_thoi_gian_khop(self):
        # Có expect_report_id thì content/creator/thời gian không còn ý nghĩa gì — id khác là
        # không khớp, dù mọi thứ khác đều khớp.
        items = [{"reportId": 99, "content": "thử nghiệm 1", "creatorId": 500013302,
                  "createdDate": "2026-07-31T15:25:40"}]
        self.assertIsNone(tvb._match_report(items, "thử nghiệm 1", 500013302, self.since,
                                             expect_report_id=42))


# ==================== search_incoming_docs (tra cứu văn bản đến — tab "Quản lý văn bản đến") ====================
HAR_INCOMING = "/Users/hnguyen/Downloads/000. Xem xoá/xem văn bản đến và search.har"


class _FakeResp:
    def __init__(self, text="", json_data=None, url="https://emoh.moh.gov.vn/x", status=200):
        self.text = text
        self._json = json_data if json_data is not None else {}
        self.url = url
        self.status_code = status

    def json(self):
        return self._json


class _FakeSession:
    """Bắt đúng payload gửi đi, không đụng mạng. `.get` trả token giả; `.post` ghi lại
    params/data rồi trả JSON giả."""
    def __init__(self):
        self.last_post = None

    def get(self, url, params=None, timeout=None):
        return _FakeResp(text="TOKEN " + "A1B2C3D4E5F6G7H8J9K0LMNP", url=url)

    def post(self, url, params=None, data=None, timeout=None):
        self.last_post = {"url": url, "params": params, "data": data}
        return _FakeResp(json_data={"items": [{"documentReceiveId": 1}], "totalRows": 496}, url=url)


class TestSearchIncomingDocs(unittest.TestCase):
    def test_nam_nhan_suy_ra_bookYear_va_khoang_ngay(self):
        s = _FakeSession()
        items, total = tvb.search_incoming_docs(s, year="2025", doc_abstract="thông tư",
                                                 doc_code="1375", publisher_name="pháp chế",
                                                 start=0, count=50)
        self.assertEqual(total, 496)
        self.assertEqual(len(items), 1)
        d = s.last_post["data"]
        self.assertEqual(d["documentForm.bookYear"], "2025")
        self.assertEqual(d["documentForm.receiveDateFrom"], "2025-01-01")
        self.assertEqual(d["documentForm.receiveDateTo"], "2025-12-31")
        self.assertEqual(d["documentForm.documentAbstract"], "thông tư")
        self.assertEqual(d["documentForm.documentCode"], "1375")
        self.assertEqual(d["documentForm.publisherName"], "pháp chế")
        # token phải nằm trên query string, không phải trong body (đúng như HAR thật)
        self.assertEqual(s.last_post["params"]["struts.token.name"], "token")
        self.assertEqual(s.last_post["params"]["token"], "A1B2C3D4E5F6G7H8J9K0LMNP")

    def test_bo_trong_nam_thi_dung_nam_hien_tai(self):
        from datetime import datetime
        s = _FakeSession()
        tvb.search_incoming_docs(s, year="")
        y = str(datetime.now().year)
        self.assertEqual(s.last_post["data"]["documentForm.bookYear"], y)
        self.assertEqual(s.last_post["data"]["documentForm.receiveDateFrom"], f"{y}-01-01")

    def test_phan_trang_truyen_start(self):
        s = _FakeSession()
        tvb.search_incoming_docs(s, year="2026", start=50, count=50)
        self.assertEqual(s.last_post["data"]["start"], 50)

    @_skip_if_missing(HAR_INCOMING)
    def test_payload_khop_bo_khoa_cua_HAR_that(self):
        har = json.load(open(HAR_INCOMING, encoding="utf-8"))
        har_keys = None
        for e in har["log"]["entries"]:
            if "searchStaffMonitorDocument.do" in e["request"]["url"]:
                from urllib.parse import parse_qsl
                har_keys = {k for k, _ in parse_qsl(e["request"]["postData"]["text"], keep_blank_values=True)}
                break
        self.assertIsNotNone(har_keys, "HAR mẫu không có request searchStaffMonitorDocument.do")
        s = _FakeSession()
        tvb.search_incoming_docs(s, year="2025")
        self.assertEqual(set(s.last_post["data"].keys()), har_keys)


HAR_VIEW_DOC = "/Users/hnguyen/Downloads/000. Xem xoá/xem cụ thể văn bản.har"


class TestFetchIncomingDocComments(unittest.TestCase):
    def test_goi_dung_endpoint_va_tra_items(self):
        s = _FakeSession()

        def _post(url, params=None, data=None, timeout=None):
            s.last_post = {"url": url, "params": params, "data": data}
            return _FakeResp(json_data={"items": [{"commentText": "ok"}], "totalRows": 1})

        s.post = _post
        items = tvb.fetch_incoming_doc_comments(s, 504697758)
        self.assertEqual(items, [{"commentText": "ok"}])
        self.assertTrue(s.last_post["url"].endswith("/assignDoc!getComments.do"))
        self.assertEqual(s.last_post["params"], {"objectId": 504697758, "objectType": 1})

    def test_loi_mang_tra_list_rong_khong_nem(self):
        s = _FakeSession()

        def _boom(*a, **k):
            raise RuntimeError("mạng lỗi")

        s.post = _boom
        self.assertEqual(tvb.fetch_incoming_doc_comments(s, 1), [])

    @_skip_if_missing(HAR_VIEW_DOC)
    def test_doc_duoc_response_getComments_that(self):
        har = json.load(open(HAR_VIEW_DOC, encoding="utf-8"))
        payload = None
        for e in har["log"]["entries"]:
            if "assignDoc!getComments.do" in e["request"]["url"]:
                payload = json.loads(e["response"]["content"]["text"])
                break
        self.assertIsNotNone(payload, "HAR mẫu không có assignDoc!getComments.do")
        items = payload.get("items") or []
        self.assertTrue(items and "commentText" in items[0] and "userName" in items[0])


# ==================== Chạy ngầm dưới khay (Windows) — phần thuần, chạy được trên mọi HĐH ====================
class TestTrayHelpers(unittest.TestCase):
    def test_launch_command_co_co_tray_va_dat_trong_ngoac_kep(self):
        cmd = tvb._tray_launch_command()
        self.assertIn("--tray", cmd)
        self.assertTrue(cmd.strip().startswith('"'))   # đường dẫn exe/python luôn bọc "..."

    def test_autostart_no_op_ngoai_windows(self):
        if tvb.IS_WINDOWS:
            self.skipTest("chạy trên Windows — set_autostart có tác dụng thật")
        self.assertFalse(tvb.autostart_enabled())
        self.assertFalse(tvb.set_autostart(True))
        self.assertFalse(tvb.set_autostart(False))

    def test_single_instance_guard_ban_dau_ok_ban_hai_bi_chan(self):
        import threading, socket
        # Dùng 1 cổng trống ngẫu nhiên thay cho SINGLE_INSTANCE_PORT thật (máy CI/dev có thể
        # đang chiếm đúng cổng đó bởi tiến trình khác — không liên quan tới thứ đang kiểm thử).
        probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        probe.bind(("127.0.0.1", 0))
        free_port = probe.getsockname()[1]
        probe.close()
        orig_port = tvb.SINGLE_INSTANCE_PORT
        tvb.SINGLE_INSTANCE_PORT = free_port
        self.addCleanup(setattr, tvb, "SINGLE_INSTANCE_PORT", orig_port)

        srv, first = tvb.single_instance_guard()
        self.assertTrue(first)
        self.assertIsNotNone(srv)
        try:
            def responder():
                srv.setblocking(True)
                srv.settimeout(2.0)
                try:
                    c, _ = srv.accept()
                    c.recv(64)
                    c.sendall(tvb.SINGLE_INSTANCE_MAGIC)
                    c.close()
                except OSError:
                    pass
            t = threading.Thread(target=responder, daemon=True)
            t.start()
            srv2, first2 = tvb.single_instance_guard()
            t.join(timeout=3)
            self.assertIsNone(srv2)
            self.assertFalse(first2)
        finally:
            srv.close()

    def test_make_tray_image_dung_kich_thuoc(self):
        try:
            img = tvb.make_tray_image(40)
        except ImportError:
            self.skipTest("thiếu Pillow trên máy đang chạy")
        self.assertEqual(img.size, (40, 40))


# ==================== Rà soát AI (parse findings, đọc file, gọi Gemini) ====================
class TestParseAiFindings(unittest.TestCase):
    def test_mang_rong_va_khong_co_mang(self):
        self.assertEqual(tvb.parse_ai_findings("[]"), [])
        self.assertEqual(tvb.parse_ai_findings("khong co gi o day"), [])
        self.assertEqual(tvb.parse_ai_findings(""), [])

    def test_go_hang_rao_code_va_chu_thua(self):
        raw = "```json\n[{\"nhom\":\"chinh_ta\",\"mo_ta\":\"x\"},{\"nhom\":\"noi_nhan\"}]\n```  xong."
        out = tvb.parse_ai_findings(raw)
        self.assertEqual(len(out), 2)
        self.assertEqual(out[0]["nhom"], "chinh_ta")

    def test_loai_phan_tu_khong_phai_dict(self):
        self.assertEqual(tvb.parse_ai_findings('["a", 1, {"nhom":"x"}]'), [{"nhom": "x"}])

    def test_mang_khong_dong_thi_nem(self):
        with self.assertRaises(ValueError):
            tvb.parse_ai_findings("[{\"nhom\":\"x\"")

    def test_json_hong_han_thi_nem(self):
        with self.assertRaises(Exception):
            tvb.parse_ai_findings('[{"nhom":}]')


class TestReadAnyDocText(unittest.TestCase):
    def test_thieu_file_tra_rong(self):
        self.assertEqual(tvb.read_any_doc_text(""), "")
        self.assertEqual(tvb.read_any_doc_text("/khong/ton/tai.pdf"), "")


class TestGeminiReviewErrors(unittest.TestCase):
    def _fake_requests(self, status, text):
        import types

        class _R:
            status_code = status
            def __init__(s):
                s.text = text
            def json(s):
                import json as _j
                return _j.loads(text)

        mod = types.SimpleNamespace()
        mod.post = lambda *a, **k: _R()
        mod.RequestException = tvb.requests.RequestException
        return mod

    def test_key_sai(self):
        orig = tvb.requests
        tvb.requests = self._fake_requests(400, '{"error":{"message":"API key not valid"}}')
        try:
            with self.assertRaises(RuntimeError) as c:
                tvb.gemini_review("k", "m", "p")
            self.assertIn("API key", str(c.exception))
        finally:
            tvb.requests = orig

    def test_qua_han_muc(self):
        orig = tvb.requests
        tvb.requests = self._fake_requests(429, "quota")
        try:
            with self.assertRaises(RuntimeError) as c:
                tvb.gemini_review("k", "m", "p")
            self.assertIn("429", str(c.exception))
        finally:
            tvb.requests = orig

    def test_thanh_cong_tra_text(self):
        orig = tvb.requests
        tvb.requests = self._fake_requests(
            200, '{"candidates":[{"content":{"parts":[{"text":"[]"}]}}]}')
        try:
            self.assertEqual(tvb.gemini_review("k", "m", "p"), "[]")
        finally:
            tvb.requests = orig


if __name__ == "__main__":
    unittest.main(verbosity=2)
