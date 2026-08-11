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


# ==================== _sig_tagged_path (tên file đánh số theo ngày) ====================
class TestSigTaggedPath(unittest.TestCase):
    def test_dinh_dang_ngay(self):
        import re
        out = tvb._sig_tagged_path("/tmp/CV gửi Vụ PC.pdf")
        self.assertRegex(os.path.basename(out), r"^CV gửi Vụ PC - \d{2}\.\d{2}\.\d{2}\.pdf$")


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


if __name__ == "__main__":
    unittest.main(verbosity=2)
