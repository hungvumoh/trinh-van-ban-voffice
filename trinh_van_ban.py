# -*- coding: utf-8 -*-
# Tác giả: Nguyễn Vũ Hùng — Cục Quản lý Dược — 2026
"""
Tự động soạn + lưu NHÁP (hoặc TRÌNH thẳng) phiếu trình trên emoh.moh.gov.vn
- KHÔNG đăng nhập: dùng lại cookie phiên từ trình duyệt (bạn dán vào).
- Khung Xem trước có 2 nút: "Lưu dự thảo" (sign=0, mặc định, an toàn — chỉ lưu, không đi vào
  luồng ký duyệt) và "Trình văn bản" (sign=1 — trình thật, văn bản bắt đầu đi vào luồng ký
  duyệt ngay). Chọn "Lưu dự thảo" thì tự mở thùng nháp trên web, xem lại, tự bấm Trình sau.

Cần: Python 3.9+, thư viện requests  ->  pip install requests
Đặt 3 file cùng thư mục: trinh_van_ban.py, du_lieu.json, noi_nhan.json
Chạy:  python trinh_van_ban.py
"""
import base64, json, os, re, subprocess, sys, tempfile, time, threading, unicodedata
from html import unescape as html_unescape   # tên "html" đã dùng làm biến cục bộ khắp file (nội dung
                                              # trang) — import tách riêng để khỏi đụng nhau
from contextlib import contextmanager
from functools import lru_cache
from datetime import datetime, timedelta
from urllib.parse import quote, unquote, urljoin
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import requests
except ImportError:
    raise SystemExit("Chưa có thư viện 'requests'. Chạy: pip install requests")

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None   # vẫn chạy được, chỉ tắt tính năng tự đọc PDF

try:
    import fitz   # PyMuPDF — dùng để tìm toạ độ chữ ký + ghi chú thích số thứ tự
except ImportError:
    fitz = None

try:
    import docx2pdf   # chuyển .docx -> .pdf (điều khiển Word cài sẵn trên máy) — không hỗ trợ .doc
except ImportError:
    docx2pdf = None

try:
    import docx   # python-docx — đọc thẳng .docx (không cần Word) để tự điền NGAY trong lúc
                   # đang chờ Word chuyển .docx -> PDF chạy song song ở luồng khác (xem
                   # convert_office_doc_to_pdf) — trước đây phải đợi PDF xong mới tách được chữ
except ImportError:
    docx = None

AUTHOR_MARK = "Nguyễn Vũ Hùng  ·  Cục Quản lý Dược  ·  2026"

BASE = "https://emoh.moh.gov.vn"
CAS = "https://emoh.moh.gov.vn:8443/passportv3"
LOGIN_URL = CAS + "/login"
CAPTCHA_URL = CAS + "/jcaptcha.jpg"
SERVICE = "https://emoh.moh.gov.vn"
# Đóng gói bằng PyInstaller (onefile) thì __file__ trỏ vào thư mục giải nén tạm (_MEIPASS),
# không phải nơi đặt file .exe — phải dùng sys.executable để tìm đúng thư mục chứa du_lieu.json,
# noi_nhan.json... (và để settings.json/id_cache.json/nguoi_dung.json ghi đúng chỗ, không mất
# khi thư mục tạm bị dọn sau khi tắt chương trình).
if getattr(sys, "frozen", False):
    HERE = os.path.dirname(os.path.abspath(sys.executable))
    # Trên macOS, sys.executable nằm trong TênApp.app/Contents/MacOS/ — nếu để nguyên thì các
    # file .json phải nhét vào tận trong ruột app (phải bấm "Show Package Contents" mới thấy),
    # bất tiện khi copy nguyên thư mục sang máy khác. Đi lên khỏi .app để dữ liệu nằm CẠNH
    # file .app trong Finder, giống hệt cách .exe + .json nằm cạnh nhau trên Windows.
    if sys.platform == "darwin" and "/Contents/MacOS" in HERE:
        HERE = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
else:
    HERE = os.path.dirname(os.path.abspath(__file__))

# Header kiểu trình duyệt (điều hướng, KHÔNG phải XHR) — dùng cho login & upload
BROWSER_HEADERS = {
    "X-Requested-With": None,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
}

# ---------- Nạp dữ liệu ----------
def load_json(name):
    with open(os.path.join(HERE, name), encoding="utf-8") as f:
        return json.load(f)

DATA = load_json("du_lieu.json")
BOOK = load_json("noi_nhan.json")
try:
    CAY = load_json("cay_don_vi.json")     # cây đầy đủ: internal + lien_thong, path neo gốc
except Exception:
    CAY = {"internal": {"root_id": "52", "nodes": []}, "lien_thong": {"nodes": []}}
ENUMS = DATA["enums"]
TPL = DATA["templates"]
# Tên loại văn bản viết HOA (khớp dòng tiêu đề in trên trang 1, VD "QUYẾT ĐỊNH") -> tên gốc
# trong ENUMS["documentType"] (VD "Quyết định"), dùng để tự nhận loại VB từ file PDF.
DOC_TYPE_TITLES = {name.upper(): name for name in ENUMS["documentType"]}

FLOW_STORE_FILE = os.path.join(HERE, "luong_trinh.json")

def _migrate_old_flow_list(old):
    """Định dạng cũ: {"flows": [{"name": "Luồng Cục"/"Luồng Phòng"/"Luồng Bộ", "flowId": ...}]}
    — 3 luồng CỨNG dùng chung cho MỌI người chạy chương trình này, dù đó là 3 luồng CÁ NHÂN chỉ
    của riêng người tạo file. Chuyển sang ghim + quy tắc cục bộ (giữ nguyên hành vi hiện tại cho
    máy này, nhưng máy người khác sẽ có sổ ghim RIÊNG, rỗng lúc mới cài — xem thảo luận)."""
    by_name = {fl.get("name"): fl.get("flowId") for fl in old.get("flows", [])}
    bo, cuc, phong = by_name.get("Luồng Bộ"), by_name.get("Luồng Cục"), by_name.get("Luồng Phòng")
    rules = []
    if bo:   rules.append({"keyword": "byt", "flowId": bo})
    if cuc:  rules.append({"keyword": "qld", "flowId": cuc})
    if phong: rules.append({"keyword": "cl", "flowId": phong})
    return {
        "pinned": [x for x in (bo, cuc, phong) if x],
        "freq": {},
        "rules": rules,
        "doc_type_rules": ({"Giấy chứng nhận": bo} if bo else {}),
        "default_flow_id": phong,
    }

def load_flow_store():
    """Sổ ghim/tần suất luồng trình + quy tắc tự nhận luồng theo ký hiệu văn bản — CỤC BỘ theo
    máy đang chạy (y hệt cơ chế ghim/tần suất "Nơi nhận" ở nguoi_dung.json), KHÔNG PHẢI danh
    sách luồng cứng dùng chung cho mọi người. Danh sách luồng thật luôn lấy động từ web
    (fetch_prepare_insert_data) — file này chỉ nhớ: luồng nào máy này hay dùng/đã ghim, và quy
    tắc "thấy chữ X trong ký hiệu thì chọn luồng nào" của riêng máy này."""
    try:
        s = load_json("luong_trinh.json")
    except Exception:
        s = {}
    migrated = "flows" in s and "pinned" not in s
    if migrated:
        s = _migrate_old_flow_list(s)
    s.setdefault("pinned", [])
    s.setdefault("freq", {})
    s.setdefault("rules", [])
    s.setdefault("doc_type_rules", {})
    s.setdefault("default_flow_id", None)
    if migrated:
        save_flow_store(s)   # ghi lại ngay định dạng mới — lần chạy sau khỏi phải migrate lại
    return s

def save_flow_store(store, path=None):
    """`path`: chỉ để test/gọi tay trỏ sang file khác — mặc định (None) luôn là sổ thật của máy
    đang chạy. Trước đây hàm này ghi cứng vào FLOW_STORE_FILE bất kể `store` truyền vào là gì,
    nên 1 dict tạm/giả lập (vd trong test) vẫn âm thầm ghi đè sổ thật — nay phải tự truyền
    path khác thì mới ghi chỗ khác được."""
    try:
        atomic_write_json(path or FLOW_STORE_FILE, store, ensure_ascii=False, indent=1)
    except Exception:
        pass

def bump_flow_freq(store, flow_id, path=None):
    store["freq"][flow_id] = store["freq"].get(flow_id, 0) + 1
    save_flow_store(store, path)

def flow_keyword_from_code(code):
    """Lấy phần chữ có Ý NGHĨA từ ký hiệu văn bản để HỌC quy tắc mới — bỏ số thứ tự (đổi theo
    từng văn bản, không lặp lại), chỉ giữ phần sau dấu '/' hoặc '-' CUỐI CÙNG (vd '936/CL' ->
    'CL', '2205/QĐ-BYT' -> 'BYT', '/GM-QLD' -> 'QLD'). Trả None nếu phần đó rỗng hoặc toàn số —
    không có gì lặp lại được để học thành quy tắc."""
    if not code:
        return None
    part = re.split(r"[/\-]", code)[-1].strip()
    if not part or part.isdigit():
        return None
    return part

def signer_pref_key(flow_id, node_id):
    return f"{flow_id}:{node_id}"

def remember_signer_pick(store, flow_id, node_id, user_id, path=None):
    """Nhớ người vừa được chọn cho 1 bước (flowId, nodeId) — dùng chung cho MỌI quy tắc/từ khoá
    dẫn tới cùng luồng+bước đó (không tách riêng theo từ khoá, tránh trùng lặp dữ liệu)."""
    if user_id is None:
        return
    key = signer_pref_key(flow_id, node_id)
    pref = store.setdefault("signer_pref", {}).setdefault(key, {"last": None, "freq": {}})
    pref["last"] = user_id
    uid_s = str(user_id)
    pref["freq"][uid_s] = pref["freq"].get(uid_s, 0) + 1
    save_flow_store(store, path)

def preferred_signer(store, flow_id, node_id, candidates):
    """Người nên chọn mặc định trong `candidates` (list dict có 'userId') theo lịch sử đã chọn
    cho đúng (flow_id, node_id) này — ưu tiên LẦN GẦN NHẤT, rồi tới tần suất cao nhất; None nếu
    chưa có lịch sử hoặc người cũ không còn trong danh sách ứng viên hiện tại (vd đã đổi vai trò)."""
    pref = store.get("signer_pref", {}).get(signer_pref_key(flow_id, node_id))
    if not pref:
        return None
    cand_ids = {c.get("userId") for c in candidates}
    if pref.get("last") in cand_ids:
        return pref["last"]
    for uid_s, _cnt in sorted(pref.get("freq", {}).items(), key=lambda kv: -kv[1]):
        try:
            uid = int(uid_s)
        except ValueError:
            continue
        if uid in cand_ids:
            return uid
    return None

# ---------- Tiện ích ----------
def now_ms():
    return str(int(time.time() * 1000))

def atomic_write_json(path, data, **kw):
    """Ghi JSON AN TOÀN: ghi ra file tạm CÙNG THƯ MỤC với `path` rồi mới os.replace() đè lên
    file đích — os.replace là 1 thao tác nguyên tử của hệ điều hành, nên nếu chương trình bị
    tắt/lỗi/mất điện GIỮA CHỪNG thì file đích vẫn còn nguyên bản CŨ (chưa từng bị đụng tới),
    không rơi vào tình trạng dở dang. Khác với việc ghi thẳng bằng open(path, "w") — cách đó
    XOÁ TRẮNG file đích ngay khi mở, nên nếu json.dump lỗi sau đó (đĩa đầy, encoding lỗi...)
    thì mất luôn dữ liệu cũ, dù lỗi bị 'except Exception: pass' nuốt lặng lẽ ở nơi gọi."""
    folder = os.path.dirname(path) or "."
    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", suffix=".json", dir=folder)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, **kw)
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise

def decode_params(pairs):
    """Giải mã tên & giá trị (khuôn lưu ở dạng URL-encoded)."""
    return [[unquote(n), unquote(v)] for n, v in pairs]

def set_param(params, name, value):
    for p in params:
        if p[0] == name:
            p[1] = value
            return
    params.append([name, value])

def set_suffix(params, suffix, value):
    """Ghi đè MỌI trường có tên kết thúc bằng suffix (kể cả có/không tiền tố)."""
    hit = False
    for p in params:
        if p[0] == suffix or p[0].endswith("." + suffix):
            p[1] = value
            hit = True
    return hit

class PipelineError(Exception):
    pass

@contextmanager
def step(log, n, total, label):
    """Bọc mỗi bước: báo bắt đầu / xong / mắc-ở-bước-nào."""
    log(f"\n▶ BƯỚC {n}/{total}: {label} …")
    t0 = time.time()
    try:
        yield
    except PipelineError as e:
        log(f"✖ MẮC Ở BƯỚC {n}/{total} — {label}\n   Lý do: {e}")
        raise
    except Exception as e:
        log(f"✖ MẮC Ở BƯỚC {n}/{total} — {label}\n   Lỗi: {e!r}")
        raise
    log(f"✓ Bước {n}/{total} xong ({time.time()-t0:.1f}s)")

def http_log(log, r):
    log(f"   ↳ HTTP {r.status_code} · {len(r.text)} ký tự · {r.url.split('?')[0].split('/')[-1]}")

# ---------- Phiên & Đăng nhập ----------
CHROME_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
             "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

def make_session():
    s = requests.Session()
    # Mặc định là header cho các lời gọi .do (XHR) SAU khi đã vào hệ thống.
    s.headers.update({
        "User-Agent": CHROME_UA,
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE,
        "Referer": BASE + "/Index.do",
    })
    return s

# Header cho bước ĐĂNG NHẬP: giống hệt gõ URL trên trình duyệt (KHÔNG origin/referer/xhr)
def _nav_headers():
    return {
        "User-Agent": CHROME_UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "vi,en-US;q=0.9,en;q=0.8",
        "Upgrade-Insecure-Requests": "1",
        # gỡ các header của app khỏi request điều hướng CAS
        "X-Requested-With": None, "Origin": None, "Referer": None,
    }

def _scrape_hidden(html):
    """Lấy các trường ẩn của form login (đặc biệt là 'lt'), chấp nhận mọi thứ tự thuộc tính."""
    fields = {}
    for m in re.finditer(r'<input[^>]+>', html):
        tag = m.group(0)
        if 'hidden' not in tag and 'name="lt"' not in tag:
            continue
        n = re.search(r'name="([^"]+)"', tag)
        v = re.search(r'value="([^"]*)"', tag)
        if n:
            fields[n.group(1)] = v.group(1) if v else ""
    return fields

def _is_login_page(html):
    return ('id="fm1"' in html) or ('name="lt"' in html)

def _is_interstitial(html):
    t = html.lower()
    return ("chưa thể đăng nhập" in t) or ("về trang đăng nhập" in t)

def _find_login_link(html):
    """Tìm link 'Về trang đăng nhập' trong trang đệm."""
    m = re.search(r'href="([^"]*passportv3/login[^"]*)"', html)
    if m:
        return m.group(1).replace("&amp;", "&")
    return LOGIN_URL + "?appCode=VOFFICE&service=" + SERVICE

def _scrape_form_action(html):
    m = re.search(r'<form[^>]*id="fm1"[^>]*action="([^"]*)"', html) \
        or re.search(r'<form[^>]*action="([^"]*passportv3/login[^"]*)"', html)
    return m.group(1).replace("&amp;", "&") if m else None

def _pass_waf(s, r, headers, log, depth=0):
    """Vượt lớp chống bot: trang trả về JS đặt cookie rồi reload.
    Tự bóc cookie, gắn vào phiên, tải lại URL. Trả về response mới."""
    txt = r.text
    if depth > 3 or 'document.cookie' not in txt:
        return r
    got = False
    for m in re.finditer(r'document\.cookie\s*=\s*"([^"=]+)=([^"+;]+)', txt):
        name, val = m.group(1), m.group(2)
        s.cookies.set(name, val, domain="emoh.moh.gov.vn")
        log(f"• Vượt lớp bảo vệ: gắn cookie {name}…")
        got = True
    if not got:
        return r
    r2 = s.get(r.url, headers=headers, timeout=30, allow_redirects=True)
    http_log(log, r2)
    return _pass_waf(s, r2, headers, log, depth + 1)   # lặp nếu còn tầng nữa

def _get(s, url, log, **kw):
    """GET có tự vượt lớp chống bot."""
    headers = kw.pop("headers", _nav_headers())
    r = s.get(url, headers=headers, timeout=30, allow_redirects=True, **kw)
    http_log(log, r)
    return _pass_waf(s, r, headers, log)

def _get_login_page(s, log):
    """Đi từ trang chủ để CAS tự dẫn tới form login. Trả về (html, url_form)."""
    log("• Vào trang chủ để hệ thống dẫn tới đăng nhập…")
    r = _get(s, BASE + "/", log)
    html, url = r.text, r.url

    if not _is_login_page(html) and _is_interstitial(html):
        link = _find_login_link(html)
        if link.startswith("/"):
            link = BASE + link
        log("• Gặp trang đệm → theo link 'Về trang đăng nhập'…")
        r = _get(s, link, log)
        html, url = r.text, r.url

    if not _is_login_page(html):
        log("• Chưa thấy form → thử mở thẳng trang đăng nhập CAS…")
        r = _get(s, LOGIN_URL + "?appCode=VOFFICE&service=" + SERVICE, log)
        html, url = r.text, r.url

    if not _is_login_page(html):
        log("   — Nội dung nhận được (để soi):")
        log("   " + " ".join(html.split())[:300])
        raise PipelineError("Không tới được form đăng nhập (xem nội dung ở trên).")
    return html, url

def _do_post(s, html, url_form, username, password, captcha, log):
    action = _scrape_form_action(html) or url_form
    action = urljoin(url_form, action)          # ghép tương đối vào ĐÚNG URL trang login (giữ cổng 8443)
    hidden = _scrape_hidden(html)
    if "lt" not in hidden:
        log("   ⚠ Không thấy vé 'lt' trong form.")
    data = dict(hidden)
    data.update({"username": username, "password": password, "captcha": captcha or "",
                 "_eventId": "submit", "submit": "ĐĂNG NHẬP"})
    data.setdefault("loginCount", "0")
    post_headers = _nav_headers()
    post_headers["Referer"] = url_form
    post_headers["Origin"] = "https://emoh.moh.gov.vn:8443"
    log(f"• Gửi thông tin đăng nhập → {action.split('?')[0]}")
    r = s.post(action, data=data, headers=post_headers, timeout=30, allow_redirects=True)
    http_log(log, r)
    r = _pass_waf(s, r, _nav_headers(), log)     # phòng khi phản hồi cũng bị WAF chặn
    return r

def cas_login(s, username, password, log, ask_captcha):
    """Đăng nhập CAS, nạp cookie vào session s. ask_captcha(path)->str hoặc None."""
    html, url_form = _get_login_page(s, log)

    r = _do_post(s, html, url_form, username, password, "", log)
    if _is_login_page(r.text):
        # Vẫn ở trang login → sai mật khẩu, hoặc cần captcha (loginCount>=3)
        if 'id="capchaRow"' in r.text and "display:none" not in r.text.split('id="capchaRow"')[1][:40]:
            need_captcha = True
        else:
            need_captcha = ("captcha" in r.text.lower() and "display:none" not in r.text)
        if not need_captcha:
            log("   — Phản hồi (để soi):")
            log("   " + " ".join(r.text.split())[:300])
            raise PipelineError("Đăng nhập thất bại: nhiều khả năng sai tài khoản/mật khẩu.")
        # Cần captcha
        log("• Cần captcha. Đang tải ảnh…")
        ci = s.get(CAPTCHA_URL, headers=_nav_headers(), timeout=30)
        path = os.path.join(HERE, "captcha.jpg")
        with open(path, "wb") as f:
            f.write(ci.content)
        code = ask_captcha(path)
        if not code:
            raise PipelineError("Chưa nhập captcha nên dừng.")
        # lấy lại form (lt mới) rồi post kèm captcha
        html, url_form = _get_login_page(s, log)
        r = _do_post(s, html, url_form, username, password, code, log)

    if not _is_login_page(r.text) and "passportv3/login" not in r.url:
        log("   → Đăng nhập thành công.")
        return True
    log("   — Phản hồi cuối (để soi):")
    log("   " + " ".join(r.text.split())[:300])
    raise PipelineError("Đăng nhập thất bại (sai mật khẩu hoặc captcha).")

# ---------- Các bước ----------
def open_forms(s, log):
    """Mở phiếu trình mới + form văn bản; trả về HTML form (chứa URL upload)."""
    log("• Mở phiếu trình mới (prepareInsert)…")
    s.post(BASE + "/voReport!prepareInsert.do", data={"dojo.preventCache": now_ms()}, timeout=30)
    log("• Mở form soạn văn bản (prepareCreateDraft)…")
    r = s.post(BASE + "/voPublishDocument!prepareCreateDraft.do",
               data={"dojo.preventCache": now_ms()}, timeout=30)
    http_log(log, r)
    html = r.text
    if "login" in r.url.lower() or "passport" in r.url.lower() or len(html) < 2000:
        raise PipelineError("Phiên có vẻ đã hết hạn (bị đẩy về đăng nhập). "
                            "Đăng nhập lại trên trình duyệt, lấy cookie mới rồi thử lại.")
    return html

def extract_upload_urls(html):
    """Lấy URL upload (kèm id mã hóa) cho từng ô kẹp file."""
    urls = {}
    # Ghép mỗi action với ô uploader gần nhất phía trước
    for m in re.finditer(r"getElementById\('browse(upload\w+)'\)[\s\S]{0,400}?action:\s*'([^']+)'", html):
        urls[m.group(1)] = m.group(2).replace("&amp;", "&")
    # Dự phòng: bắt trực tiếp các action upload theo thứ tự
    if not urls:
        found = [u.replace("&amp;", "&") for u in
                 re.findall(r"action:\s*'([^']*vou/file/upload[^']*)'", html)]
        if len(found) >= 1: urls["uploadReportFile"] = found[0]
        if len(found) >= 2: urls["uploadDraftFile"] = found[1]
    return urls

def upload(s, url, filepath, log):
    """Upload 1 file. Dựng multipart GIỐNG HỆT trình duyệt: MAX_FILE_SIZE rồi Filedata. Gửi
    NGUYÊN TÊN file gốc (kể cả dấu tiếng Việt/khoảng trắng) — xác nhận qua HAR thật: server
    lưu/hiển thị đúng tên có dấu bình thường, không cần đổi sang ASCII."""
    name = os.path.basename(filepath)
    log(f"• Upload '{name}'…")
    up_headers = {
        "X-Requested-With": None,
        "Accept": ("text/html,application/xhtml+xml,application/xml;q=0.9,"
                   "image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7"),
        "Referer": BASE + "/Index.do?request_locale=en_US&mainMenu=3&trId=2.0",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "iframe",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
    }
    with open(filepath, "rb") as fh:
        files = [
            ("MAX_FILE_SIZE", (None, "5")),                     # trường bắt buộc, đứng TRƯỚC (như trình duyệt)
            ("Filedata", (name, fh, "application/pdf")),
        ]
        r = s.post(url, files=files, headers=up_headers, timeout=120)
    http_log(log, r)
    txt = r.text
    # Response chứa postMessage([tên, "ID", "uploadXxx", token])
    m = re.search(r'\[\s*"[^"]*"\s*,\s*"(\d{4,})"', txt)
    if not m:
        m = re.search(r'"(\d{6,})"', txt) or re.search(r'(\d{6,})', txt)
    if not m:
        raise PipelineError("Không đọc được ID file từ kết quả upload.\n--- Phản hồi ---\n" + txt[:900])
    fid = m.group(1)
    log(f"   → ID file = {fid}")
    return fid

def upload_many(s, url, paths, log):
    """Upload nhiều file vào cùng ô. Trả về (chuỗi_id nối ';', id_file_chính)."""
    ids = []
    for fp in paths:
        ids.append(upload(s, url, fp, log))
    attach = ";".join(ids) + (";" if ids else "")   # server dùng dạng 'a;b;'
    main = ids[0] if ids else ""
    return attach, main

def reload_token(s, log):
    log("• Xin token mới (reloadToken)…")
    r = s.get(BASE + "/token!reloadToken.do", params={"dojo.preventCache": now_ms()}, timeout=30)
    http_log(log, r)
    m = re.search(r"[A-Z0-9]{24,}", r.text)
    if not m:
        raise PipelineError("Không tách được token từ reloadToken.\n--- Phản hồi ---\n" + repr(r.text[:300]))
    return m.group(0)

# ---------- Luồng trình (chuỗi người ký/duyệt) ----------
def _b64(s):
    return base64.b64encode((s or "").encode("utf-8")).decode("ascii")

def fetch_flow_nodes(s, flow_id, doc_abstract, log):
    """Lấy danh sách người ký/duyệt của 1 luồng trình (searchNodeInFlow)."""
    log(f"• Lấy danh sách người trong luồng trình (flowId={flow_id})…")
    url = (f"{BASE}/voReport!searchNodeInFlow.do?flowId={flow_id}"
           f"&token.getTokenParamString()&docAbs={quote(doc_abstract or '')}")
    r = s.post(url, data={"q": "*", "start": 0, "count": 200, "startval": 0}, timeout=30)
    http_log(log, r)
    try:
        data = r.json()
    except Exception:
        raise PipelineError("Không đọc được JSON luồng trình.\n--- Phản hồi ---\n" + r.text[:400])
    items = data.get("items") or []
    if not items:
        raise PipelineError(f"Luồng trình flowId={flow_id} không trả về ai cả (kiểm tra lại flowId).")
    log(f"   → {len(items)} người trong luồng.")
    return items

def build_flow_json(items):
    """Dựng lại chuỗi myJsonString (đúng định dạng server cần: tên/vai trò/đơn vị mã hoá base64)."""
    nodes = [{
        "deptId": n.get("deptId"), "userId": n.get("userId"), "roleId": n.get("roleId"),
        "actionType": n.get("actionType"), "nodeId": n.get("nodeId"), "order": n.get("order"),
        "roleName": _b64(n.get("roleName")), "fullName": _b64(n.get("fullName")),
        "deptName": _b64(n.get("deptName")),
    } for n in items]
    return json.dumps(nodes, ensure_ascii=False, separators=(",", ":"))

# ---------- Luồng trình KHÁC (lấy động từ web, ngoài 3 luồng quen) ----------
def _parse_select_options(html_text, select_name_substr):
    """Rút các <option value=ID>Tên</option> trong 1 <select> có tên chứa `select_name_substr`
    (dùng chung cho cả 2 select tĩnh mà prepareInsert.do nhúng sẵn trong HTML: danh sách luồng
    trình VÀ danh sách hồ sơ công việc — cả 2 đều theo đúng tài khoản đang đăng nhập)."""
    m = re.search(rf'<select[^>]*{re.escape(select_name_substr)}[^>]*>(.*?)</select>', html_text, re.S)
    if not m:
        return None
    out = []
    for opt_m in re.finditer(r'<option\b[^>]*\bvalue="([^"]*)"[^>]*>(.*?)</option>', m.group(1), re.S):
        val, name = opt_m.group(1).strip(), html_unescape(opt_m.group(2)).strip()
        if not val or val == "-1":
            continue   # dòng "---Chọn---"
        out.append({"id": val, "name": name})
    return out

def fetch_prepare_insert_data(s, log):
    """Gọi voReport!prepareInsert.do 1 lần, lấy 2 danh sách nhúng sẵn trong HTML theo ĐÚNG tài
    khoản đang đăng nhập — không cần sửa code/HAR mỗi khi tài khoản/luồng/hồ sơ khác nhau:
    - Luồng trình khác 3 luồng quen (<select name="reportForm.flowAsignId">).
    - Hồ sơ công việc (<select name="reportForm.profileFlowAsignId">) — BẮT BUỘC phải đúng của
      tài khoản đang trình, vì đây chính là chỗ trước đây bị hardcode theo hồ sơ của 1 người
      (xem save_report_draft/TPL) khiến phiếu trình tạo ra bị gắn nhầm hồ sơ khi đổi tài khoản
      khác — server vẫn báo "thành công" nhưng phiếu trình lạc mất, không thấy trong thùng nháp
      của tài khoản đang dùng."""
    log("• Lấy danh sách luồng trình + hồ sơ công việc từ web (prepareInsert)…")
    r = s.post(BASE + "/voReport!prepareInsert.do", data={"dojo.preventCache": now_ms()}, timeout=30)
    flow_opts = _parse_select_options(r.text, "flowAsignId") or []
    profile_opts = _parse_select_options(r.text, "profileFlowAsignId") or []
    flows = [{"flowId": o["id"], "name": o["name"]} for o in flow_opts]
    profiles = [{"fileId": o["id"], "name": o["name"]} for o in profile_opts]
    log(f"   → {len(flows)} luồng khác + {len(profiles)} hồ sơ công việc từ web.")
    return flows, profiles

def fetch_current_user_identity(s, log):
    """Lấy (userId, fullname) của TÀI KHOẢN ĐANG ĐĂNG NHẬP qua notice!getTop.do — xác nhận qua
    2 file HAR độc lập: field "userId"/"fullname" của message đầu tiên luôn khớp đúng người
    đang xem thông báo (không phải người được nhắc tới trong nội dung thông báo, đó là
    "userActionId"). Dùng để điền reportForm.creator/creatorId đúng theo tài khoản đang chạy,
    thay vì cứng theo 1 người (xem TPL["saveReport"]) — 2 trường này khác nhau giữa các tài
    khoản (đã kiểm chứng qua HAR) và có thể là lý do phiếu trình "lưu thành công" nhưng không
    hiện trong thùng nháp của tài khoản đang dùng (voReport!onSearchMyReport.do có vẻ lọc theo
    đúng danh tính phiên đăng nhập).
    Trả về None nếu không lấy được (vd tài khoản chưa có thông báo nào) — lúc đó gọi nơi dùng
    hàm này tự rơi về giá trị tĩnh cũ, không được tự bịa danh tính."""
    r = s.post(BASE + "/notice!getTop.do", data={"dojo.preventCache": now_ms()}, timeout=30)
    try:
        data = r.json()
    except Exception:
        log("   — Không đọc được danh tính tài khoản (notice!getTop.do lỗi JSON).")
        return None
    msg = data.get("message") or {}
    user_id, full_name = msg.get("userId"), msg.get("fullname")
    if not user_id or not full_name:
        log("   — Không lấy được danh tính tài khoản từ notice!getTop.do (có thể chưa có thông báo nào).")
        return None
    return user_id, full_name

def view_node_roles(s, flow_id, node_id):
    """Giải mã 1 bước của luồng (node) chưa có sẵn người: trả về TOÀN BỘ biến thể chức danh
    khả dĩ ở bước đó — list các (roleId, roleName, deptId, deptName). deptId đã được server tự
    quy đổi từ placeholder (-1=phòng người trình, -2=đơn vị người trình) sang phòng/đơn vị thật
    của người đang đăng nhập.
    LƯU Ý: dù nhiều biến thể chung 1 roleId (vd 4052 lặp lại cho cả "Cục Trưởng"/"Ký Văn
    Bản"/"Phó Cục Trưởng"), thử nghiệm thực tế cho thấy roleName CÓ ảnh hưởng tới danh sách
    người trả về ở viewPersonAsign.do (khác quan sát ban đầu từ dữ liệu mẫu) — nên phải tra
    riêng từng biến thể, không được chỉ lấy biến thể đầu tiên."""
    r = s.post(f"{BASE}/voReport!viewNode.do?flowId={flow_id}&nodeId={node_id}",
               data={"dojo.preventCache": now_ms()}, timeout=30)
    try:
        data = r.json()
    except Exception:
        raise PipelineError(f"Không đọc được vai trò/đơn vị cho bước (nodeId={node_id}).")
    items = data.get("items") or []
    if not items:
        raise PipelineError(f"Bước (nodeId={node_id}) không xác định được vai trò/đơn vị.")
    it = items[0]
    role_ids, role_names = it.get("roleId") or [], it.get("roleName") or []
    dept_ids, dept_names = it.get("deptId") or [], it.get("deptName") or []
    if not role_ids or not dept_ids:
        raise PipelineError(f"Bước (nodeId={node_id}) thiếu vai trò hoặc đơn vị.")
    dept_id = dept_ids[0]
    dept_name = dept_names[0] if dept_names else ""
    variants, seen = [], set()
    for i, rid in enumerate(role_ids):
        rname = role_names[i] if i < len(role_names) else ""
        key = (rid, rname)
        if key in seen:
            continue
        seen.add(key)
        variants.append({"roleId": rid, "roleName": rname, "deptId": dept_id, "deptName": dept_name})
    return variants

def view_person_candidates(s, role_id, dept_id, role_name):
    """Danh sách người đủ điều kiện giữ 1 vai trò tại 1 đơn vị (có thể 1 người, có thể nhiều)."""
    url = (f"{BASE}/voReport!viewPersonAsign.do?roleId={role_id}&deptId={dept_id}"
           f"&deptName=&roleName={quote(_b64(role_name))}")
    r = s.post(url, data={"dojo.preventCache": now_ms()}, timeout=30)
    try:
        data = r.json()
    except Exception:
        raise PipelineError(f"Không đọc được danh sách người cho vai trò roleId={role_id}.")
    return data.get("items") or []

def resolve_flow_signers(s, flow_id, doc_abstract, log):
    """Với 1 luồng CHƯA có sẵn người (luồng ngoài 3 luồng quen): lấy từng bước, bước nào thiếu
    người thì tự tra (viewNode + viewPersonAsign) — tra RIÊNG từng biến thể chức danh (vd "Cục
    Trưởng" và "Phó Cục Trưởng" ở cùng 1 bước "Lãnh đạo Cục ký"), không gộp/bỏ bớt biến thể nào,
    vì mỗi biến thể có thể ra danh sách người khác nhau (đã xác nhận qua thực tế dùng: "Cục
    Trưởng" chỉ 1 người trong khi "Phó Cục Trưởng" có 3 người — bỏ sót sẽ mất lựa chọn của
    người dùng). Trả về list bước, mỗi bước thêm khoá 'candidates' (list người, mỗi người kèm
    'roleId'/'roleName' của biến thể mà họ giữ) để GUI tự điền (tổng cộng đúng 1 người) hoặc
    hiện ô chọn (từ 2 người trở lên, có ghi kèm chức danh cho dễ phân biệt)."""
    items = fetch_flow_nodes(s, flow_id, doc_abstract, log)
    out = []
    for n in items:
        node = dict(n)
        if node.get("userId") is not None:
            node["candidates"] = []
            out.append(node)
            continue
        log(f"   • Tra người cho bước: {node.get('name')!r} …")
        variants = view_node_roles(s, flow_id, node["nodeId"])
        candidates = []
        for v in variants:
            for p in view_person_candidates(s, v["roleId"], v["deptId"], v["roleName"]):
                p = dict(p)
                p["roleId"], p["roleName"] = v["roleId"], v["roleName"]
                candidates.append(p)
        if variants:
            node["deptId"], node["deptName"] = variants[0]["deptId"], variants[0]["deptName"]
        node["candidates"] = candidates
        if len(candidates) == 1:
            node["userId"] = candidates[0].get("userId")
            node["fullName"] = candidates[0].get("fullName")
            node["roleId"] = candidates[0].get("roleId")
            node["roleName"] = candidates[0].get("roleName")
        out.append(node)
    return out

ID_CACHE_FILE = os.path.join(HERE, "id_cache.json")
def load_id_cache():
    try:
        with open(ID_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}
def save_id_cache(c):
    try:
        atomic_write_json(ID_CACHE_FILE, c, ensure_ascii=False)
    except Exception:
        pass
ID_CACHE = load_id_cache()

def _match_name(target, pairs):
    """pairs = [(id, name)]. Khớp theo tên chuẩn hóa (bỏ dấu + gọn khoảng trắng)."""
    t = " ".join(_norm(target).split())
    for nid, nm in pairs:
        if " ".join(_norm(nm).split()) == t:
            return nid
    return None

def _get_children(s, parent_id, log):
    r = s.post(BASE + "/departmentAction!getChildrenNode.do",
               data={"parentItemId": parent_id, "checboxStatus": "false",
                     "dojo.preventCache": now_ms()}, timeout=30)
    try:
        arr = r.json()
    except Exception:
        arr = json.loads(r.text)
    if isinstance(arr, dict):
        arr = arr.get("items", [])
    return [(str(n.get("id")), n.get("name", "")) for n in arr if n.get("id") is not None]

def _lienthong_roots(s, log):
    r = s.post(BASE + "/departmentAction!getTreeLinkDocument.do",
               data={"dojo.preventCache": now_ms()}, timeout=30)
    try:
        d = r.json()
    except Exception:
        d = json.loads(r.text)
    items = d.get("items", []) if isinstance(d, dict) else d
    return [(str(n.get("id")), n.get("name", "")) for n in items if n.get("id") is not None]

def resolve_node(s, node, tree, log):
    """node={'name','path'}; tree='internal'|'lien_thong'. Trả về ID (đi cây, có cache)."""
    path = node["path"]
    key = tree + "|" + " > ".join(path)
    if key in ID_CACHE:
        return ID_CACHE[key]
    if tree == "internal":
        cur = str(CAY["internal"]["root_id"])     # 52 = Bộ Y Tế
        rest = path[1:]                            # bỏ 'Bộ Y Tế'
    else:
        roots = _lienthong_roots(s, log)
        cur = _match_name(path[0], roots)
        if not cur:
            raise PipelineError(f"Không thấy gốc liên thông '{path[0]}' trên hệ thống.")
        rest = path[1:]
    for seg in rest:
        kids = _get_children(s, cur, log)
        nid = _match_name(seg, kids)
        if not nid:
            raise PipelineError(f"Kẹt ở cấp '{seg.strip()}' (đơn vị: {node['name']}). "
                                f"Tên có thể lệch so với cây thật — gửi log này để chỉnh.")
        cur = nid
    ID_CACHE[key] = cur
    save_id_cache(ID_CACHE)
    return cur

def resolve_nodes(s, nodes, tree, log):
    """Danh sách node -> (chuỗi_tên, chuỗi_id)."""
    names, ids = [], []
    for nd in nodes:
        log(f"   • Giải ID: {nd['name']}…")
        i = resolve_node(s, nd, tree, log)
        names.append(nd["name"].strip()); ids.append(i)
    return ";".join(names), ";".join(ids)

def save_document(s, cfg, doc, draft_attach, draft_sign, log):
    """Lưu 1 văn bản dự thảo (onInsertDraft). `doc` mang loại VB/số ký hiệu/trích yếu RIÊNG
    của văn bản này (mỗi văn bản trong 1 phiếu trình có thể khác loại/khác số/khác trích yếu —
    chỉ nơi nhận/độ khẩn/độ mật/luồng trình là dùng chung, lấy từ cfg). Trả về publishDocumentId."""
    params = decode_params([list(p) for p in TPL["insertDraft"]["params"]])
    P = "publishDocumentCreateForm."
    dt = ENUMS["documentType"][doc["doc_type"]]
    set_param(params, P + "documentTypeId", dt)
    set_param(params, P + "documentType", doc["doc_type"])
    set_param(params, P + "code", doc["code"])
    set_param(params, P + "documentAbstract", doc["abstract"])
    if cfg.get("priority"):
        set_param(params, P + "priorityId", ENUMS["priority"][cfg["priority"]])
        set_param(params, P + "priority", cfg["priority"])
    if cfg.get("security"):
        set_param(params, P + "securityTypeId", ENUMS["security"][cfg["security"]])
        set_param(params, P + "securityType", cfg["security"])
    set_param(params, P + "attachDraftId", draft_attach)   # dự thảo + tài liệu gửi kèm (nối ';')
    set_param(params, P + "signRequere", draft_sign)       # chỉ dự thảo cần ký
    set_param(params, P + "publishDocumentId", "")
    set_param(params, P + "status", "")
    set_param(params, P + "createDatePublish", datetime.now().strftime("%Y-%m-%d"))
    # Nơi nhận: xóa hết rồi đặt lại theo lựa chọn
    for suf in ["receiveInside", "receiveInsideId", "receiveEdoc", "receiveEdocId",
                "receiveToKnow", "receiveToKnowId", "receiveReport", "receiveReportId",
                "receiveOutside", "receiveOutsideId",
                "receiveSaveDepartment", "receiveSaveDepartmentId"]:
        set_suffix(params, suf, "")
    for cat, (nm, idsuf, tree) in {
        "inside":  ("receiveInside",  "receiveInsideId",  "internal"),
        "report":  ("receiveReport",  "receiveReportId",  "internal"),
        "edoc":    ("receiveEdoc",    "receiveEdocId",    "lien_thong"),
        "save":    ("receiveSaveDepartment", "receiveSaveDepartmentId", "internal"),
        "know":    ("receiveToKnow",  "receiveToKnowId",  "internal"),
    }.items():
        nodes = cfg.get("recv_" + cat) or []
        if nodes:
            log(f"• Giải ID nơi nhận ({nm})…")
            n, i = resolve_nodes(s, nodes, tree, log)
            set_suffix(params, nm, n); set_suffix(params, idsuf, i)
    set_param(params, "dojo.preventCache", now_ms())

    tok = reload_token(s, log)
    log("• Lưu văn bản dự thảo (onInsertDraft)…")
    r = s.post(BASE + "/voPublishDocument!onInsertDraft.do",
               params={"struts.token.name": "token", "token": tok},
               data=dict(params), timeout=60)
    http_log(log, r)
    m = re.search(r'"?publishDocumentId"?\s*[:=]\s*"?(\d{4,})', r.text)
    if not m:
        m = re.search(r'(\d{9,})', r.text)
    if not m:
        raise PipelineError("Không lấy được publishDocumentId sau khi lưu.\n--- Phản hồi ---\n" + r.text[:400])
    pid = m.group(1)
    log(f"   → publishDocumentId = {pid}")
    return pid

def flow_items_for_cfg(s, cfg, log):
    """Node list của luồng đang chọn (order/actionType/roleName/userId...), dùng để đánh số chữ
    ký VÀ để dựng myJsonString khi lưu. Ưu tiên `flow_nodes_override` — danh sách đã resolve sẵn
    ở khung "Chọn người ký" (GUI luôn resolve luồng đang chọn, kể cả luồng đã có sẵn người, xem
    FlowSignerPanel) — chỉ tự fetch lại nếu cfg được dựng tay, không qua khung đó."""
    items = cfg.get("flow_nodes_override")
    if items is not None:
        return items
    flow_id = cfg.get("flow_id")
    if not flow_id:
        return []
    return fetch_flow_nodes(s, flow_id, cfg.get("report_content", ""), log)

def save_report_draft(s, cfg, report_attach, report_sign, documents, log, sign="0"):
    """Lưu phiếu trình (onUpdate). `documents`: list các văn bản đã lưu (mỗi dict có sẵn '_pid'
    = publishDocumentId từ save_document()) — ghi thành nhiều dòng draftDocumentGridForm[0],
    [1], ... (giống hệt cách trang web tự thêm dòng khi bạn bấm "Thêm văn bản" nhiều lần trong
    1 phiếu trình — xác nhận qua file HAR).
    `sign`: "0" = LƯU NHÁP (mặc định, an toàn — không đi vào luồng ký duyệt), "1" = TRÌNH THẬT
    (văn bản bắt đầu đi vào luồng ký duyệt) — xác nhận qua HAR: request giống hệt lưu nháp,
    chỉ khác đúng tham số này, không có trường nào khác cần thêm."""
    q = dict(TPL["saveReport"]["query"])
    q = {unquote(k): unquote(v) for k, v in q.items()}

    flow_id = cfg.get("flow_id") or q.get("flowId")
    if flow_id:
        items = flow_items_for_cfg(s, cfg, log)
        q["flowId"] = flow_id
        q["myJsonString"] = build_flow_json(items)

    # Hồ sơ công việc (fileId/profileFlowAsignId) — KHÔNG được để nguyên giá trị tĩnh trong mẫu
    # (đó là hồ sơ của riêng 1 tài khoản cụ thể). Đổi tài khoản mà vẫn dùng hồ sơ cũ, server vẫn
    # báo "thành công" nhưng phiếu trình bị gắn nhầm hồ sơ của người khác — tìm không thấy trong
    # thùng nháp của tài khoản đang dùng. Nếu cfg có chọn hồ sơ (từ danh sách lấy động theo đúng
    # tài khoản), dùng giá trị đó; nếu không mới rơi về giá trị tĩnh trong mẫu (tài khoản gốc).
    work_profile_id = cfg.get("work_profile_id")
    work_profile_name = cfg.get("work_profile_name")
    if work_profile_id:
        q["fileId"] = work_profile_id

    q["sign"] = sign
    q["token"] = reload_token(s, log)

    params = decode_params([list(p) for p in TPL["saveReport"]["params"]])
    if work_profile_id:
        set_param(params, "reportForm.profileFlowAsignId", work_profile_id)
        set_param(params, "reportForm.profileFlowAsign", work_profile_name or "")

    # creator/creatorId — KHÔNG được để nguyên giá trị tĩnh trong mẫu (đó là danh tính riêng 1
    # tài khoản). Server có vẻ lọc "thùng nháp phiếu trình" theo đúng creatorId khớp phiên đăng
    # nhập (xem onSearchMyReport.do) — nếu để nguyên, phiếu trình vẫn "lưu thành công" nhưng bị
    # gắn nhầm danh tính người khác nên không hiện ra. Lấy động theo tài khoản đang chạy; nếu
    # không lấy được thì rơi về giá trị tĩnh trong mẫu (an toàn cho tài khoản gốc).
    # (officeId/officeName TẠM giữ nguyên = Cục Quản lý Dược — hiện chỉ dùng nội bộ Cục này;
    # đơn vị khác dùng thì tự đổi cứng 2 giá trị đó trong mẫu, chưa cần tự động ở bước này.)
    identity = fetch_current_user_identity(s, log)
    if identity:
        user_id, full_name = identity
        set_param(params, "reportForm.creatorId", user_id)
        set_param(params, "reportForm.creator", full_name)
        log(f"   → Người trình (creator): {full_name} (userId={user_id})")

    set_param(params, "reportForm.content", cfg["report_content"])
    set_param(params, "reportForm.attachId", report_attach)   # phiếu trình + tài liệu không gửi (nối ';')
    set_param(params, "reportForm.signRequere", report_sign)  # chỉ phiếu trình cần ký
    set_param(params, "reportForm.reportId", "")
    set_param(params, "reportForm.editorDate", datetime.now().strftime("%Y-%m-%d"))
    now_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    for i, doc in enumerate(documents):
        pfx = f"draftDocumentGridForm[{i}]."
        set_param(params, pfx + "publishDocumentId", doc["_pid"])
        set_param(params, pfx + "code", doc["code"])
        set_param(params, pfx + "documentType", doc["doc_type"])
        set_param(params, pfx + "documentAbstract", doc["abstract"])
        set_param(params, pfx + "createDate", now_iso)
    set_param(params, "dojo.preventCache", now_ms())

    action_label = "TRÌNH" if sign == "1" else "lưu NHÁP"
    log(f"• {action_label.capitalize()} phiếu trình ({len(documents)} văn bản, onUpdate, sign={sign})…")
    r = s.post(BASE + "/voReport!onUpdate.do", params=q, data=dict(params), timeout=60)
    http_log(log, r)
    # QUAN TRỌNG: server này có thể trả HTTP 200 ngay cả khi từ chối yêu cầu (báo lỗi trong nội
    # dung, không đổi mã HTTP) — status_code=200 KHÔNG chắc chắn là đã lưu được. Log nguyên văn
    # phản hồi để biết chắc, chứ không suy đoán.
    body = r.text.strip()
    log(f"   ↳ Nội dung phản hồi ({len(body)} ký tự): {body[:500]!r}")
    if r.status_code != 200:
        raise PipelineError(f"onUpdate trả mã {r.status_code}.\n--- Phản hồi ---\n{body[:400]}")
    log(f"   → Đã gửi yêu cầu {action_label} (xem dòng 'Nội dung phản hồi' ở trên — "
        "nếu trống hoàn toàn thì có thể là bình thường, nếu có chữ thì đọc kỹ xem có báo lỗi không).")

# ---------- Xác minh SAU KHI lưu (đọc lại từ server, không suy đoán từ phản hồi onUpdate) ----------
# Đã xác nhận qua 2 file HAR thật (trình lại 1 phiếu cũ, và tạo phiếu hoàn toàn mới): phản hồi
# của chính onUpdate.do KHÔNG mang lại reportId hay bất kỳ thông tin đọc được nào (cả 2 lần đều
# rỗng dù server báo HTTP 200) — nên không thể suy ra "đã lưu" chỉ từ phản hồi đó. Thay vào đó,
# gọi lại 1-2 API TRA CỨU (đọc, không ghi gì) để tìm đúng phiếu trình vừa lưu bằng cách khớp
# content + creatorId + thời điểm tạo — nếu tìm thấy, đó là bằng chứng THẬT (server đã ghi nhận),
# không phải suy đoán từ mã HTTP.
def _search_my_report(s, grid=None):
    """Gọi onSearchMyReport.do — `grid="prepareProcessDocument"` = đúng hộp "đang trình/đang xử
    lý" (đã xác nhận qua HAR: trả về gọn, vài dòng); `grid=None` = danh sách chung KHÔNG lọc
    trạng thái (bao gồm cả nháp lẫn đã trình — đây là hộp dùng để tìm thấy cả văn bản còn ở
    thùng nháp). Tham số postData khác nhau giữa 2 biến thể là ĐÚNG theo HAR thật, không phải
    thiếu sót — mỗi biến thể có bộ tham số riêng của đúng màn hình tương ứng trên web."""
    now = datetime.now()
    data = {
        "searchForm.content": "", "reportSearchForm.createDateFrom": now.replace(day=1).strftime("%Y-%m-%d"),
        "reportSearchForm.createDateTo": now.strftime("%Y-%m-%d"), "reportSearchForm.content": "",
        "q": "*", "start": 0, "count": 50, "startval": 0,
    }
    params = None
    if grid:
        params = {"grid": grid}
        data["reportSearchForm.reportOfficeNumber"] = ""
    else:
        data["reportSearchForm.reportType"] = "-1"
        data["reportSearchForm.stateId"] = "-1"
        data["reportSearchForm.status"] = "-1"
    r = s.post(BASE + "/voReport!onSearchMyReport.do", params=params, data=data, timeout=30)
    return r.json().get("items") or []

def _match_report(items, content, creator_id, since_dt):
    """Khớp đúng phiếu trình VỪA lưu trong `items` (kết quả _search_my_report): content nguyên
    văn + tạo sau `since_dt` (trừ hao 10 giây cho lệch giờ máy/server) — đã kiểm chứng bằng HAR
    thật (createdDate khớp chính xác tới từng giây với lúc gọi onUpdate). Chỉ so thêm creatorId
    khi có (`creator_id` có thể None nếu không lấy được danh tính tài khoản — vẫn cứ khớp theo
    content + thời gian, không bỏ cuộc hoàn toàn chỉ vì thiếu 1 lớp so khớp phụ)."""
    target = (content or "").strip()
    matches = []
    for it in items:
        if (it.get("content") or "").strip() != target:
            continue
        if creator_id is not None and str(it.get("creatorId")) != str(creator_id):
            continue
        try:
            created = datetime.strptime(it["createdDate"], "%Y-%m-%dT%H:%M:%S")
        except Exception:
            continue
        if created < since_dt - timedelta(seconds=10):
            continue   # phiếu trình cũ trùng nội dung, tạo từ trước lúc mình lưu — bỏ qua
        matches.append((created, it))
    if not matches:
        return None
    matches.sort(key=lambda x: x[0], reverse=True)
    return matches[0][1]

def find_current_pending_step(s, report_id, log=lambda *a: None):
    """Bước đang chờ xử lý (status=1) trong luồng của đúng phiếu trình này — trả về dict có
    receiveUser/displayPositionName, hoặc None nếu không tra được (không phải lỗi nghiêm trọng,
    chỉ là chưa hiện được câu 'đang chờ ai')."""
    try:
        r = s.post(BASE + "/voReport!onSearchReportProcess.do", params={"reportId": report_id},
                   data={"q": "*", "start": 0, "count": 12, "startval": 0}, timeout=30)
        items = r.json().get("items") or []
    except Exception as e:
        log(f"   • Không tra được người đang giữ phiếu trình: {e!r}")
        return None
    return next((it for it in items if it.get("status") == 1), None)

def find_latest_history_note(s, report_id, log=lambda *a: None):
    """Dòng mới nhất trong nhật ký hành động của phiếu trình (Trình ký/Hủy trình ký/...) — trả
    về dict có note/createAt/fullname, hoặc None nếu không tra được."""
    try:
        r = s.post(BASE + "/voReport!getReportHistory.do", params={"objectId": report_id},
                   data={"q": "*", "start": 0, "count": 20, "startval": 0}, timeout=30)
        items = r.json().get("items") or []
    except Exception as e:
        log(f"   • Không tra được nhật ký phiếu trình: {e!r}")
        return None
    return items[0] if items else None

VERIFY_RETRY_DELAYS = [3, 4, 5, 6]   # giây nghỉ TRƯỚC mỗi lần thử — hệ thống này khá chậm, nên
                                     # thử vài lần cách quãng thay vì tin ngay 1 lần đầu; DỪNG
                                     # NGAY khi có kết quả, không chờ hết các lần còn lại (tổng
                                     # cộng tối đa ~18 giây nếu chưa lần nào thấy).

def verify_report_saved(s, cfg, sign, since_dt, log=lambda *a: None):
    """Xác minh (đọc lại từ server) rằng phiếu trình vừa lưu/trình thật sự tồn tại — KHÔNG bao
    giờ raise (best-effort): thất bại ở đây không có nghĩa là việc lưu thất bại, chỉ là chưa
    xác minh được, để nơi gọi hiển thị đúng trạng thái 'chưa chắc' thay vì báo lỗi giả.
    Mỗi lần thử tra ĐỒNG THỜI cả 2 nơi — hộp "đang trình" (grid=prepareProcessDocument) VÀ danh
    sách chung (bao gồm cả thùng nháp) — thấy ở đâu trước thì biết trạng thái luôn: thấy ở hộp
    "đang trình" là bằng chứng mạnh nhất (đã thật sự vào luồng ký duyệt); chỉ thấy ở danh sách
    chung (không thấy ở "đang trình") nghĩa là đã lưu nhưng có thể CHƯA thực sự vào luồng — đáng
    chú ý riêng nếu đang Trình (sign=1), vì lẽ ra phải thấy ở cả 2 nơi.
    Trả về dict: {"verified", "report_id", "in_process", "pending", "history_note"}."""
    result = {"verified": False, "report_id": None, "in_process": False,
              "pending": None, "history_note": None}
    identity = None
    try:
        identity = fetch_current_user_identity(s, log)
    except Exception as e:
        log(f"   • Không lấy được danh tính tài khoản: {e!r} — vẫn thử xác minh, chỉ bớt 1 lớp so khớp.")
    if not identity:
        log("   • Không xác định được tài khoản đang đăng nhập — vẫn thử khớp theo nội dung + thời gian.")
    creator_id = identity[0] if identity else None
    content = cfg.get("report_content", "")

    item, in_process = None, False
    for attempt, delay in enumerate(VERIFY_RETRY_DELAYS, start=1):
        log(f"   • Chờ {delay}s rồi tra lại lần {attempt}/{len(VERIFY_RETRY_DELAYS)}…")
        time.sleep(delay)
        try:
            item_process = _match_report(_search_my_report(s, grid="prepareProcessDocument"),
                                          content, creator_id, since_dt)
        except Exception as e:
            log(f"   • Tra hộp 'đang trình' lỗi: {e!r}")
            item_process = None
        try:
            item_all = _match_report(_search_my_report(s, grid=None), content, creator_id, since_dt)
        except Exception as e:
            log(f"   • Tra danh sách phiếu trình (kể cả nháp) lỗi: {e!r}")
            item_all = None
        if item_process or item_all:
            item, in_process = (item_process or item_all), bool(item_process)
            where = "hộp 'đang trình'" if in_process else "danh sách phiếu trình (chưa thấy ở 'đang trình')"
            log(f"   ✓ Tìm thấy ở {where} — dừng thử lại.")
            break
        log("   • Chưa thấy ở đâu cả, thử lại.")
    else:
        log("   ⚠ Đã thử đủ số lần vẫn chưa thấy — không coi là lỗi lưu, chỉ là chưa xác minh được.")

    if not item:
        return result
    result["verified"] = True
    result["report_id"] = item.get("reportId")
    result["in_process"] = in_process
    if sign == "1":
        try:
            result["pending"] = find_current_pending_step(s, result["report_id"], log)
        except Exception as e:
            log(f"   • Không tra được người đang giữ: {e!r}")
        try:
            result["history_note"] = find_latest_history_note(s, result["report_id"], log)
        except Exception as e:
            log(f"   • Không tra được nhật ký: {e!r}")
    return result

# ---------- Luồng chính ----------
# Các "phase" (mốc lớn) để GUI hiện checklist tiến trình thân thiện, gộp nhiều bước kỹ thuật
# nhỏ (xem step() ở dưới) lại thành các mốc dễ hiểu cho người dùng phổ thông.
PIPELINE_PHASES = [
    ("prepare", "Chuẩn bị file"),
    ("upload", "Tải văn bản lên hệ thống"),
    ("save_docs", "Lưu văn bản"),
    ("save_report", "Lưu phiếu trình"),
    ("verify", "Xác minh với hệ thống"),
]

def run_pipeline(s, cfg, log, check_only=False, phase_cb=lambda key: None):
    # Mỗi văn bản dự thảo trong cfg["documents"] có file + loại VB/số/trích yếu RIÊNG (xác nhận
    # qua HAR "luồng trình 2 văn bản một lúc"): mỗi văn bản upload riêng, onInsertDraft riêng
    # (ra publishDocumentId riêng), rồi phiếu trình liệt kê tất cả qua draftDocumentGridForm[i].
    all_documents = [dict(d) for d in (cfg.get("documents") or [])]
    documents = [d for d in all_documents if d.get("file_draft_main")]
    report_main = cfg.get("file_report_main") or None
    report_files = ([report_main] if report_main else []) + list(cfg.get("files_report_extra") or [])
    borrowed_report_as_draft = False
    if not documents:            # thử nghiệm: chưa chọn văn bản nào thì dùng tạm file phiếu trình
        documents = [{"doc_type": "", "code": "", "abstract": "",
                      "file_draft_main": None, "files_draft_extra": []}]
        borrowed_report_as_draft = True

    n_docs = len(documents)
    total = (5 + n_docs) if check_only else (6 + 2 * n_docs)   # +1 so với trước = bước "Xác minh"
    log(f"Chế độ: {'CHỈ KIỂM TRA (không ghi)' if check_only else 'LƯU NHÁP thật'} · {total} bước"
        + (f" · {n_docs} văn bản" if not borrowed_report_as_draft else ""))
    n = [0]
    def nn():
        n[0] += 1
        return n[0]

    phase_cb("prepare")
    with step(log, nn(), total, "Mở 2 form (phiếu trình + văn bản)"):
        html = open_forms(s, log)

    with step(log, nn(), total, "Tìm ô kẹp file trong form"):
        urls = extract_upload_urls(html)
        log(f"   Tìm thấy {len(urls)} ô: {', '.join(urls) or '(KHÔNG có!)'}")
        if not urls:
            raise PipelineError("Không thấy URL upload trong form.")
        url_report = urls.get("uploadReportFile") or list(urls.values())[0]
        url_draft  = urls.get("uploadDraftFile")  or list(urls.values())[-1]

    for doc in documents:
        main = doc.get("file_draft_main")
        doc["_files"] = (report_files[:1] if borrowed_report_as_draft else
                          ([main] if main else []) + list(doc.get("files_draft_extra") or []))

    with step(log, nn(), total, "Đánh số chữ ký lên Phiếu trình + các Văn bản (theo luồng trình)"):
        if not cfg.get("auto_stamp", True):
            log("   (Bỏ qua — đã tắt tự đánh số.)")
        else:
            flow_items = flow_items_for_cfg(s, cfg, log)
            if borrowed_report_as_draft:
                if report_main and report_main.lower().endswith(".pdf"):
                    log("  › File phiếu trình:")
                    stamped = stamp_signature_numbers(report_main, flow_items, log,
                                                       stamps=cfg.get("stamps_report_override"))
                    report_files = [stamped] + list(cfg.get("files_report_extra") or [])
                log("   (Nhóm VĂN BẢN đang dùng tạm file phiếu trình để thử — không đánh số riêng.)")
            else:
                if report_main and report_main.lower().endswith(".pdf"):
                    log("  › File phiếu trình:")
                    stamped = stamp_signature_numbers(report_main, flow_items, log,
                                                       stamps=cfg.get("stamps_report_override"))
                    report_files = [stamped] + list(cfg.get("files_report_extra") or [])
                for i, doc in enumerate(documents):
                    main = doc.get("file_draft_main")
                    if main and main.lower().endswith(".pdf"):
                        log(f"  › Văn bản {i+1}/{n_docs}: {os.path.basename(main)}")
                        stamped = stamp_signature_numbers(main, flow_items, log,
                                                           stamps=doc.get("stamps_override"))
                        doc["_files"] = [stamped] + list(doc.get("files_draft_extra") or [])

    phase_cb("upload")
    with step(log, nn(), total, f"Upload nhóm PHIẾU TRÌNH ({len(report_files)} file) → lấy ID"):
        report_attach, report_sign = upload_many(s, url_report, report_files, log)
        log(f"   attachId={report_attach}  (ký: {report_sign})")

    for i, doc in enumerate(documents):
        label = f"Upload Văn bản {i+1}/{n_docs} ({len(doc['_files'])} file) → lấy ID"
        with step(log, nn(), total, label):
            attach, sign = upload_many(s, url_draft, doc["_files"], log)
            doc["_attach"], doc["_sign"] = attach, sign
            log(f"   attachDraftId={attach}  (ký: {sign})")

    if check_only:
        with step(log, nn(), total, "Xin token (kiểm tra) — KHÔNG ghi gì"):
            reload_token(s, log)
        log("\n✔ KIỂM TRA OK: phiên, upload, token đều hoạt động. Chưa tạo gì cả.")
        return

    phase_cb("save_docs")
    for i, doc in enumerate(documents):
        with step(log, nn(), total, f"Xin token + LƯU văn bản {i+1}/{n_docs} (onInsertDraft)"):
            doc["_pid"] = save_document(s, cfg, doc, doc["_attach"], doc["_sign"], log)

    submit_sign = cfg.get("submit_sign", "0")
    step_label = "Xin token + TRÌNH phiếu trình (onUpdate sign=1)" if submit_sign == "1" \
        else "Xin token + LƯU NHÁP phiếu trình (onUpdate sign=0)"
    phase_cb("save_report")
    since_dt = datetime.now()   # mốc thời gian TRƯỚC lúc gọi onUpdate — dùng để lọc khi dò lại
                                # reportId ở bước xác minh (loại bỏ phiếu trình cũ trùng nội dung)
    with step(log, nn(), total, step_label):
        save_report_draft(s, cfg, report_attach, report_sign, documents, log, sign=submit_sign)

    phase_cb("verify")
    with step(log, nn(), total, "Xác minh lại với hệ thống (đọc lại, không ghi gì)"):
        verify_result = verify_report_saved(s, cfg, submit_sign, since_dt, log)
        if verify_result["verified"]:
            log(f"   ✓ Đã xác minh: reportId={verify_result['report_id']} có tồn tại trên hệ thống.")
        else:
            log("   ⚠ Chưa xác minh được — không có nghĩa là chưa lưu, chỉ là chưa đọc lại "
                "được ngay, tự kiểm tra thêm trên web nếu cần.")

    if submit_sign == "1":
        log(f"\n✔ XONG TẤT CẢ {total} BƯỚC. Phiếu trình đã được TRÌNH.")
    else:
        log(f"\n✔ XONG TẤT CẢ {total} BƯỚC. Phiếu trình đã lưu NHÁP.")
        log("→ Mở web → thùng nháp phiếu trình → kiểm tra → tự bấm TRÌNH.")
    return {"submit_sign": submit_sign, **verify_result}

# ---------- Giao diện ----------
# ================= LƯU CÀI ĐẶT & DỮ LIỆU THÔNG MINH =================
SETTINGS_FILE = os.path.join(HERE, "settings.json")
STORE_FILE = os.path.join(HERE, "nguoi_dung.json")
KR_SERVICE = "emoh_trinh_vb"
try:
    import keyring
except Exception:
    keyring = None

def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_settings(d):
    try:
        atomic_write_json(SETTINGS_FILE, d, ensure_ascii=False, indent=1)
    except Exception:
        pass

def save_password(username, pw):
    if keyring:
        try:
            keyring.set_password(KR_SERVICE, username, pw); return True
        except Exception:
            return False
    return False

def load_password(username):
    if keyring and username:
        try:
            return keyring.get_password(KR_SERVICE, username)
        except Exception:
            return None
    return None

def load_store():
    try:
        with open(STORE_FILE, encoding="utf-8") as f:
            s = json.load(f)
    except Exception:
        s = {}
    s.setdefault("freq", {}); s.setdefault("pinned", []); s.setdefault("templates", {})
    return s

def save_store(store, path=None):
    """`path`: chỉ để test/gọi tay trỏ sang file khác — mặc định (None) luôn là sổ thật."""
    try:
        atomic_write_json(path or STORE_FILE, store, ensure_ascii=False, indent=1)
    except Exception:
        pass

def bump_freq(store, name, path=None):
    store["freq"][name] = store["freq"].get(name, 0) + 1
    save_store(store, path)

def toggle_pin(store, name, path=None):
    p = store["pinned"]
    p.remove(name) if name in p else p.append(name)
    save_store(store, path)

# ---------- Tìm kiếm không dấu + viết tắt + mờ ----------
# @lru_cache: tên đơn vị/vai trò lặp lại hàng nghìn lần trong cây (mỗi phím gõ quét lại
# toàn bộ pool) nhưng bản thân chuỗi tên không đổi giữa các lần gõ -> nhớ kết quả để
# khỏi normalize/regex lại từ đầu mỗi phím (đỡ giật khi gõ tiếng Việt qua bộ gõ Telex,
# vì 1 ký tự có dấu thường phát sinh vài sự kiện KeyRelease liên tiếp).
@lru_cache(maxsize=None)
def _norm(s):
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d")

@lru_cache(maxsize=None)
def _initials(name):
    return "".join(w[0] for w in re.split(r"[\s\-_/.]+", _norm(name)) if w)

def _is_subseq(q, t):
    it = iter(t)
    return all(c in it for c in q)

def _score(q, name):
    if not q:
        return 0
    n = _norm(name); ini = _initials(name)
    if n.startswith(q): return 100
    if ini.startswith(q): return 95
    if q in n: return 80
    if _is_subseq(q, ini): return 70
    if _is_subseq(q, n): return 40
    return -1

def _node_key(node):
    return " > ".join(node["path"])

def search_nodes(query, nodes, store, k=15):
    """Tìm trên danh sách node {name,parent,path}. Trả về list node."""
    q = _norm(query.strip())
    freq = store.get("freq", {}); pinned = set(store.get("pinned", []))
    if not q:
        cand = [nd for nd in nodes if _node_key(nd) in pinned or freq.get(_node_key(nd), 0) > 0]
        cand.sort(key=lambda nd: (-(1000 if _node_key(nd) in pinned else 0) - freq.get(_node_key(nd), 0), nd["name"]))
        return cand[:k]
    scored = []
    for nd in nodes:
        sc = _score(q, nd["name"])
        if sc < 0:
            continue
        kkey = _node_key(nd)
        boost = (1000 if kkey in pinned else 0) + min(freq.get(kkey, 0), 100)
        scored.append((sc, boost, nd))
    scored.sort(key=lambda x: (-x[0], -x[1], len(x[2]["name"])))
    return [nd for _, _, nd in scored[:k]]

# ---------- Đọc PDF Dự thảo văn bản: tự điền Loại VB + Số/ký hiệu + Trích yếu ----------
# Mẫu Công văn: "Số: <số>/<ký hiệu>" ở đầu văn bản; <số> có thể trống (chưa cấp số).
# Một số mẫu có tên loại (Giấy chứng nhận...) lại ghi "Số hiệu:" thay vì "Số:".
CODE_RE = re.compile(r'S[ốôo]\s*(?:hi[eệ]u)?\s*:?\s*(\d*)\s*/\s*([A-Za-zÀ-ỹĐđ0-9\-]+)')
# Trích yếu (mẫu Công văn): đoạn ngay sau "V/v" (hoặc "Về việc") tới khi gặp dòng trống
# hoặc khối tiêu đề bên cột phải (Kính gửi / Cộng hòa.../ Độc lập.../ Hà Nội, ngày...).
ABSTRACT_RE = re.compile(r'(?:V\s*/\s*v|Về\s+việc)\s*:?\s*(.*)', re.IGNORECASE)
ABSTRACT_STOP = ("kinh gui", "cong hoa", "doc lap", "ha noi, ngay", "ha noi,ngay")

# Trích yếu (mẫu có tên loại — Quyết định/Kế hoạch/Giấy mời...): không có "V/v", trích yếu
# nằm ngay dưới dòng tên loại (VD "QUYẾT ĐỊNH") tới khi gặp dòng mở đầu phần nội dung/căn cứ.
TITLE_ABSTRACT_STOP = ("can cu", "thuc hien", "nham", "xet de nghi", "theo de nghi", "xet ")

def extract_code_from_text(text):
    m = CODE_RE.search(text)
    if not m:
        return None
    num, code = m.group(1), m.group(2)
    return (num + "/" + code) if num else ("/" + code)

def extract_abstract_from_text(text):
    m = ABSTRACT_RE.search(text)
    if not m:
        return None
    tail = text[m.start(1):]
    out = []
    for ln in tail.splitlines()[:6]:      # tối đa 6 dòng cho an toàn
        s = ln.strip()
        if not s or any(_norm(s).startswith(mk) for mk in ABSTRACT_STOP):
            break
        out.append(s)
    joined = " ".join(" ".join(out).split())
    return joined or None

def _sentence_case(s):
    """VIẾT HOA TOÀN BỘ -> Chỉ hoa chữ đầu (dùng cho khối tiêu đề 2 dòng kiểu Giấy chứng nhận)."""
    s = " ".join(s.split())
    return s[:1].upper() + s[1:].lower() if s else s

_DOC_TYPE_NAMES_BY_LEN = sorted(DOC_TYPE_TITLES.keys(), key=len, reverse=True)

def extract_doc_type_title(lines):
    """Dò 15 dòng đầu trang 1 tìm dòng viết HOA khớp 1 tên loại trong ENUMS[documentType]
    (VD "QUYẾT ĐỊNH" -> "Quyết định"). Trả về (tên loại, chỉ số dòng) hoặc (None, None) nếu
    không thấy — mẫu Công văn vốn không có dòng tên loại nên sẽ luôn rơi vào trường hợp này.

    2 bước, ưu tiên độ chắc chắn: (1) quét hết tìm dòng khớp TUYỆT ĐỐI cả dòng trước (đáng tin
    nhất) — có thì dùng luôn; (2) chỉ khi không dòng nào khớp tuyệt đối mới quét lại tìm dòng
    BẮT ĐẦU BẰNG 1 tên loại rồi còn chữ khác phía sau (VD "GIẤY CHỨNG NHẬN THỰC HÀNH TỐT" —
    tên loại "Giấy chứng nhận" + phụ đề dính cùng dòng). Nếu 1 dòng bắt đầu bằng nhiều tên loại
    (VD tên loại này là tiền tố của tên loại khác), ưu tiên tên DÀI hơn (cụ thể hơn)."""
    upper_lines = []
    for i, raw in enumerate(lines[:15]):
        s = " ".join(raw.split())
        if s and _is_upper_title(s):
            upper_lines.append((i, s))
    for i, s in upper_lines:
        dt = DOC_TYPE_TITLES.get(s)
        if dt:
            return dt, i
    for i, s in upper_lines:
        for name in _DOC_TYPE_NAMES_BY_LEN:
            if s.startswith(name + " "):
                return DOC_TYPE_TITLES[name], i
    return None, None

def extract_abstract_after_title(lines, title_idx):
    """Trích yếu cho văn bản có tên loại. Nếu dòng ngay sau tên loại cũng viết HOA (khối
    tiêu đề 2 dòng, VD "GIẤY CHỨNG NHẬN" + "ĐỦ ĐIỀU KIỆN KINH DOANH DƯỢC") thì ghép 2 dòng
    và chỉ hoa chữ đầu. Ngược lại lấy các dòng thường ngay dưới, dừng khi gặp dòng mở đầu
    phần nội dung (Căn cứ/Thực hiện/Nhằm...) hoặc 1 dòng viết HOA khác (dòng chức danh, VD
    "BỘ TRƯỞNG BỘ Y TẾ")."""
    nxt = lines[title_idx + 1].strip() if title_idx + 1 < len(lines) else ""
    if nxt and _is_upper_title(nxt):
        return _sentence_case(lines[title_idx] + " " + nxt)
    out = []
    for ln in lines[title_idx + 1: title_idx + 1 + 6]:
        s = ln.strip()
        if not s or _is_upper_title(s):
            break
        if any(_norm(s).startswith(mk) for mk in TITLE_ABSTRACT_STOP):
            break
        out.append(s)
    joined = " ".join(" ".join(out).split())
    return joined or None

def extract_draft_fields(path):
    """Đọc trang 1 file PDF dự thảo. Trả về (loại VB, code, abstract) — phần không suy ra
    được = None. loại VB chỉ có giá trị khi nhận ra được dòng tên loại (Quyết định, Kế
    hoạch...); mẫu Công văn không có dòng này nên loại VB luôn là None (giữ nguyên combobox)."""
    if PdfReader is None:
        raise RuntimeError("Chưa cài thư viện 'pypdf'. Chạy: pip install pypdf")
    reader = PdfReader(path)
    if not reader.pages:
        return None, None, None
    text = reader.pages[0].extract_text() or ""
    text = unicodedata.normalize("NFC", text)
    code = extract_code_from_text(text)

    doc_type, title_idx = extract_doc_type_title(text.splitlines())
    if doc_type is not None:
        abstract = extract_abstract_after_title(text.splitlines(), title_idx)
        return doc_type, code, abstract
    return None, code, extract_abstract_from_text(text)

def convert_office_doc_to_pdf(path):
    """Chuyển .docx sang .pdf bằng gói docx2pdf (điều khiển Word cài sẵn trên máy). KHÔNG hỗ
    trợ .doc (định dạng cũ) — gói docx2pdf tự chặn cứng chỉ nhận .docx. Lưu PDF cạnh file gốc,
    cùng tên. Trả về đường dẫn PDF vừa tạo."""
    if docx2pdf is None:
        raise RuntimeError("Chưa cài thư viện 'docx2pdf'. Chạy: pip install docx2pdf")
    if not path.lower().endswith(".docx"):
        raise RuntimeError("Chỉ hỗ trợ .docx — file .doc (định dạng cũ) hãy tự mở bằng Word, "
                            "'Save As' sang .docx hoặc .pdf rồi chọn lại.")
    src = os.path.abspath(path)
    pdf_path = os.path.splitext(src)[0] + ".pdf"
    docx2pdf.convert(src, pdf_path)
    if not os.path.exists(pdf_path):
        raise RuntimeError("Word không tạo ra file PDF (không rõ lý do) — tự xuất PDF tay rồi chọn lại.")
    return pdf_path

def extract_phieu_trinh_content(path):
    """Đọc trang 1 file PHIẾU TRÌNH (không phải văn bản dự thảo): tìm dòng "PHẦN I..." làm
    mốc, lấy nội dung ngay sau CỤM IN ĐẬM ĐẦU TIÊN xuất hiện sau mốc đó (nhãn trường — không
    cần biết trước tên nhãn cụ thể như "Tên văn bản trình:"/"Nội dung xin ý kiến:"/..., tự
    nhận theo định dạng đậm/thường chung của mẫu phiếu trình) cho tới cụm in đậm tiếp theo
    (nhãn trường kế tiếp), gộp tất cả thành 1 dòng. Trả None nếu không tìm được mốc/nhãn (mẫu
    phiếu trình khác, hoặc PDF quét ảnh không có lớp chữ) — không suy đoán, để điền tay.
    Cần pymupdf (fitz) vì cần biết chữ nào IN ĐẬM — pypdf (dùng cho extract_draft_fields) chỉ
    đọc được chữ thô, không giữ định dạng."""
    if fitz is None:
        raise RuntimeError("Chưa cài thư viện 'pymupdf'. Chạy: pip install pymupdf")
    doc = fitz.open(path)
    try:
        if len(doc) == 0:
            return None
        page = doc[0]
        spans = []
        for block in page.get_text("dict").get("blocks", []):
            for line in block.get("lines", []):
                for sp in line.get("spans", []):
                    text = unicodedata.normalize("NFC", sp.get("text", "")).strip()
                    if not text:
                        continue
                    bold = bool(sp.get("flags", 0) & 16) or "bold" in (sp.get("font") or "").lower()
                    bbox = sp["bbox"]
                    spans.append({"y0": bbox[1], "x0": bbox[0], "text": text, "bold": bold})
        spans.sort(key=lambda s: (round(s["y0"]), s["x0"]))

        anchor_idx = next((i for i, s in enumerate(spans)
                            if re.match(r"ph[ầâ]n\s*i\b", s["text"], re.IGNORECASE)), None)
        if anchor_idx is None:
            return None
        label_idx = next((i for i in range(anchor_idx + 1, len(spans)) if spans[i]["bold"]), None)
        if label_idx is None:
            return None
        i = label_idx
        while i < len(spans) and spans[i]["bold"]:   # nhãn có thể tách nhiều mảnh in đậm liền nhau
            i += 1
        parts = []
        while i < len(spans) and not spans[i]["bold"]:
            parts.append(spans[i]["text"])
            i += 1
        content = " ".join(parts).strip()
        return content or None
    finally:
        doc.close()

# ---------- Đọc thẳng .docx (không cần đợi Word chuyển sang PDF) ----------
# Mục tiêu: tự điền NGAY khi vừa chọn file .docx, thay vì phải đợi Word convert xong (có thể
# mất vài giây tới vài chục giây). Việc chuyển sang PDF (convert_office_doc_to_pdf) vẫn chạy
# như cũ — vẫn cần PDF thật để xem trước/đóng dấu/upload — chỉ là bước ĐỌC ĐỂ TỰ ĐIỀN không
# còn phải chờ nó nữa. Dùng lại đúng các hàm tách chữ (extract_code_from_text/
# extract_abstract_from_text/extract_doc_type_title/extract_abstract_after_title) như bản PDF,
# chỉ khác nguồn lấy "lines".

DOCX_PAGE1_MAX_LINES = 40   # .docx không có ranh giới trang thật (khác PDF) — lấy tạm N đoạn
                            # văn đầu tiên làm "trang 1", đủ rộng cho quốc hiệu/số ký hiệu/tên
                            # loại/trích yếu của mọi mẫu văn bản hành chính đã gặp

def _iter_docx_paragraphs(document):
    """Duyệt TOÀN BỘ đoạn văn theo đúng thứ tự xuất hiện trong file, kể cả đoạn nằm trong Ô
    BẢNG — mẫu Công văn chuẩn hành chính thường đặt khối "Số:.../V/v..." trong 1 bảng 2 cột
    (quốc hiệu bên phải, tên cơ quan bên trái), python-docx mặc định tách riêng
    document.paragraphs/document.tables và BỎ QUA chữ trong bảng, nên phải tự duyệt cây XML
    để không bỏ sót đúng đoạn quan trọng nhất."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn

    def _walk(parent_elm, parent):
        for child in parent_elm:
            if child.tag == qn("w:p"):
                yield Paragraph(child, parent)
            elif child.tag == qn("w:tbl"):
                table = Table(child, parent)
                for row in table.rows:
                    for cell in row.cells:
                        yield from _walk(cell._tc, cell)

    yield from _walk(document.element.body, document)

def _run_is_bold(run, paragraph):
    """run.bold trả None nếu đậm/không đậm là do KẾ THỪA từ style (không đặt riêng ở run) —
    phải lần lên style cha của đoạn văn mới biết chắc, nếu không sẽ hiểu nhầm là "không đậm"."""
    if run.bold is not None:
        return run.bold
    style = paragraph.style
    while style is not None:
        b = style.font.bold
        if b is not None:
            return b
        style = style.base_style
    return False

def extract_draft_fields_docx(path):
    """Như extract_draft_fields() nhưng đọc thẳng .docx bằng python-docx — không cần Word."""
    if docx is None:
        raise RuntimeError("Chưa cài thư viện 'python-docx'. Chạy: pip install python-docx")
    d = docx.Document(path)
    lines = [unicodedata.normalize("NFC", p.text) for p in _iter_docx_paragraphs(d) if p.text.strip()]
    lines = lines[:DOCX_PAGE1_MAX_LINES]
    if not lines:
        return None, None, None
    text = "\n".join(lines)
    code = extract_code_from_text(text)
    doc_type, title_idx = extract_doc_type_title(lines)
    if doc_type is not None:
        abstract = extract_abstract_after_title(lines, title_idx)
        return doc_type, code, abstract
    return None, code, extract_abstract_from_text(text)

def extract_phieu_trinh_content_docx(path):
    """Như extract_phieu_trinh_content() nhưng đọc thẳng .docx bằng python-docx — không cần
    Word. In đậm đọc từ run.bold (kèm suy ra từ style cha qua _run_is_bold), đáng tin hơn cách
    đoán qua tên font trong bản PDF. Các run liền kề CÙNG trạng thái đậm/thường trong 1 đoạn
    văn được gộp thành 1 "khối" — Word hay tách run vụn vặt (gõ sửa, gạch chân chính tả...)
    không liên quan gì tới định dạng đậm/thường thật, gộp lại để khỏi vỡ thuật toán dò nhãn."""
    if docx is None:
        raise RuntimeError("Chưa cài thư viện 'python-docx'. Chạy: pip install python-docx")
    d = docx.Document(path)
    spans = []
    for p in _iter_docx_paragraphs(d):
        cur_text, cur_bold = "", None
        for r in p.runs:
            t = unicodedata.normalize("NFC", r.text)
            if not t:
                continue
            b = _run_is_bold(r, p)
            if b == cur_bold:
                cur_text += t
            else:
                if cur_text.strip():
                    spans.append({"text": cur_text.strip(), "bold": cur_bold})
                cur_text, cur_bold = t, b
        if cur_text.strip():
            spans.append({"text": cur_text.strip(), "bold": cur_bold})

    anchor_idx = next((i for i, s in enumerate(spans)
                        if re.match(r"ph[ầâ]n\s*i\b", s["text"], re.IGNORECASE)), None)
    if anchor_idx is None:
        return None
    label_idx = next((i for i in range(anchor_idx + 1, len(spans)) if spans[i]["bold"]), None)
    if label_idx is None:
        return None
    i = label_idx
    while i < len(spans) and spans[i]["bold"]:
        i += 1
    parts = []
    while i < len(spans) and not spans[i]["bold"]:
        parts.append(spans[i]["text"])
        i += 1
    content = " ".join(parts).strip()
    return content or None

def flow_id_for_code(code, flow_store):
    """Suy ra luồng trình theo ký hiệu văn bản, dùng quy tắc "từ khoá -> luồng" của CHÍNH máy
    này (flow_store["rules"], thứ tự ưu tiên = thứ tự trong list) — không còn hardcode flowId
    trong code, vì mỗi máy/mỗi tài khoản có thể có luồng khác nhau. Không khớp quy tắc nào thì
    rơi về flow_store["default_flow_id"] (nếu máy này đã đặt)."""
    n = _norm(code or "")
    for rule in flow_store.get("rules", []):
        kw = rule.get("keyword")
        if kw and _norm(kw) in n:
            return rule.get("flowId")
    return flow_store.get("default_flow_id")

def flow_id_for_doc(doc_type, code, flow_store):
    """Như flow_id_for_code(), nhưng loại VB có luồng cố định (flow_store["doc_type_rules"])
    thì ưu tiên luồng đó, không cần xét ký hiệu."""
    forced = flow_store.get("doc_type_rules", {})
    if doc_type in forced:
        return forced[doc_type]
    return flow_id_for_code(code, flow_store)

# ---------- Đánh số chữ ký lên PDF dự thảo (theo luồng trình) ----------
# Vị trí ký = chú thích PDF (Text annot, icon "Comment", nội dung = số) — không phải chữ in
# lên trang. Số đóng dấu = đúng "order" của bước trong luồng đang chọn (actionType∈{1,4,5} =
# bước "ký" thật, khác bước phê duyệt/cấp số/ban hành) — xác nhận khớp 100% qua HAR trên cả 3
# luồng cá nhân cũ (order 1/2-3/5 khớp đúng số đã hardcode trước đây). Chức danh trên PDF được
# so khớp với `roleName` thật của luồng đang chọn (không còn giới hạn 4 tên cứng) — nên áp dụng
# được cho BẤT KỲ luồng nào, không riêng 3 luồng cũ.
SIG_CHUYEN_VIEN_ROLE = "CHUYÊN VIÊN"   # quy ước cố định, luôn = 0 — KHÔNG phải 1 bước thật của
                                        # luồng nào (không xuất hiện trong searchNodeInFlow.do ở
                                        # cả 3 luồng cá nhân đã kiểm chứng), là người soạn thảo,
                                        # đứng trước cả bước ký đầu tiên của luồng.

def _norm_role_text(s):
    return unicodedata.normalize("NFC", (s or "")).strip().upper()

def build_role_number_map(items):
    """Dựng {chức danh viết HOA: số đóng dấu} từ node list của luồng đang chọn (kết quả
    searchNodeInFlow.do / resolve_flow_signers, đã có order/actionType/roleName) — thay cho
    bảng cứng trước đây. Chỉ lấy bước "ký" thật (actionType 1/4/5); actionType 0 (phê duyệt),
    2/3 (văn thư cấp số/ban hành) không đóng dấu số, không đưa vào bảng."""
    m = {SIG_CHUYEN_VIEN_ROLE: 0}
    for n in items or []:
        if n.get("actionType") in (1, 4, 5):
            rn = n.get("roleName")
            if rn:
                m[_norm_role_text(rn)] = n.get("order")
    return m

def _is_upper_title(text):
    """Dòng chức danh thật luôn viết HOA hết (phân biệt với câu văn thường có nhắc tới
    tên vai trò, ví dụ 'Kính gửi: ..., Phó Cục trưởng ...' — câu này không viết hoa hết)."""
    letters = [c for c in text if c.isalpha()]
    return bool(letters) and all(c == c.upper() for c in letters)

def _sig_page_lines(page):
    """Toạ độ từng dòng chữ trên trang, sắp theo thứ tự trên->dưới."""
    lines = []
    d = page.get_text("dict")
    for block in d.get("blocks", []):
        for ln in block.get("lines", []):
            text = "".join(sp.get("text", "") for sp in ln.get("spans", [])).strip()
            if not text:
                continue
            # Font nhúng trong PDF đôi khi ánh xạ chữ có dấu tiếng Việt ra Unicode không
            # chuẩn NFC (khác dạng chữ trong role_map) — chuẩn hoá về NFC để so khớp chuỗi
            # ("CỤC TRƯỞNG"...) không bị trật, dù nhìn bằng mắt vẫn giống hệt nhau.
            text = unicodedata.normalize("NFC", text)
            x0, y0, x1, y1 = ln["bbox"]
            lines.append({"text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1})
    lines.sort(key=lambda l: (round(l["y0"]), l["x0"]))
    return lines

# Mục lớn đánh số La Mã đứng riêng 1 dòng (vd "I.", "II.", "V.") — CHỮ HOA hết nên qua được
# _is_upper_title(), nhưng KHÔNG PHẢI chức danh. Nếu không loại, dòng 1 chữ cái như "V" khớp
# nhầm qua vế "up in role" vì "V" tình cờ là 1 ký tự con nằm trong "CHUYÊN VIÊN" — gây phát
# hiện "TRÙNG số" giả (2 mục La Mã khác nhau cùng bị gán số 0), chặn cả file dù không có gì sai
# thật (xác nhận qua ca thực tế: mục "I." và "V." trên cùng 1 trang, cả 2 báo trùng số 0).
_ROMAN_NUMERAL_RE = re.compile(r'^[IVXLCDM]+$')
SIG_MIN_TITLE_LEN = 4   # chức danh thật luôn là cụm nhiều chữ (vd "CỤC TRƯỞNG") — dòng ngắn hơn
                        # ngưỡng này không đủ để so khớp kiểu "nằm trong nhau" (tier 1/2), chỉ
                        # còn so khớp NGUYÊN VĂN (tier 3) là an toàn dù ngắn.

def _sig_find_hits(lines, role_map):
    """Các dòng CHỮ HOA khớp 1 chức danh trong role_map (của luồng đang chọn). Trả về
    [(chỉ số dòng, tên chức danh khớp được)]. So khớp CẢ HAI CHIỀU — vì roleName lấy từ server
    (vd "Phó Cục trưởng") có thể dài/ngắn hơn dòng in thật trên PDF (vd chỉ in "CỤC TRƯỞNG",
    không có "PHÓ"): ưu tiên khớp đúng nguyên văn, rồi tới chức danh NẰM TRONG dòng in (dòng in
    dài hơn, có thể có tiền tố "KT."/"PHÓ" phía trước), cuối cùng mới tới dòng in NẰM TRONG chức
    danh (dòng in ngắn/chung chung hơn chức danh thật của luồng). Loại trước các dòng quá ngắn/
    số La Mã đứng riêng — không thể là chức danh thật, chỉ gây khớp nhầm ở vế so khớp lỏng."""
    hits = []
    for i, ln in enumerate(lines):
        if not _is_upper_title(ln["text"]):
            continue
        up = ln["text"].upper()
        core = up.rstrip(".)").strip()
        if _ROMAN_NUMERAL_RE.match(core):
            continue
        best = None   # (mức ưu tiên, độ dài chức danh, tên chức danh)
        for role in role_map:
            if role == up:
                tier = 3
            elif len(up) >= SIG_MIN_TITLE_LEN and role in up:
                tier = 2
            elif len(up) >= SIG_MIN_TITLE_LEN and up in role:
                tier = 1
            else:
                continue
            cand = (tier, len(role), role)
            if best is None or cand > best:
                best = cand
        if best:
            hits.append((i, best[2]))
    return hits

def _sig_cluster_hits(lines, hits):
    """Gộp các dòng liền kề, cùng vai trò, không có khoảng trống (VD 'KT. CỤC TRƯỞNG' +
    'PHÓ CỤC TRƯỞNG') thành 1 vị trí ký — lấy dòng CUỐI cụm (chức danh thật) làm mốc.
    So sánh trực tiếp 2 hit liên tiếp (không đòi hỏi liền chỉ số trong 'lines'), vì cùng
    hàng ngang có thể xen dòng chấm chấm/nội dung cột khác không liên quan (VD khung
    'PHẦN III' bên trái chạy song song với khối chữ ký bên phải)."""
    clusters = []
    i = 0
    while i < len(hits):
        j = i
        while j + 1 < len(hits) and hits[j + 1][1] == hits[j][1]:
            cur, nxt = lines[hits[j][0]], lines[hits[j + 1][0]]
            gap = nxt["y0"] - cur["y1"]
            overlap = min(nxt["x1"], cur["x1"]) - max(nxt["x0"], cur["x0"])
            if gap <= 6 and overlap > 0:   # cùng cột (chồng lấn x) và không có khoảng trống
                j += 1
            else:
                break
        clusters.append((hits[j][0], hits[j][1]))   # (chỉ số dòng mốc, vai trò)
        i = j + 1
    return clusters

def _sig_find_name_line(lines, anchor_idx, max_gap=220):
    """Dòng tên người ký gần nhất phía dưới mốc, cùng cột (chồng lấn theo trục x)."""
    anchor = lines[anchor_idx]
    width = anchor["x1"] - anchor["x0"]
    for k in range(anchor_idx + 1, len(lines)):
        cand = lines[k]
        if cand["y0"] <= anchor["y1"]:
            continue
        if cand["y0"] - anchor["y1"] > max_gap:
            break
        overlap = min(cand["x1"], anchor["x1"]) - max(cand["x0"], anchor["x0"])
        if overlap > 0.3 * width:
            return cand
    return None

def find_signature_stamps(doc, flow_items, log):
    """Quét TOÀN BỘ file (mọi trang, kể cả phụ lục), trả về các vị trí cần đánh số.
    `flow_items`: node list của luồng đang chọn (searchNodeInFlow.do / resolve_flow_signers)."""
    role_map = build_role_number_map(flow_items)
    stamps = []
    for pno in range(len(doc)):
        lines = _sig_page_lines(doc[pno])
        hits = _sig_find_hits(lines, role_map)
        if not hits:
            continue
        for anchor_idx, role in _sig_cluster_hits(lines, hits):
            number = role_map.get(role)
            if number is None:
                log(f"   • Trang {pno+1}: thấy '{role}' nhưng luồng hiện tại không có số "
                    f"cho vai trò này — bỏ qua.")
                continue
            anchor = lines[anchor_idx]
            name_line = _sig_find_name_line(lines, anchor_idx)
            y_mid = ((anchor["y1"] + name_line["y0"]) / 2) if name_line else (anchor["y1"] + 40)
            x_mid = (anchor["x0"] + anchor["x1"]) / 2
            stamps.append({"page": pno, "x": x_mid, "y": y_mid, "number": number, "role": role,
                            "title": anchor["text"]})
            log(f"   • Trang {pno+1}: '{anchor['text']}' → số {number} (x={x_mid:.0f}, y={y_mid:.0f})")
    return stamps

def _sig_tagged_path(path):
    folder = os.path.dirname(path)
    stem, ext = os.path.splitext(os.path.basename(path))
    today = datetime.now().strftime("%d.%m.%y")
    return os.path.join(folder, f"{stem} - {today}{ext}")

def stamp_signature_numbers(path, flow_items, log, stamps=None):
    """Đọc 1 file PDF (phiếu trình hoặc dự thảo văn bản), tìm vị trí ký theo luồng, ghi
    chú thích (Text annot) số thứ tự. Lưu bản đã đánh số NGAY CẠNH file gốc, tên gắn thêm
    " - <ngày hôm nay dd.mm.yy>" (file gốc giữ nguyên, không sửa). Trả về đường dẫn để upload —
    là bản đã đánh số nếu có đánh được gì, hoặc chính path gốc nếu không tìm thấy vị trí ký.
    Nếu `stamps` được truyền vào (VD: đã được người dùng chỉnh trong khung xem trước),
    dùng nguyên danh sách đó thay vì tự quét lại file. `flow_items`: node list của luồng
    đang chọn (xem flow_items_for_cfg)."""
    if fitz is None:
        raise PipelineError("Chưa cài thư viện 'pymupdf'. Chạy: pip install pymupdf")
    if not flow_items:
        raise PipelineError("Chưa chọn Luồng trình — không biết đánh số theo luồng nào.")

    doc = fitz.open(path)
    try:
        if stamps is None:
            stamps = find_signature_stamps(doc, flow_items, log)
        if not stamps:
            log("   ⚠ Không tìm thấy vị trí ký nào trong file — giữ nguyên file gốc, không đánh số.")
            return path

        by_number = {}
        for st in stamps:
            by_number.setdefault(st["number"], []).append(st)
        dups = {n: v for n, v in by_number.items() if len(v) > 1}
        if dups:
            detail = "; ".join(
                f"số {n} lặp {len(v)} lần (trang {', '.join(str(x['page']+1) for x in v)}, "
                f"chức danh: {', '.join(x['title'] for x in v)})"
                for n, v in dups.items())
            raise PipelineError(
                f"Phát hiện TRÙNG số chữ ký trong file (1 file không được trùng số) — "
                f"dừng lại để kiểm tra tay: {detail}")

        expected_max = max((n.get("order") for n in flow_items if n.get("actionType") in (1, 4, 5)),
                            default=None)
        actual_max = max(st["number"] for st in stamps)
        if expected_max is not None and actual_max != expected_max:
            log(f"   ⚠ Cảnh báo: luồng đang chọn thường có số cao nhất = {expected_max}, "
                f"nhưng file này chỉ thấy tới số {actual_max}. Kiểm tra lại file/luồng trước khi trình.")

        for st in stamps:
            page = doc[st["page"]]
            point = fitz.Point(st["x"] - 8, st["y"] - 8)
            annot = page.add_text_annot(point, str(st["number"]), icon="Comment")
            annot.update()

        out_path = _sig_tagged_path(path)
        try:
            doc.save(out_path)
        except Exception as e:
            raise PipelineError(f"Không lưu được file đã đánh số vào '{out_path}': {e}")
        log(f"   → Đã đánh {len(stamps)} số chữ ký, lưu tại: {out_path}")
        return out_path
    finally:
        doc.close()

# ---------- Ô chọn nơi nhận (tìm kiếm + chip + bộ mẫu) ----------
class RecipientBox(ttk.LabelFrame):
    CATS = [("inside", "Nhận nội bộ"), ("report", "Báo cáo"), ("edoc", "Liên thông"),
            ("save", "Nơi lưu"), ("know", "Để biết")]

    def __init__(self, parent, cay, store):
        super().__init__(parent, text="Nơi nhận  (gõ để tìm — không dấu/viết tắt đều được)", padding=8)
        self.store = store
        self.pools = {
            "inside": cay["internal"]["nodes"], "save": cay["internal"]["nodes"],
            "know": cay["internal"]["nodes"], "report": cay["internal"]["nodes"],
            "edoc": cay["lien_thong"]["nodes"],
        }
        self.buckets = {c: [] for c, _ in self.CATS}   # list các node
        self.cat = tk.StringVar(value="inside")

        rowc = ttk.Frame(self); rowc.pack(fill="x")
        ttk.Label(rowc, text="Thêm vào:").pack(side="left")
        for c, label in self.CATS:
            ttk.Radiobutton(rowc, text=label, value=c, variable=self.cat,
                            command=self._refilter).pack(side="left", padx=2)

        self.entry = ttk.Entry(self); self.entry.pack(fill="x", pady=(4, 0))
        self.entry.bind("<KeyRelease>", self._on_key)
        self.entry.bind("<Down>", lambda e: self._focus_list())
        self.entry.bind("<Return>", self._add_top)

        self.lb = tk.Listbox(self, height=8, activestyle="dotbox")
        self.lb.pack(fill="x")
        self.lb.bind("<Return>", self._add_selected)
        self.lb.bind("<Double-Button-1>", self._add_selected)
        self.lb.bind("<Button-3>", self._context)
        self.lb.bind("<Button-2>", self._context)

        self.chips = ttk.Frame(self); self.chips.pack(fill="x", pady=(6, 0))
        rowt = ttk.Frame(self); rowt.pack(fill="x", pady=(6, 0))
        ttk.Button(rowt, text="💾 Lưu bộ mẫu", command=self._save_template).pack(side="left")
        ttk.Button(rowt, text="📂 Nạp bộ mẫu", command=self._load_template).pack(side="left", padx=4)
        ttk.Label(rowt, text="  (chuột phải mục để Ghim ★)", foreground="gray").pack(side="left")

        self._matches = []
        self._debounce_id = None
        self._refilter(); self._render_chips()

    def _pool(self):
        return self.pools[self.cat.get()]

    def _label(self, nd):
        par = nd.get("parent")
        return nd["name"].strip() + (f"   ‹{par.strip()}›" if par else "")

    def _on_key(self, e):
        if e.keysym in ("Up", "Down", "Return"):
            return
        # Gộp các sự kiện KeyRelease liên tiếp (bộ gõ tiếng Việt hay phát sinh nhiều
        # sự kiện cho 1 ký tự có dấu) thành 1 lần lọc duy nhất sau khi ngừng gõ.
        if self._debounce_id is not None:
            self.after_cancel(self._debounce_id)
        self._debounce_id = self.after(120, self._debounced_refilter)

    def _debounced_refilter(self):
        self._debounce_id = None
        try:
            self._refilter()
        except tk.TclError:
            pass

    def _refilter(self, *_):
        self._matches = search_nodes(self.entry.get(), self._pool(), self.store)
        pinned = set(self.store.get("pinned", []))
        self.lb.delete(0, "end")
        for nd in self._matches:
            mark = "★ " if _node_key(nd) in pinned else "    "
            self.lb.insert("end", mark + self._label(nd))

    def _focus_list(self):
        if self._matches:
            self.lb.focus_set(); self.lb.selection_clear(0, "end")
            self.lb.selection_set(0); self.lb.activate(0)

    def _add_top(self, *_):
        if self._matches:
            self._add(self._matches[0])

    def _add_selected(self, *_):
        sel = self.lb.curselection()
        if sel:
            self._add(self._matches[sel[0]])

    def _add(self, node):
        c = self.cat.get()
        if not any(_node_key(x) == _node_key(node) for x in self.buckets[c]):
            self.buckets[c].append(node)
            bump_freq(self.store, _node_key(node))
        self.entry.delete(0, "end")
        self._refilter(); self._render_chips()
        self.entry.focus_set()

    def _remove(self, c, key):
        self.buckets[c] = [x for x in self.buckets[c] if _node_key(x) != key]
        self._render_chips()

    def _render_chips(self):
        for w in self.chips.winfo_children():
            w.destroy()
        any_sel = False
        for c, label in self.CATS:
            if not self.buckets[c]:
                continue
            any_sel = True
            row = ttk.Frame(self.chips); row.pack(fill="x", anchor="w", pady=1)
            ttk.Label(row, text=label + ":", width=12).pack(side="left", anchor="n")
            wrap = ttk.Frame(row); wrap.pack(side="left", fill="x")
            for nd in self.buckets[c]:
                chip = tk.Frame(wrap, bg="#e3f2fd", bd=1, relief="solid")
                chip.pack(side="left", padx=2, pady=1)
                tk.Label(chip, text=nd["name"].strip(), bg="#e3f2fd").pack(side="left", padx=(4, 0))
                tk.Button(chip, text="✕", bd=0, bg="#e3f2fd", padx=2,
                          command=lambda cc=c, k=_node_key(nd): self._remove(cc, k)).pack(side="left")
        if not any_sel:
            ttk.Label(self.chips, text="(chưa chọn nơi nhận nào)", foreground="gray").pack(anchor="w")

    def _context(self, e):
        idx = self.lb.nearest(e.y)
        if idx < 0 or idx >= len(self._matches):
            return
        key = _node_key(self._matches[idx])
        pinned = set(self.store.get("pinned", []))
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=("Bỏ ghim" if key in pinned else "Ghim ★"),
                      command=lambda: (toggle_pin(self.store, key), self._refilter()))
        m.tk_popup(e.x_root, e.y_root)

    def _save_template(self):
        from tkinter import simpledialog
        if not any(self.buckets.values()):
            messagebox.showinfo("Bộ mẫu", "Chưa chọn nơi nhận nào để lưu."); return
        name = simpledialog.askstring("Lưu bộ mẫu", "Đặt tên bộ mẫu:")
        if name:
            self.store["templates"][name] = {c: list(v) for c, v in self.buckets.items()}
            save_store(self.store)
            messagebox.showinfo("Bộ mẫu", f"Đã lưu bộ mẫu '{name}'.")

    def _load_template(self):
        tpls = self.store.get("templates", {})
        if not tpls:
            messagebox.showinfo("Bộ mẫu", "Chưa có bộ mẫu nào."); return
        m = tk.Menu(self, tearoff=0)
        for name in tpls:
            m.add_command(label=name, command=lambda n=name: self._apply_template(n))
        m.tk_popup(self.winfo_pointerx(), self.winfo_pointery())

    def _apply_template(self, name):
        tpl = self.store["templates"].get(name, {})
        for c, _ in self.CATS:
            existing = {_node_key(x) for x in self.buckets[c]}
            for nd in tpl.get(c, []):
                if isinstance(nd, dict) and _node_key(nd) not in existing:
                    self.buckets[c].append(nd)
        self._render_chips()

    def get(self, cat):
        return list(self.buckets[cat])   # danh sách node


class FileList(ttk.Frame):
    """Danh sách file: nút thêm + listbox + nút bỏ. get() -> [đường dẫn]."""
    def __init__(self, parent, label):
        super().__init__(parent)
        self.pack(fill="x", pady=2)
        self.paths = []
        top = ttk.Frame(self); top.pack(fill="x")
        ttk.Label(top, text=label, width=22).pack(side="left")
        ttk.Button(top, text="Thêm file…", command=self._add).pack(side="left")
        ttk.Button(top, text="Bỏ chọn", command=self._remove).pack(side="left", padx=4)
        self.lb = tk.Listbox(self, height=2)
        self.lb.pack(fill="x", padx=(0, 0))

    def _add(self):
        ps = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf"), ("Tất cả", "*.*")])
        for p in ps:
            if p not in self.paths:
                self.paths.append(p); self.lb.insert("end", os.path.basename(p))

    def _remove(self):
        sel = list(self.lb.curselection())
        for i in reversed(sel):
            self.lb.delete(i); del self.paths[i]

    def get(self):
        return list(self.paths)


class DocumentSection(ttk.LabelFrame):
    """1 văn bản trong nhóm VĂN BẢN — file (chính + tài liệu gửi kèm) + loại VB/số ký hiệu/
    trích yếu RIÊNG của văn bản này (1 phiếu trình có thể gửi nhiều văn bản khác loại/khác số).
    Chọn file chính sẽ tự điền loại VB/số/trích yếu (và tự chọn Luồng trình dùng chung) như
    màn hình cũ, chỉ khác là áp dụng riêng cho văn bản này."""

    def __init__(self, parent, app, on_remove):
        super().__init__(parent, text="Văn bản", padding=6)
        self.app = app
        self.on_remove = on_remove
        self._extract_seq = 0
        self.file_draft = tk.StringVar()

        # BƯỚC bạn cần làm — nổi bật, chữ đậm, luôn hiện.
        f = ttk.Frame(self); f.pack(fill="x", pady=2)
        ttk.Label(f, text="  Dự thảo văn bản:", width=18, font=("", 10, "bold")).pack(side="left")
        ttk.Entry(f, textvariable=self.file_draft).pack(side="left", fill="x", expand=True)
        ttk.Button(f, text="Chọn…", command=self._pick).pack(side="left")
        self.btn_remove = ttk.Button(f, text="✕", width=3, command=lambda: self.on_remove(self))
        self.btn_remove.pack(side="left", padx=(4, 0))

        self.extra = FileList(self, "  + Tài liệu gửi kèm:")

        f = ttk.Frame(self); f.pack(fill="x", pady=2)
        ttk.Label(f, text="  Loại văn bản:", width=18).pack(side="left")
        self.doc_type = ttk.Combobox(f, values=sorted(ENUMS["documentType"].keys()), state="readonly")
        self.doc_type.set("Công văn"); self.doc_type.pack(side="left", fill="x", expand=True)

        f = ttk.Frame(self); f.pack(fill="x", pady=2)
        ttk.Label(f, text="  Số/ký hiệu:", width=18).pack(side="left")
        self.code = ttk.Entry(f); self.code.pack(side="left", fill="x", expand=True)

        f = ttk.Frame(self); f.pack(fill="x", pady=2)
        ttk.Label(f, text="  Trích yếu:", width=18).pack(side="left")
        # tk.Text nhiều dòng thay vì ttk.Entry 1 dòng — trích yếu thường dài, Entry 1 dòng phải
        # cuộn ngang bằng chuột mới đọc hết được.
        self.abstract = tk.Text(f, height=2, wrap="word")
        self.abstract.pack(side="left", fill="x", expand=True)

    def _pick(self):
        p = filedialog.askopenfilename(
            filetypes=[("PDF/Word", "*.pdf *.docx"), ("Tất cả", "*.*")])
        if not p:
            return
        self.file_draft.set(p)
        self._extract_seq += 1
        seq = self._extract_seq
        threading.Thread(target=self._prepare_worker, args=(p, seq), daemon=True).start()

    def _prepare_worker(self, path, seq):
        """Đọc nhanh để tự điền ngay khi vừa chọn file — .docx qua python-docx (không cần
        Word), .pdf qua pypdf như cũ. Việc chuyển .docx sang PDF THẬT không còn làm ở đây nữa —
        dời sang lúc bấm CHẠY (xem App._run/_run_after_conversion): gộp lại thành 1 lượt gọi
        Word DUY NHẤT, tuần tự cho mọi văn bản trong phiếu trình, tránh nhiều luồng nền cùng lúc
        tranh nhau 1 tiến trình Word (từng gây lỗi 'Word.Application.Quit'), đồng thời đảm bảo
        PDF luôn sẵn sàng trước khi khung Xem trước mở ra."""
        if path.lower().endswith(".docx"):
            self._extract_docx_worker(path, seq)
            return
        if not path.lower().endswith(".pdf"):
            return   # định dạng khác PDF/Word — không tự đọc được, để nguyên cho người dùng tự điền
        self._extract_worker(path, seq)

    def _extract_docx_worker(self, path, seq):
        """Đọc thẳng .docx bằng python-docx, không cần Word."""
        doc_type = code = abstract = None
        err = None
        try:
            doc_type, code, abstract = extract_draft_fields_docx(path)
        except Exception as e:
            err = str(e)
        if err:
            self.app.after(0, lambda: self.app.log(
                f"• Không đọc nhanh được .docx để tự điền: {err} — vẫn có thể tự điền lại "
                "sau khi bấm CHẠY (lúc đó có bản PDF)."))
            return
        self.app.after(0, self._apply_extract, seq, doc_type, code, abstract, None)

    def _extract_worker(self, path, seq):
        doc_type = code = abstract = None
        err = None
        try:
            doc_type, code, abstract = extract_draft_fields(path)
        except Exception as e:
            err = str(e)
        self.app.after(0, self._apply_extract, seq, doc_type, code, abstract, err)

    def _apply_extract(self, seq, doc_type, code, abstract, err):
        if seq != self._extract_seq:
            return   # đã chọn file khác cho văn bản này trong lúc đọc — bỏ kết quả cũ
        if err:
            self.app.log(f"• Không đọc được PDF để tự điền: {err}")
            return
        if doc_type:
            self.doc_type.set(doc_type)
        if code:
            self.code.delete(0, "end"); self.code.insert(0, code)
            fid = flow_id_for_doc(doc_type, code, self.app.flow_store)
            fname = self.app._flow_name_for_id(fid) if fid else None
            if fname:
                self.app.flow.set(fname)
                self.app._on_flow_changed()
        if abstract:
            self.abstract.delete("1.0", "end"); self.abstract.insert("1.0", abstract)
            if not self.app.report_content.get("1.0", "end-1c").strip():   # không ghi đè nếu đã có/văn bản khác đã điền
                self.app.report_content.delete("1.0", "end"); self.app.report_content.insert("1.0", abstract)
        bits = []
        if doc_type: bits.append(f"Loại VB={doc_type!r}")
        if code: bits.append(f"Số/ký hiệu={code!r}")
        if abstract: bits.append(f"Trích yếu={abstract!r}")
        if bits:
            self.app.log("• Đã tự điền từ PDF: " + "  ".join(bits))
        else:
            self.app.log("• Không tìm thấy Số/ký hiệu hoặc Trích yếu trong PDF — điền tay.")

    def get(self):
        return {
            "doc_type": self.doc_type.get(),
            "code": self.code.get(),
            "abstract": self.abstract.get("1.0", "end-1c"),
            "file_draft_main": self.file_draft.get(),
            "files_draft_extra": self.extra.get(),
        }


# ---------- Cuộn bằng con lăn chuột (dùng chung cho mọi Canvas cuộn được) ----------
# Không dùng cách bind_all/unbind_all trên "Enter"/"Leave" của chính Canvas: các widget con
# (Entry, Label, ảnh...) phủ gần kín Canvas nên "Leave" kích hoạt gần như ngay khi con trỏ
# chạm vào bất kỳ widget con nào — khiến cuộn chuột gần như không hoạt động. Thay vào đó,
# đăng ký MỖI Canvas cuộn được vào 1 sổ dùng chung, và bind_all 1 LẦN DUY NHẤT cho cả ứng
# dụng; khi có sự kiện, tra xem con trỏ đang ở trong Canvas nào (dò lên theo .master) rồi
# cuộn đúng Canvas đó. Nhiều cửa sổ cùng có Canvas cuộn được (VD màn chính + khung xem trước)
# vẫn hoạt động độc lập, không "unbind" mất phần cuộn của cửa sổ kia khi 1 cửa sổ đóng lại.
_SCROLLABLES = []   # [(canvas, cuộn_ngang: bool)]
_WHEEL_HANDLERS_BOUND = False

def _wheel_target(canvas, x_root, y_root):
    try:
        w = canvas.winfo_containing(x_root, y_root)
    except (tk.TclError, KeyError):
        return False
    while w is not None:
        if w is canvas:
            return True
        w = w.master
    return False

def _dispatch_wheel(e, horizontal):
    for canvas, allow_h in list(_SCROLLABLES):
        try:
            if not canvas.winfo_exists():
                _SCROLLABLES.remove((canvas, allow_h)); continue
        except tk.TclError:
            continue
        if horizontal and not allow_h:
            continue
        if _wheel_target(canvas, e.x_root, e.y_root):
            try:
                (canvas.xview_scroll if horizontal else canvas.yview_scroll)(
                    int(-e.delta / 120), "units")
            except tk.TclError:
                pass
            return

def bind_mousewheel_scroll(canvas, horizontal=False):
    """Cho phép cuộn `canvas` bằng con lăn chuột dù con trỏ đang ở trên widget con nào
    bên trong nó. `horizontal=True` để thêm Shift+cuộn = cuộn ngang."""
    global _WHEEL_HANDLERS_BOUND
    _SCROLLABLES.append((canvas, horizontal))
    canvas.bind("<Destroy>", lambda e: _SCROLLABLES.remove((canvas, horizontal))
                if (canvas, horizontal) in _SCROLLABLES else None)
    if not _WHEEL_HANDLERS_BOUND:
        canvas.bind_all("<MouseWheel>", lambda e: _dispatch_wheel(e, False))
        canvas.bind_all("<Shift-MouseWheel>", lambda e: _dispatch_wheel(e, True))
        _WHEEL_HANDLERS_BOUND = True

def make_scrollable_frame(parent):
    """Bọc nội dung trong Canvas + thanh cuộn dọc + cuộn chuột. Trả về frame bên trong để
    pack nội dung vào."""
    outer = ttk.Frame(parent); outer.pack(fill="both", expand=True)
    canvas = tk.Canvas(outer, highlightthickness=0)
    vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    canvas.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    inner = ttk.Frame(canvas)
    win = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_inner_configure(_e):
        canvas.configure(scrollregion=canvas.bbox("all"))
    inner.bind("<Configure>", _on_inner_configure)

    def _on_canvas_configure(e):
        canvas.itemconfig(win, width=e.width)
    canvas.bind("<Configure>", _on_canvas_configure)

    bind_mousewheel_scroll(canvas)
    return inner


class FlowSignerPanel(ttk.LabelFrame):
    """Tải danh sách bước của LUỒNG ĐANG CHỌN (bất kỳ luồng nào, không riêng luồng nào) — bước
    nào đã có sẵn người (vd luồng cá nhân tự tạo, bake sẵn người ký) thì bỏ qua âm thầm; bước
    nào máy tự biết chắc người (đúng 1 ứng viên) thì tự điền; bước nào không chắc (≥2 ứng viên)
    thì hiện ô cho người dùng chọn. Tự ẩn nếu luồng đang chọn đã có sẵn đủ người (không cần hỏi
    gì thêm) — tự hiện nếu có ít nhất 1 bước cần tự điền/chọn, để minh bạch cho người dùng thấy."""
    def __init__(self, parent, session, log, flow_store):
        super().__init__(parent, text="Chọn người ký cho luồng vừa chọn", padding=8)
        self.session = session
        self.log = log
        self.flow_store = flow_store   # sổ cục bộ — nhớ người hay chọn cho từng (luồng, bước)
        self._nodes = []          # kết quả resolve_flow_signers() cho luồng đang hiện
        self._pick_vars = {}      # nodeId -> (tk.StringVar, [candidates])
        self._loading = False
        self._flow_id = None
        self.status = ttk.Label(self, text="", foreground="gray")
        self.status.pack(anchor="w")
        self.rows_frame = ttk.Frame(self)
        self.rows_frame.pack(fill="x", pady=(4, 0))

    def load(self, flow_id, doc_abstract):
        """Tải lại toàn bộ cho `flow_id` mới — huỷ kết quả cũ. Tự ẩn trong lúc tải (tránh hiện
        nhầm kết quả của luồng trước đó); CHẠY/GỬI bị chặn qua is_ready() cho tới khi xong."""
        self._flow_id = flow_id
        self._nodes = []
        self._pick_vars = {}
        self._loading = True
        self.pack_forget()
        for w in self.rows_frame.winfo_children():
            w.destroy()
        self.status.config(text="Đang tải người ký cho luồng này…", foreground="gray")

        def worker():
            try:
                nodes = resolve_flow_signers(self.session, flow_id, doc_abstract, self.log)
            except Exception as e:
                err_msg = str(e)   # tính ngay trong khối except — "e" bị Python xoá khi except kết thúc
                self.after(0, lambda: self._load_failed(flow_id, err_msg))
                return
            self.after(0, lambda: self._load_done(flow_id, nodes))
        threading.Thread(target=worker, daemon=True).start()

    def _load_failed(self, flow_id, e):
        if flow_id != self._flow_id:
            return   # đã đổi sang luồng khác trong lúc tải — bỏ kết quả cũ
        self._loading = False
        self.status.config(text=f"✖ Không tải được: {e}", foreground="#c62828")
        self.pack(fill="x", pady=(6, 0))

    def _load_done(self, flow_id, nodes):
        if flow_id != self._flow_id:
            return
        self._loading = False
        self._nodes = nodes
        need_pick = [n for n in nodes if len(n.get("candidates") or []) >= 2]
        auto = [n for n in nodes if n.get("userId") is not None and len(n.get("candidates") or []) < 2]
        if not need_pick and not any(n.get("candidates") for n in nodes):
            self.pack_forget()   # luồng đã có sẵn đủ người (vd luồng cá nhân) — không cần hỏi gì
            return
        self.pack(fill="x", pady=(6, 0))
        self.status.config(
            text=f"{len(nodes)} bước — {len(auto)} tự điền, {len(need_pick)} cần bạn chọn.",
            foreground="gray")
        for n in nodes:
            # Xếp DỌC (tên bước ở trên, combobox/tên người ở dưới) thay vì nằm ngang — nằm
            # ngang trước đây chia đôi bề rộng hàng cho nhãn + combobox, tên người dài dễ bị
            # cắt chữ, phải kéo rộng cả cửa sổ mới đọc hết. Xếp dọc thì combobox luôn được toàn
            # bộ bề rộng khung, không phụ thuộc cửa sổ to hay nhỏ.
            row = ttk.Frame(self.rows_frame); row.pack(fill="x", pady=(2, 4))
            ttk.Label(row, text=n.get("name") or "(không rõ tên bước)", wraplength=380,
                      justify="left", foreground="gray").pack(anchor="w", fill="x")
            cands = n.get("candidates") or []
            if len(cands) >= 2:
                # Ghi kèm chức danh vào nhãn — 1 bước có thể gồm nhiều biến thể chức danh khác
                # nhau (vd "Cục Trưởng" chỉ 1 người, "Phó Cục Trưởng" 3 người khác), phải phân
                # biệt rõ đang chọn ai VỚI TƯ CÁCH gì, không chỉ chọn tên.
                pref_id = preferred_signer(self.flow_store, flow_id, n["nodeId"], cands)
                default_idx = 0
                if pref_id is not None:
                    default_idx = next(i for i, c in enumerate(cands) if c.get("userId") == pref_id)
                names = [f"{c.get('fullName') or ''} — {c.get('roleName') or ''}"
                         + ("  (lần trước)" if i == default_idx and pref_id is not None else "")
                         for i, c in enumerate(cands)]
                var = tk.StringVar(value=names[default_idx])
                n["userId"] = cands[default_idx].get("userId")
                n["fullName"] = cands[default_idx].get("fullName")
                n["roleId"] = cands[default_idx].get("roleId")
                n["roleName"] = cands[default_idx].get("roleName")
                cb = ttk.Combobox(row, values=names, textvariable=var, state="readonly")
                cb.pack(fill="x")
                def on_pick(_e, n=n, var=var, cands=cands, names=names):
                    idx = names.index(var.get())
                    n["userId"] = cands[idx].get("userId")
                    n["fullName"] = cands[idx].get("fullName")
                    n["roleId"] = cands[idx].get("roleId")
                    n["roleName"] = cands[idx].get("roleName")
                cb.bind("<<ComboboxSelected>>", on_pick)
                self._pick_vars[n["nodeId"]] = (var, cands)
            else:
                name = n.get("fullName") or "(chưa xác định được người)"
                fg = "black" if n.get("userId") is not None else "#c62828"
                ttk.Label(row, text=name, foreground=fg).pack(anchor="w")

    def is_ready(self):
        """False nếu còn đang tải, hoặc còn bước nào chưa có người (kể cả chưa chọn xong)."""
        if self._loading or not self._nodes:
            return False
        return all(n.get("userId") is not None for n in self._nodes)

    def get_nodes(self):
        return self._nodes


class _ConvertingDialog(tk.Toplevel):
    """Cửa sổ chờ nhỏ, hiện trong lúc chuyển .docx -> PDF tuần tự ngay sau khi bấm CHẠY (trước
    khi mở khung Xem trước hoặc chạy Chỉ kiểm tra) — để người dùng biết đang có việc chạy ngầm
    (Word có thể mất vài giây tới vài chục giây), không phải chương trình bị treo."""

    def __init__(self, master, text):
        super().__init__(master)
        self.title("Đang chuẩn bị…")
        self.resizable(False, False)
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", lambda: None)   # không cho tự đóng giữa chừng
        self.label = ttk.Label(self, text=text, padding=24)
        self.label.pack()
        self.update_idletasks()
        x = master.winfo_rootx() + master.winfo_width() // 2 - self.winfo_width() // 2
        y = master.winfo_rooty() + master.winfo_height() // 2 - self.winfo_height() // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.grab_set()

    def set_status(self, text):
        self.label.config(text=text)

    def close(self):
        try:
            self.grab_release()   # xem _close_dlg() ở _open_role_dialog — tránh kẹt input trên macOS
        except Exception:
            pass
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Trợ lý trình văn bản")
        self.session = None
        self.settings = load_settings()
        self.store = load_store()
        self.flow_store = load_flow_store()
        self.file_report = tk.StringVar()
        self.doc_sections = []
        self.container = ttk.Frame(self); self.container.pack(fill="both", expand=True)
        self._show_login()

    def _clear(self):
        for w in self.container.winfo_children():
            w.destroy()

    def _row(self, parent, label, bold=False):
        f = ttk.Frame(parent); f.pack(fill="x", padx=12, pady=3)
        font = ("", 10, "bold") if bold else ("", 10)
        ttk.Label(f, text=label, width=18, font=font).pack(side="left")
        return f

    # ---------- MÀN 1: ĐĂNG NHẬP ----------
    def _show_login(self):
        self._clear()
        self.geometry("440x380")
        pad = ttk.Frame(self.container, padding=16); pad.pack(fill="both", expand=True)
        ttk.Label(pad, text="Đăng nhập hệ thống", font=("", 12, "bold")).pack(anchor="w", pady=(0, 8))
        f = self._row(pad, "Tên đăng nhập:")
        self.username = ttk.Entry(f); self.username.pack(side="left", fill="x", expand=True)
        f = self._row(pad, "Mật khẩu:")
        self.password = ttk.Entry(f, show="•"); self.password.pack(side="left", fill="x", expand=True)
        self.password.bind("<Return>", lambda e: self._do_login())

        self.remember = tk.BooleanVar(value=bool(self.settings.get("remember")))
        ttk.Checkbutton(pad, text="Nhớ đăng nhập (mật khẩu lưu an toàn trong Keychain)",
                        variable=self.remember).pack(anchor="w", pady=2)

        # Điền sẵn từ lần trước
        saved_user = self.settings.get("username", "")
        if saved_user:
            self.username.insert(0, saved_user)
            if self.settings.get("remember"):
                pw = load_password(saved_user)
                if pw:
                    self.password.insert(0, pw)
        self.password.focus() if saved_user else self.username.focus()

        self.btn_login = ttk.Button(pad, text="ĐĂNG NHẬP", command=self._do_login)
        self.btn_login.pack(pady=10)
        if not keyring:
            ttk.Label(pad, text="(Muốn nhớ mật khẩu: chạy  pip install keyring)",
                      foreground="gray").pack(anchor="w")
        self.login_log = tk.Text(pad, height=6, bg="#101418", fg="#d0d0d0")
        self.login_log.pack(fill="both", expand=True)

        ttk.Label(pad, text=AUTHOR_MARK, font=("", 8), foreground="#999999").pack(
            anchor="e", pady=(4, 0))

    def _llog(self, msg):
        self.login_log.insert("end", msg + "\n"); self.login_log.see("end"); self.update_idletasks()

    def _do_login(self):
        u, p = self.username.get().strip(), self.password.get()
        if not u or not p:
            messagebox.showwarning("Thiếu", "Nhập tên đăng nhập và mật khẩu."); return
        self.btn_login.config(state="disabled")
        self.login_log.delete("1.0", "end")
        s = make_session()
        try:
            cas_login(s, u, p, self._llog, self._ask_captcha)   # chạy luồng chính (captcha an toàn)
        except PipelineError as e:
            self._llog("✖ " + str(e)); self.btn_login.config(state="normal"); return
        except Exception as e:
            self._llog("✖ Lỗi đăng nhập: " + repr(e)); self.btn_login.config(state="normal"); return
        self.session = s
        self._logged_user = u          # lưu lại trước khi xóa màn login
        # Nhớ đăng nhập
        if self.remember.get():
            self.settings["username"] = u
            self.settings["remember"] = True
            if save_password(u, p):
                pass
            else:
                self._llog("   (Không lưu được mật khẩu — thiếu 'keyring'. Chỉ nhớ tên đăng nhập.)")
        else:
            self.settings["username"] = u   # vẫn nhớ tên cho tiện
            self.settings["remember"] = False
            try:
                if keyring:
                    keyring.delete_password(KR_SERVICE, u)
            except Exception:
                pass
        save_settings(self.settings)
        self._show_main()   # thành công → mở giao diện chính

    # ---------- Khung cuộn (dùng chung) ----------
    def _make_scrollable(self, parent):
        return make_scrollable_frame(parent)

    # ---------- MÀN 2: SOẠN & LƯU NHÁP ----------
    def _show_main(self):
        self._clear()
        # Tự tính theo ĐÚNG màn hình đang chạy — không đoán 1 số cố định cho mọi máy (dễ tràn
        # màn hình nhỏ hoặc quá bé so với màn hình lớn). Có trần hợp lý để không quá khổ trên
        # màn hình rất lớn.
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(int(sw * 0.85), 1400), min(int(sh * 0.85), 950)
        self.geometry(f"{w}x{h}")
        self.minsize(480, 360)
        pad = self._make_scrollable(self.container)
        ttk.Label(pad, text=f"Đã đăng nhập: {getattr(self, '_logged_user', '')}",
                  foreground="#2e7d32", font=("", 9, "bold")).pack(anchor="w", padx=12, pady=(10, 4))

        # ---- Phần trên: 2 cột — trái = file đính kèm, phải = thông tin văn bản ----
        # Dùng grid (không phải pack) với 2 cột "uniform" bằng nhau — pack(expand=True) trước
        # đây KHÔNG đảm bảo chia đôi thật sự, cột nào có nội dung "đòi" nhiều chỗ hơn (vd cột
        # trái nhiều Entry/Listbox) sẽ lấn cột kia, khiến cột phải bị bóp hẹp và cắt chữ.
        top = ttk.Frame(pad); top.pack(fill="x", padx=12, pady=(4, 0))
        top.columnconfigure(0, weight=1, uniform="cols")
        top.columnconfigure(1, weight=1, uniform="cols")
        left_col = ttk.Frame(top); left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        right_col = ttk.Frame(top); right_col.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        # --- Cột trái: File — 2 nhóm, mỗi nhóm: file chính + tài liệu thêm ---
        g1 = ttk.LabelFrame(left_col, text="Nhóm PHIẾU TRÌNH (không gửi đi)", padding=6)
        g1.pack(fill="x")
        f = self._row(g1, "  File phiếu trình:", bold=True)
        ttk.Entry(f, textvariable=self.file_report).pack(side="left", fill="x", expand=True)
        ttk.Button(f, text="Chọn…", command=self._pick_file_report).pack(side="left")
        self.extra_report = FileList(g1, "  + Tài liệu thêm (không gửi):")

        g2 = ttk.LabelFrame(left_col, text="Nhóm VĂN BẢN (gửi đi)", padding=6)
        g2.pack(fill="x", pady=(6, 0))
        self.doc_sections_frame = ttk.Frame(g2); self.doc_sections_frame.pack(fill="x")
        self.doc_sections = []
        self._add_document_section()
        ttk.Button(g2, text="+ Thêm văn bản", command=self._add_document_section).pack(anchor="w", pady=(4, 0))
        self.auto_stamp_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(g2, text="Tự đánh số chữ ký lên Phiếu trình + các Văn bản theo Luồng trình đã chọn",
                        variable=self.auto_stamp_var).pack(anchor="w", pady=(2, 0))
        ttk.Label(left_col, text="(File chính của mỗi văn bản = cái cần ký. Bỏ trống hết để thử nghiệm sẽ dùng tạm file phiếu trình.)",
                  foreground="gray", wraplength=300).pack(anchor="w", pady=(4, 0))

        # --- Cột phải: Luồng trình (dùng chung) + khẩn/mật + nội dung phiếu ---
        # Danh sách luồng LUÔN lấy động từ web (theo đúng tài khoản đang đăng nhập) — không còn
        # 3 luồng cứng trong code nữa. Trong lúc chờ tải, combobox tạm hiện 1 dòng placeholder.
        self._FLOW_LOADING = "(đang tải danh sách luồng…)"
        self._flow_by_name = {}
        f = self._row(right_col, "Luồng trình:")
        self.flow = ttk.Combobox(f, values=[self._FLOW_LOADING], state="readonly")
        self.flow.set(self._FLOW_LOADING)
        self.flow.pack(side="left", fill="x", expand=True)
        self.flow.bind("<<ComboboxSelected>>", self._on_flow_changed)

        self.flow_panel = FlowSignerPanel(right_col, self.session, self.log, self.flow_store)
        # chưa pack() — panel tự hiện/ẩn tuỳ luồng đang chọn đã có sẵn đủ người hay chưa

        # Hồ sơ công việc — BẮT BUỘC lấy đúng theo tài khoản đang đăng nhập (xem
        # fetch_prepare_insert_data): dùng nhầm hồ sơ của tài khoản khác vẫn "lưu thành công"
        # nhưng phiếu trình lạc mất, không thấy trong thùng nháp.
        self._profile_by_name = {}
        f = self._row(right_col, "Hồ sơ công việc:")
        self.work_profile = ttk.Combobox(f, values=[], state="readonly")
        self.work_profile.pack(side="left", fill="x", expand=True)

        # Khẩn + mật
        f = self._row(right_col, "Độ khẩn:")
        self.priority = ttk.Combobox(f, values=[""] + sorted(ENUMS["priority"].keys()), state="readonly")
        self.priority.set("Khẩn"); self.priority.pack(side="left", fill="x", expand=True)
        f = self._row(right_col, "Độ mật:")
        self.security = ttk.Combobox(f, values=[""] + sorted(ENUMS["security"].keys()), state="readonly")
        self.security.set("Bình thường"); self.security.pack(side="left", fill="x", expand=True)

        # Nội dung phiếu trình
        f = self._row(right_col, "Nội dung phiếu:")
        # tk.Text nhiều dòng thay vì ttk.Entry 1 dòng — cùng lý do với "Trích yếu" ở DocumentSection.
        self.report_content = tk.Text(f, height=2, wrap="word")
        self.report_content.pack(side="left", fill="x", expand=True)

        # ---- Phần dưới: Nơi nhận (rộng hết chiều ngang) ----
        self.recip = RecipientBox(pad, CAY, self.store)
        self.recip.pack(fill="x", padx=12, pady=(10, 0))

        # Nút
        f = ttk.Frame(pad); f.pack(fill="x", padx=12, pady=8)
        self.check_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Chỉ kiểm tra (không ghi gì) — bật cho lần chạy đầu",
                        variable=self.check_var).pack(anchor="w")
        ttk.Button(f, text="CHẠY", command=self._run).pack(side="left", pady=4)
        self.readiness_label = ttk.Label(f, text="", font=("", 9))
        self.readiness_label.pack(side="left", padx=(10, 0))
        self._refresh_readiness()   # tự cập nhật định kỳ — xem _refresh_readiness()

        # Log
        self.logbox = tk.Text(pad, height=10, bg="#101418", fg="#d0d0d0")
        self.logbox.pack(fill="both", expand=True, padx=12, pady=(6, 12))

        ttk.Label(pad, text=AUTHOR_MARK, font=("", 8), foreground="#999999").pack(
            anchor="e", padx=12, pady=(0, 6))

        self._fetch_flow_data()   # sau cùng — logbox đã có sẵn để self.log() gọi từ luồng nền

    def _pick(self, var):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf"), ("Tất cả", "*.*")])
        if p: var.set(p)

    def _pick_file_report(self):
        """Chọn File phiếu trình — khác _pick() thường ở chỗ: nếu là .docx, tự chuyển
        sang PDF trước (nền); rồi đọc "Nội dung phiếu" ngay từ chính phiếu trình (xem
        extract_phieu_trinh_content), GHI ĐÈ ô đó (nội dung phiếu trình luôn đáng tin hơn trích
        yếu văn bản — trích yếu chỉ điền khi ô còn trống, xem DocumentSection._apply_extract)."""
        p = filedialog.askopenfilename(
            filetypes=[("PDF/Word", "*.pdf *.docx"), ("Tất cả", "*.*")])
        if not p:
            return
        self.file_report.set(p)
        threading.Thread(target=self._prepare_report_worker, args=(p,), daemon=True).start()

    def _prepare_report_worker(self, path):
        """Đọc nhanh để tự điền "Nội dung phiếu" ngay khi vừa chọn file — .docx qua python-docx
        (không cần Word), .pdf qua pymupdf như cũ. Việc chuyển .docx sang PDF thật dời sang lúc
        bấm CHẠY (xem _run()/_run_after_conversion) — xem lý do ở DocumentSection._prepare_worker."""
        if path.lower().endswith(".docx"):
            self._extract_report_docx_worker(path)
            return
        if not path.lower().endswith(".pdf"):
            return   # định dạng khác PDF/Word — không tự đọc được, để nguyên cho người dùng tự điền
        self._extract_report_worker(path)

    def _extract_report_docx_worker(self, path):
        """Đọc thẳng .docx bằng python-docx, không cần Word."""
        content, err = None, None
        try:
            content = extract_phieu_trinh_content_docx(path)
        except Exception as e:
            err = str(e)
        if err:
            self.after(0, lambda: self.log(
                f"• Không đọc nhanh được .docx để tự điền nội dung phiếu: {err} — vẫn có thể "
                "tự điền lại sau khi bấm CHẠY (lúc đó có bản PDF)."))
            return
        self.after(0, self._apply_report_extract, path, content, None)

    def _extract_report_worker(self, path):
        content, err = None, None
        try:
            content = extract_phieu_trinh_content(path)
        except Exception as e:
            err = str(e)
        self.after(0, self._apply_report_extract, path, content, err)

    def _apply_report_extract(self, path, content, err):
        if path != self.file_report.get():
            return   # đã chọn file phiếu trình khác trong lúc đọc — bỏ kết quả cũ
        if err:
            self.log(f"• Không đọc được nội dung phiếu trình để tự điền: {err}")
            return
        if content:
            self.report_content.delete("1.0", "end")
            self.report_content.insert("1.0", content)
            self.log(f"• Đã tự điền Nội dung phiếu từ file phiếu trình: {content!r}")
        else:
            self.log("• Không nhận ra được nội dung trong file phiếu trình (mẫu khác/PDF quét "
                      "ảnh không có lớp chữ) — điền tay nếu cần.")

    FLOW_SEPARATOR = "──────── Luồng khác (từ web) ────────"
    FLOW_QUEN_MAX = 5   # số luồng "quen" tối đa hiện ở đầu combobox (ghim + hay dùng nhất)

    def _fetch_flow_data(self):
        """Lấy (1 lần, cùng 1 request prepareInsert.do): toàn bộ luồng của tài khoản đang đăng
        nhập, VÀ toàn bộ hồ sơ công việc — cả 2 đều đúng theo tài khoản đang đăng nhập, không
        phải giá trị tĩnh ghi cứng trong luong_trinh.json/du_lieu.json."""
        def worker():
            try:
                flows, profiles = fetch_prepare_insert_data(self.session, self.log)
            except Exception as e:
                self.log(f"• Không lấy được danh sách luồng/hồ sơ công việc từ web: {e}")
                return
            self.after(0, lambda: self._apply_flow_list(flows))
            self.after(0, lambda: self._apply_work_profiles(profiles))
        threading.Thread(target=worker, daemon=True).start()

    def _apply_flow_list(self, live_flows):
        """Xếp lên đầu combobox tối đa FLOW_QUEN_MAX luồng: ưu tiên đã ghim, rồi theo tần suất
        dùng TRÊN MÁY NÀY (luong_trinh.json — sổ cục bộ, không dùng chung cho mọi người) — phần
        còn lại xếp dưới 1 dòng phân cách, giữ nguyên thứ tự web trả về."""
        by_id = {fl["flowId"]: fl["name"] for fl in live_flows}
        pinned = set(self.flow_store.get("pinned", []))
        freq = self.flow_store.get("freq", {})
        known_ids = [fid for fid in by_id if fid in pinned or freq.get(fid, 0) > 0]
        known_ids.sort(key=lambda fid: (0 if fid in pinned else 1, -freq.get(fid, 0)))
        quen_ids = known_ids[:self.FLOW_QUEN_MAX]

        self._flow_by_name = {by_id[fid]: fid for fid in quen_ids}
        rest = [fl for fl in live_flows if fl["flowId"] not in quen_ids]
        values = list(self._flow_by_name.keys())
        if rest:
            values.append(self.FLOW_SEPARATOR)
            for fl in rest:
                self._flow_by_name.setdefault(fl["name"], fl["flowId"])
            values += [fl["name"] for fl in rest]

        if not values:
            self.log("   — Không lấy được luồng nào từ web (giữ nguyên ô trống, tự kiểm tra lại mạng).")
            return
        self.flow.config(values=values)
        self.flow.set(values[0])
        self._on_flow_changed()   # tải khung chọn người ngay cho lựa chọn mặc định

    def _apply_work_profiles(self, profiles):
        if not profiles:
            self.log("   — Không tìm thấy hồ sơ công việc nào của tài khoản này trên web.")
            return
        self._profile_by_name = {p["name"]: p["fileId"] for p in profiles}
        self.work_profile.config(values=list(self._profile_by_name.keys()))
        # Mặc định: ưu tiên hồ sơ có chữ "chung" (hồ sơ dùng chung, không gắn 1 vụ việc cụ thể);
        # nếu không có thì lấy hồ sơ đầu tiên trong danh sách của tài khoản này.
        default_name = next((n for n in self._profile_by_name if "chung" in n.lower()),
                             next(iter(self._profile_by_name)))
        self.work_profile.set(default_name)

    def _on_flow_changed(self, _event=None):
        name = self.flow.get()
        if name == self.FLOW_SEPARATOR:
            # dòng phân cách không phải luồng thật — quay lại luồng đầu danh sách
            self.flow.set(next(iter(self._flow_by_name), self._FLOW_LOADING))
            name = self.flow.get()
        flow_id = self._flow_by_name.get(name)
        if not flow_id:
            return   # danh sách luồng chưa tải xong (vẫn đang hiện placeholder) — chưa làm gì được
        self.flow_panel.load(flow_id, self.report_content.get("1.0", "end-1c"))

    def _flow_name_for_id(self, flow_id):
        for name, fid in self._flow_by_name.items():
            if fid == flow_id:
                return name
        return None

    def _add_document_section(self):
        ds = DocumentSection(self.doc_sections_frame, self, self._remove_document_section)
        ds.pack(fill="x", pady=(0, 6))
        self.doc_sections.append(ds)
        self._update_doc_section_remove_buttons()

    def _remove_document_section(self, ds):
        if len(self.doc_sections) <= 1:
            return   # luôn giữ ít nhất 1 văn bản
        self.doc_sections.remove(ds)
        ds.destroy()
        self._update_doc_section_remove_buttons()

    def _update_doc_section_remove_buttons(self):
        only_one = len(self.doc_sections) <= 1
        for ds in self.doc_sections:
            ds.btn_remove.config(state=("disabled" if only_one else "normal"))

    def log(self, msg):
        self.logbox.insert("end", msg + "\n"); self.logbox.see("end"); self.update_idletasks()

    def _collect_cfg(self):
        return {
            "priority": self.priority.get(),
            "security": self.security.get(),
            "report_content": self.report_content.get("1.0", "end-1c"),
            "flow_id": self._flow_by_name.get(self.flow.get()),
            "flow_name": self.flow.get(),
            "flow_nodes_override": self.flow_panel.get_nodes() or None,
            "work_profile_id": self._profile_by_name.get(self.work_profile.get()),
            "work_profile_name": self.work_profile.get(),
            "auto_stamp": self.auto_stamp_var.get(),
            "recv_inside": self.recip.get("inside"),
            "recv_report": self.recip.get("report"),
            "recv_edoc": self.recip.get("edoc"),
            "recv_save": self.recip.get("save"),
            "recv_know": self.recip.get("know"),
            "file_report_main": self.file_report.get(),
            "files_report_extra": self.extra_report.get(),
            "documents": [ds.get() for ds in self.doc_sections],
        }

    def _refresh_readiness(self):
        """1 dòng trạng thái cạnh nút CHẠY — liếc là biết đã đủ để bấm chưa, khỏi phải tự rà
        từng ô. Tự cập nhật định kỳ (không cần nối callback riêng vào từng ô/luồng/nơi nhận —
        đơn giản hơn, chi phí không đáng kể)."""
        missing = []
        if not self.file_report.get():
            missing.append("File phiếu trình")
        if not self.report_content.get("1.0", "end-1c").strip():
            missing.append("Nội dung phiếu")
        multi = len(self.doc_sections) > 1
        for i, ds in enumerate(self.doc_sections):
            tag = f" (văn bản {i+1})" if multi else ""
            if not ds.abstract.get("1.0", "end-1c").strip():
                missing.append(f"Trích yếu văn bản{tag}")
            if not ds.code.get().strip():
                missing.append(f"Số/ký hiệu{tag}")
        if not any(self.recip.get(c) for c, _ in RecipientBox.CATS):
            missing.append("Nơi nhận")
        if not self.flow_panel.is_ready():
            missing.append("chọn người ký cho luồng")
        if missing:
            self.readiness_label.config(text="⚠ Còn thiếu: " + ", ".join(missing), foreground="#c62828")
        else:
            self.readiness_label.config(text="✓ Sẵn sàng", foreground="#2e7d32")
        self.after(700, self._refresh_readiness)

    def _maybe_learn_flow_rule(self, flow_id, cfg):
        """Nếu luồng đang chọn KHÔNG do quy tắc nào tự suy ra được (từ khoá lạ, chưa có trong
        sổ), hỏi có muốn nhớ quy tắc mới không — để lần sau ký hiệu tương tự tự chọn đúng luồng
        này, người dùng mới không phải tự sửa tay luong_trinh.json."""
        code = next((d.get("code") for d in (cfg.get("documents") or []) if d.get("code")), None)
        if not code or flow_id_for_code(code, self.flow_store):
            return   # không có ký hiệu để học, hoặc đã có quy tắc khớp sẵn rồi
        kw = flow_keyword_from_code(code)
        if not kw:
            return
        flow_name = cfg.get("flow_name") or flow_id
        if messagebox.askyesno("Nhớ quy tắc chọn luồng?",
                f"Ký hiệu văn bản có '{kw}' — từ nay tự động chọn luồng '{flow_name}' khi gặp "
                f"ký hiệu có '{kw}' không?"):
            self.flow_store.setdefault("rules", []).insert(0, {"keyword": kw, "flowId": flow_id})
            save_flow_store(self.flow_store)
            self.log(f"• Đã nhớ quy tắc: ký hiệu có '{kw}' → luồng '{flow_name}'.")

    def _run(self):
        """Bấm CHẠY: nếu còn file .docx (phiếu trình hoặc bất kỳ văn bản nào) chưa có bản PDF,
        chuyển hết TUẦN TỰ (1 lượt gọi Word duy nhất cho cả phiếu trình, tránh nhiều luồng nền
        tranh nhau 1 tiến trình Word — xem DocumentSection._prepare_worker) rồi mới tiếp tục —
        áp dụng NHƯ NHAU cho cả 'Chỉ kiểm tra' lẫn chạy thật, vì cả 2 đều cần PDF thật (kể cả
        Chỉ kiểm tra cũng đánh số chữ ký lên PDF, xem bước tương ứng trong run_pipeline())."""
        if not self.flow_panel.is_ready():
            messagebox.showwarning("Chưa xong",
                                    "Luồng trình này còn bước chưa xác định được người ký "
                                    "(đang tải hoặc chưa chọn đủ, hoặc danh sách luồng chưa tải xong). "
                                    "Đợi hoặc chọn xong rồi bấm CHẠY lại.")
            return
        if not self.file_report.get():
            messagebox.showwarning("Thiếu", "Chưa chọn file phiếu trình."); return

        pending = []   # [(kind, idx_hoặc_None, path)] — mọi file CHÍNH còn là .docx
        if self.file_report.get().lower().endswith(".docx"):
            pending.append(("report", None, self.file_report.get()))
        for i, ds in enumerate(self.doc_sections):
            p = ds.file_draft.get()
            if p.lower().endswith(".docx"):
                pending.append(("draft", i, p))

        if not pending:
            self._run_after_conversion()
            return

        total = len(pending)
        dlg = _ConvertingDialog(self, f"Đang chuyển sang PDF (1/{total})…")
        def worker():
            for k, (kind, idx, path) in enumerate(pending, start=1):
                self.after(0, lambda p=path, k=k: dlg.set_status(
                    f"Đang chuyển sang PDF ({k}/{total}): {os.path.basename(p)}…"))
                try:
                    pdf_path = convert_office_doc_to_pdf(path)
                except Exception as e:
                    err = str(e)   # tính ngay trong khối except — "e" bị Python xoá khi except kết thúc
                    self.after(0, lambda path=path, err=err: self._conversion_run_failed(dlg, path, err))
                    return
                self.after(0, lambda kind=kind, idx=idx, pdf_path=pdf_path:
                           self._apply_run_conversion(kind, idx, pdf_path))
            self.after(0, lambda: (dlg.close(), self._run_after_conversion()))
        threading.Thread(target=worker, daemon=True).start()

    def _conversion_run_failed(self, dlg, path, err):
        dlg.close()
        messagebox.showerror("Không chuyển được sang PDF",
            f"File: {os.path.basename(path)}\n\nLỗi: {err}\n\n"
            "Chưa thể tiếp tục (kể cả 'Chỉ kiểm tra') khi file này chưa có bản PDF. "
            "Bấm CHẠY thử lại, hoặc tự mở Word 'Save As' sang PDF rồi chọn lại file đó.")

    def _apply_run_conversion(self, kind, idx, pdf_path):
        """Cập nhật đường dẫn sau khi convert xong 1 file, và tự điền NỐT các ô còn trống —
        đọc từ chính bản PDF vừa có (đáng tin hơn/đầy đủ hơn bản đọc nhanh lúc chọn file, vì
        giờ đã chắc chắn có PDF). CHỈ điền ô đang trống, không đè lên bất kỳ ô nào người dùng
        (hoặc bước đọc nhanh lúc chọn file) đã điền trước đó."""
        self.log(f"• Đã chuyển sang PDF: {os.path.basename(pdf_path)}")
        if kind == "report":
            self.file_report.set(pdf_path)
            try:
                content = extract_phieu_trinh_content(pdf_path)
            except Exception as e:
                self.log(f"• Không đọc được nội dung phiếu trình từ PDF vừa chuyển: {e}")
                return
            if content and not self.report_content.get("1.0", "end-1c").strip():
                self.report_content.delete("1.0", "end"); self.report_content.insert("1.0", content)
                self.log(f"• Đã tự điền thêm Nội dung phiếu từ PDF vừa chuyển: {content!r}")
            return

        ds = self.doc_sections[idx]
        ds.file_draft.set(pdf_path)
        try:
            doc_type, code, abstract = extract_draft_fields(pdf_path)
        except Exception as e:
            self.log(f"• Không đọc được văn bản từ PDF vừa chuyển: {e}")
            return
        if doc_type and not ds.doc_type.get():
            ds.doc_type.set(doc_type)
        if code and not ds.code.get().strip():
            ds.code.delete(0, "end"); ds.code.insert(0, code)
            fid = flow_id_for_doc(doc_type, code, self.flow_store)
            fname = self._flow_name_for_id(fid) if fid else None
            if fname:
                self.flow.set(fname)
                self._on_flow_changed()
        if abstract and not ds.abstract.get("1.0", "end-1c").strip():
            ds.abstract.delete("1.0", "end"); ds.abstract.insert("1.0", abstract)
            if not self.report_content.get("1.0", "end-1c").strip():
                self.report_content.delete("1.0", "end"); self.report_content.insert("1.0", abstract)

    def _run_after_conversion(self):
        """Phần logic CHẠY thật sự — chạy sau khi mọi file .docx (nếu có) đã chuyển xong sang
        PDF, nên đọc lại cfg ở đây (không dùng cfg cũ) để lấy đúng đường dẫn .pdf mới."""
        flow_id = self._flow_by_name.get(self.flow.get())
        cfg = self._collect_cfg()

        if flow_id:
            bump_flow_freq(self.flow_store, flow_id)   # để lần sau luồng này tự nổi lên đầu combobox
            for n in self.flow_panel.get_nodes():
                if len(n.get("candidates") or []) >= 2:
                    remember_signer_pick(self.flow_store, flow_id, n["nodeId"], n.get("userId"))
            if not self.check_var.get():   # chỉ hỏi ở lần chạy thật, đỡ phiền lúc test "chỉ kiểm tra"
                self._maybe_learn_flow_rule(flow_id, cfg)

        if self.check_var.get():
            # Chế độ an toàn (lần đầu): chạy thẳng như trước, KHÔNG mở khung xem trước, không ghi gì.
            self.logbox.delete("1.0", "end")
            self.log("=== BẮT ĐẦU XỬ LÝ (CHỈ KIỂM TRA) ===")
            s = self.session
            def worker():
                try:
                    run_pipeline(s, cfg, self.log, check_only=True)
                except PipelineError as e:
                    self.log("\n✖ DỪNG: " + str(e))
                except Exception as e:
                    self.log("\n✖ LỖI KHÔNG NGỜ: " + repr(e))
            threading.Thread(target=worker, daemon=True).start()
            return

        if fitz is None:
            messagebox.showerror("Thiếu thư viện",
                                  "Khung xem trước cần thư viện 'pymupdf'. Chạy: pip install pymupdf")
            return
        PreviewWindow(self, cfg, self.session)

    def _ask_captcha(self, path):
        try:
            import webbrowser
            webbrowser.open("file://" + os.path.abspath(path))
        except Exception:
            pass
        self._llog(f"   (Ảnh captcha đã lưu: {path})")
        from tkinter import simpledialog
        return simpledialog.askstring("Captcha", "Nhập mã trong ảnh captcha vừa mở:")


# ---------- MÀN XEM TRƯỚC + GỬI ----------
# ---------- Checklist tiến trình + banner kết quả (cho khung Xem trước & Gửi) ----------
# Thay cho việc chỉ có 1 ô log chạy chữ kỹ thuật — người dùng phổ thông không có cách nào yên
# tâm là "đã xong chưa/có lỗi không" nếu phải tự đọc log. Checklist hiện các mốc lớn (xem
# PIPELINE_PHASES) tự tick khi xong; banner hiện kết quả cuối cùng to, rõ, có màu.
class StepChecklist(ttk.Frame):
    ICONS = {"pending": "○", "running": "⏳", "done": "✓", "error": "✕"}
    COLORS = {"pending": "#9e9e9e", "running": "#f57f17", "done": "#2e7d32", "error": "#c62828"}

    def __init__(self, parent, phases):
        """phases: [(key, nhãn tiếng Việt), ...] — xem PIPELINE_PHASES."""
        super().__init__(parent)
        self.order = [k for k, _ in phases]
        self._rows = {}
        for key, label in phases:
            row = ttk.Frame(self); row.pack(fill="x", anchor="w", pady=1)
            icon = ttk.Label(row, text=self.ICONS["pending"], width=2, foreground=self.COLORS["pending"])
            icon.pack(side="left")
            text = ttk.Label(row, text=label, foreground=self.COLORS["pending"])
            text.pack(side="left")
            self._rows[key] = (icon, text)

    def reset(self):
        for key in self.order:
            self.set_status(key, "pending")

    def set_status(self, key, status):
        if key not in self._rows:
            return
        icon, text = self._rows[key]
        icon.config(text=self.ICONS[status], foreground=self.COLORS[status])
        text.config(foreground=self.COLORS[status])

    def mark_running(self, key):
        """Đánh dấu mọi bước TRƯỚC `key` là xong, và `key` đang chạy."""
        idx = self.order.index(key)
        for k in self.order[:idx]:
            self.set_status(k, "done")
        self.set_status(key, "running")

    def complete_all(self):
        for key in self.order:
            self.set_status(key, "done")

    def mark_error(self, key):
        """Bước `key` bị lỗi — các bước trước đó vẫn coi là đã xong (đúng thực tế: chúng đã
        chạy thành công trước khi mắc ở `key`)."""
        if key not in self._rows:
            return
        idx = self.order.index(key)
        for k in self.order[:idx]:
            self.set_status(k, "done")
        self.set_status(key, "error")


class PreviewWindow(tk.Toplevel):
    """Xem lại thông tin + file đã chuẩn bị trước khi gửi thật.
    - Trái: thông tin đã lưu + danh sách file (2 nhóm: Phiếu trình / Văn bản).
    - Phải: xem PDF; với file CHÍNH của mỗi nhóm, các vị trí đánh số chữ ký (tính theo
      Luồng trình) hiện thành các dấu tròn có thể kéo/sửa/xoá/thêm ngay trên trang.
    - File không phải PDF: bấm vào chỉ mở thư mục chứa file (Explorer)."""

    MARK_R = 11
    BANNER_STYLE = {   # kind -> (nền, chữ)
        "success": ("#e8f5e9", "#1b5e20"),
        "ambiguous": ("#fff8e1", "#7a5b00"),
        "error": ("#fdecea", "#b71c1c"),
    }

    def __init__(self, master, cfg, session):
        super().__init__(master)
        self.cfg = cfg
        self.session = session
        self._stamps_cache = {}     # path -> [stamp dict]
        self._current_doc = None
        self._current_path = None
        self._current_page = 0
        self._current_kind = None
        self._editable = False
        self._add_mode = False
        self._zoom = 1.0
        self._mark_items = {}       # idx -> (oval_id, text_id)
        self._drag_idx = None
        self._drag_last = None
        self._tree_paths = {}       # iid -> (path, kind)
        self._sending = False       # True trong lúc luồng nền đang Lưu/Trình — chặn đóng cửa
                                     # sổ "lặng lẽ" giữa chừng (xem _on_close)
        self._current_phase = None  # phase (xem PIPELINE_PHASES) đang chạy — để biết tô đỏ
                                     # đúng bước nào nếu lỗi

        self.title("Xem trước & Gửi")
        self.geometry("1180x760")
        self.minsize(820, 520)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # ---- Thanh trên: nút Lưu dự thảo / Trình văn bản ----
        top = ttk.Frame(self, padding=(10, 8)); top.pack(fill="x")
        self.btn_send = ttk.Button(top, text="Lưu dự thảo", command=lambda: self._send("0"))
        self.btn_send.pack(side="left")
        self.btn_submit = ttk.Button(top, text="Trình văn bản", command=lambda: self._send("1"))
        self.btn_submit.pack(side="left", padx=(6, 0))
        ttk.Button(top, text="Đóng", command=self._on_close).pack(side="left", padx=6)
        if not cfg.get("auto_stamp", True):
            ttk.Label(top, foreground="#c62828",
                      text="⚠ Đang TẮT 'Tự đánh số chữ ký' — các dấu bên dưới sẽ KHÔNG được ghi khi Gửi.")\
                .pack(side="left", padx=12)
        self.status_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.status_var, foreground="gray").pack(side="right")

        # ---- Khung tiến trình: checklist các mốc lớn + banner kết quả cuối ----
        # Mục đích: người dùng phổ thông bấm Lưu/Trình xong phải THẤY NGAY (không cần đọc log
        # kỹ thuật) đang chạy tới đâu, và cuối cùng có chắc chắn thành công hay không.
        progress = ttk.Frame(self, padding=(10, 4)); progress.pack(fill="x")
        self.checklist = StepChecklist(progress, PIPELINE_PHASES)
        self.checklist.pack(fill="x", anchor="w")
        self.banner_var = tk.StringVar(value="")
        self.banner_label = tk.Label(progress, textvariable=self.banner_var, justify="left",
                                      anchor="w", padx=10, pady=8, wraplength=1000, font=("", 10))
        # banner ẩn cho tới khi có kết quả (xem _set_banner) — không choán chỗ lúc mới mở khung
        btnrow = ttk.Frame(progress); btnrow.pack(fill="x", pady=(4, 0))
        self.btn_toggle_log = ttk.Button(btnrow, text="▾ Xem chi tiết kỹ thuật",
                                          command=self._toggle_log)
        self.btn_toggle_log.pack(side="left")

        body = ttk.Frame(self); body.pack(fill="both", expand=True)
        left = ttk.Frame(body, width=340); left.pack(side="left", fill="y")
        left.pack_propagate(False)
        right = ttk.Frame(body, padding=(0, 8, 8, 8)); right.pack(side="left", fill="both", expand=True)

        # Chia cột trái thành 2 khung có thể kéo giãn: "Thông tin đã lưu" (tự cuộn riêng nếu
        # dài) ở trên, "File" ở dưới — để danh sách file LUÔN có chỗ hiện ra, không bị đẩy
        # khuất khi phần thông tin (trích yếu dài, nhiều nơi nhận...) chiếm hết chỗ.
        vpaned = ttk.PanedWindow(left, orient="vertical")
        vpaned.pack(fill="both", expand=True, padx=8, pady=8)
        info_pane = ttk.Frame(vpaned)
        tree_pane = ttk.Frame(vpaned)
        vpaned.add(info_pane, weight=1)
        vpaned.add(tree_pane, weight=2)

        def _set_initial_sash():
            try:
                total = vpaned.winfo_height()
                if total > 50:
                    vpaned.sashpos(0, min(280, int(total * 0.4)))
            except tk.TclError:
                pass
        self.after(50, _set_initial_sash)

        self._build_info(info_pane)
        self._build_tree(tree_pane, cfg)
        self._build_viewer(right)

        # Log gửi kỹ thuật — ẩn mặc định (xem progress/btn_toggle_log ở trên), người phổ thông
        # không cần thấy; bấm "Xem chi tiết kỹ thuật" mới hiện.
        self.logbox = tk.Text(self, height=6, bg="#101418", fg="#d0d0d0")

        # Quét sẵn file chính của phiếu trình + MỌI văn bản (không chờ người dùng bấm) + chọn
        # file mặc định để hiện (ưu tiên văn bản đầu tiên có file, không thì phiếu trình).
        report_main = cfg.get("file_report_main")
        if report_main and report_main.lower().endswith(".pdf"):
            self._ensure_scanned(report_main)
        draft_mains = [d.get("file_draft_main") for d in (cfg.get("documents") or [])]
        for dm in draft_mains:
            if dm and dm.lower().endswith(".pdf"):
                self._ensure_scanned(dm)

        init_path = next((dm for dm in draft_mains if dm), None) or report_main
        for iid, (p, k) in self._tree_paths.items():
            if p == init_path:
                self.tree.selection_set(iid); self.tree.see(iid)
                break

    # ---------- Trái: thông tin + cây file ----------
    def _build_info(self, parent):
        parent = make_scrollable_frame(parent)   # tự cuộn riêng khi trích yếu/nơi nhận dài
        info = ttk.LabelFrame(parent, text="Thông tin đã lưu", padding=8)
        info.pack(fill="x", pady=(0, 8))
        rows = [
            ("Độ khẩn", self.cfg.get("priority")),
            ("Độ mật", self.cfg.get("security")),
            ("Nội dung phiếu", self.cfg.get("report_content")),
            ("Luồng trình", self.cfg.get("flow_name")),
            ("Hồ sơ công việc", self.cfg.get("work_profile_name")),
        ]
        for label, val in rows:
            r = ttk.Frame(info); r.pack(fill="x", anchor="w", pady=1)
            ttk.Label(r, text=label + ":", width=13, foreground="gray").pack(side="left", anchor="n")
            ttk.Label(r, text=val or "(trống)", wraplength=200, justify="left").pack(
                side="left", fill="x", expand=True)

        documents = self.cfg.get("documents") or []
        if documents:
            docs_frame = ttk.LabelFrame(parent, text=f"Văn bản ({len(documents)})", padding=8)
            docs_frame.pack(fill="x", pady=(0, 8))
            for i, doc in enumerate(documents):
                r = ttk.Frame(docs_frame); r.pack(fill="x", anchor="w", pady=(0 if i == 0 else 6, 0))
                summary = f"{doc.get('doc_type') or '(chưa chọn loại)'} — {doc.get('code') or '(chưa có số)'}"
                ttk.Label(r, text=f"VB {i+1}:", width=13, foreground="gray").pack(side="left", anchor="n")
                ttk.Label(r, text=summary, wraplength=200, justify="left").pack(
                    side="left", fill="x", expand=True)
                if doc.get("abstract"):
                    r2 = ttk.Frame(docs_frame); r2.pack(fill="x", anchor="w")
                    ttk.Label(r2, text="", width=13).pack(side="left")
                    ttk.Label(r2, text=doc["abstract"], wraplength=200, justify="left",
                              foreground="#444").pack(side="left", fill="x", expand=True)

        cats = [("recv_inside", "Nhận nội bộ"), ("recv_report", "Báo cáo"), ("recv_edoc", "Liên thông"),
                ("recv_save", "Nơi lưu"), ("recv_know", "Để biết")]
        any_recv = any(self.cfg.get(k) for k, _ in cats)
        if any_recv:
            rec = ttk.LabelFrame(parent, text="Nơi nhận", padding=8)
            rec.pack(fill="x", pady=(0, 8))
            for key, label in cats:
                nodes = self.cfg.get(key) or []
                if not nodes:
                    continue
                names = ", ".join(nd["name"].strip() for nd in nodes)
                r = ttk.Frame(rec); r.pack(fill="x", anchor="w", pady=1)
                ttk.Label(r, text=label + ":", width=13, foreground="gray").pack(side="left", anchor="n")
                ttk.Label(r, text=names, wraplength=200, justify="left").pack(
                    side="left", fill="x", expand=True)

    def _build_tree(self, parent, cfg):
        box = ttk.LabelFrame(parent, text="File", padding=4)
        box.pack(fill="both", expand=True)
        tree_wrap = ttk.Frame(box); tree_wrap.pack(fill="both", expand=True)
        tsb = ttk.Scrollbar(tree_wrap, orient="vertical")
        self.tree = ttk.Treeview(tree_wrap, show="tree", height=14, yscrollcommand=tsb.set)
        tsb.config(command=self.tree.yview)
        tsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        bind_mousewheel_scroll(self.tree)

        def add_group(title, entries):
            gid = self.tree.insert("", "end", text=title, open=True)
            for path, kind in entries:
                if not path:
                    continue
                iid = self.tree.insert(gid, "end", text=os.path.basename(path))
                self._tree_paths[iid] = (path, kind)

        add_group("📄 PHIẾU TRÌNH (không gửi đi)",
                   [(cfg.get("file_report_main"), "report_main")] +
                   [(p, "report_extra") for p in (cfg.get("files_report_extra") or [])])
        documents = cfg.get("documents") or []
        n_docs = len(documents)
        for i, doc in enumerate(documents):
            title = f"📤 VĂN BẢN {i+1}/{n_docs} (gửi đi)" if n_docs > 1 else "📤 VĂN BẢN (gửi đi)"
            add_group(title,
                      [(doc.get("file_draft_main"), f"draft_main:{i}")] +
                      [(p, f"draft_extra:{i}") for p in (doc.get("files_draft_extra") or [])])

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        ttk.Label(parent, foreground="gray", wraplength=320, justify="left",
                  text="PDF của Phiếu trình/Dự thảo: kéo dấu để di chuyển, nhấp đúp để sửa số/xoá, "
                       "bấm '➕ Thêm dấu' rồi nhấp vào trang để thêm. File khác/PDF phụ: bấm để mở "
                       "thư mục chứa file.").pack(anchor="w", pady=(6, 0))

    def _on_tree_select(self, _e):
        sel = self.tree.selection()
        if not sel or sel[0] not in self._tree_paths:
            return
        path, kind = self._tree_paths[sel[0]]
        self._select_file(path, kind)

    # ---------- Phải: khung xem PDF ----------
    def _build_viewer(self, parent):
        bar = ttk.Frame(parent); bar.pack(fill="x")
        ttk.Button(bar, text="◀", width=3, command=self._prev_page).pack(side="left")
        self.page_var = tk.StringVar(value="—")
        ttk.Label(bar, textvariable=self.page_var).pack(side="left", padx=4)
        ttk.Button(bar, text="▶", width=3, command=self._next_page).pack(side="left")
        self.btn_addmark = ttk.Button(bar, text="➕ Thêm dấu", command=self._toggle_add_mode, state="disabled")
        self.btn_addmark.pack(side="left", padx=12)

        wrap = ttk.Frame(parent); wrap.pack(fill="both", expand=True, pady=(6, 0))
        xsb = ttk.Scrollbar(wrap, orient="horizontal")
        ysb = ttk.Scrollbar(wrap, orient="vertical")
        self.canvas = tk.Canvas(wrap, bg="#5c5c5c",
                                 xscrollcommand=xsb.set, yscrollcommand=ysb.set)
        xsb.config(command=self.canvas.xview); ysb.config(command=self.canvas.yview)
        ysb.pack(side="right", fill="y"); xsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Button-1>", self._canvas_click)
        bind_mousewheel_scroll(self.canvas, horizontal=True)

        self._page_img = None   # giữ tham chiếu tránh bị garbage-collect

    # ---------- Chọn file trong cây ----------
    def _select_file(self, path, kind):
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            self._open_folder(path)
            return
        try:
            doc = fitz.open(path)
        except Exception as e:
            messagebox.showerror("Lỗi mở PDF", f"Không mở được file:\n{e}", parent=self)
            return
        if self._current_doc is not None:
            try: self._current_doc.close()
            except Exception: pass
        self._current_doc = doc
        self._current_path = path
        self._current_kind = kind
        self._current_page = 0
        is_main = kind == "report_main" or kind.startswith("draft_main")
        self._editable = is_main and self.cfg.get("auto_stamp", True)
        self._add_mode = False
        self.btn_addmark.config(state=("normal" if self._editable else "disabled"),
                                 text="➕ Thêm dấu")
        self._render_current()
        if is_main:
            self._ensure_scanned(path)
        else:
            self._draw_marks()

    def _open_folder(self, path):
        try:
            if sys.platform == "win32":
                subprocess.Popen(["explorer", f"/select,{os.path.normpath(path)}"])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", path])   # hiện trong Finder, tương đương /select trên Windows
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(path)])
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không mở được thư mục:\n{e}", parent=self)

    # ---------- Vẽ trang ----------
    def _render_current(self):
        if self._current_doc is None:
            return
        page = self._current_doc[self._current_page]
        target_w = max(self.canvas.winfo_width(), 640)
        self._zoom = max(0.5, min(2.2, target_w / max(page.rect.width, 1)))
        pix = page.get_pixmap(matrix=fitz.Matrix(self._zoom, self._zoom))
        self._page_img = tk.PhotoImage(data=pix.tobytes("ppm"))
        self.canvas.delete("page")
        self.canvas.create_image(0, 0, anchor="nw", image=self._page_img, tags="page")
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self.page_var.set(f"Trang {self._current_page + 1}/{len(self._current_doc)}")
        self._draw_marks()

    def _prev_page(self):
        if self._current_doc and self._current_page > 0:
            self._current_page -= 1
            self._render_current()

    def _next_page(self):
        if self._current_doc and self._current_page < len(self._current_doc) - 1:
            self._current_page += 1
            self._render_current()

    # ---------- Quét vị trí ký (nền) ----------
    def _ensure_scanned(self, path):
        if path in self._stamps_cache:
            if path == self._current_path:
                self._draw_marks()
            return
        self.status_var.set(f"Đang quét vị trí ký: {os.path.basename(path)}…")

        def worker():
            try:
                d = fitz.open(path)
                try:
                    stamps = find_signature_stamps(d, self.cfg.get("flow_nodes_override"), self._plog)
                finally:
                    d.close()
            except Exception as e:
                err_msg = str(e)   # tính ngay trong khối except — "e" bị Python xoá khi except kết thúc
                self.after(0, lambda: self._scan_failed(path, err_msg))
                return
            self.after(0, lambda: self._scan_done(path, stamps))
        threading.Thread(target=worker, daemon=True).start()

    def _scan_done(self, path, stamps):
        self._stamps_cache[path] = stamps
        self.status_var.set("")
        if path == self._current_path:
            self._draw_marks()

    def _scan_failed(self, path, e):
        self.status_var.set("")
        self._plog(f"✖ Không quét được vị trí ký trong {os.path.basename(path)}: {e}")

    # ---------- Vẽ + tương tác các dấu ----------
    def _draw_marks(self):
        self.canvas.delete("markgrp")
        self._mark_items = {}
        if not self._editable or self._current_path not in self._stamps_cache:
            return
        stamps = self._stamps_cache[self._current_path]
        r = self.MARK_R
        for i, st in enumerate(stamps):
            if st["page"] != self._current_page:
                continue
            cx, cy = st["x"] * self._zoom, st["y"] * self._zoom
            oval = self.canvas.create_oval(cx - r, cy - r, cx + r, cy + r,
                                            fill="#ffca28", outline="#e65100", width=2)
            txt = self.canvas.create_text(cx, cy, text=str(st["number"]),
                                           fill="#000", font=("", 10, "bold"))
            tag = f"mark{i}"
            self.canvas.itemconfig(oval, tags=("markgrp", tag))
            self.canvas.itemconfig(txt, tags=("markgrp", tag))
            self._mark_items[i] = (oval, txt)
            self.canvas.tag_bind(tag, "<ButtonPress-1>", lambda e, idx=i: self._mark_press(e, idx))
            self.canvas.tag_bind(tag, "<B1-Motion>", lambda e, idx=i: self._mark_drag(e, idx))
            self.canvas.tag_bind(tag, "<ButtonRelease-1>", lambda e, idx=i: self._mark_release(e, idx))
            self.canvas.tag_bind(tag, "<Double-Button-1>", lambda e, idx=i: self._mark_edit(idx))

    def _mark_press(self, e, idx):
        self._drag_idx = idx
        self._drag_last = (e.x, e.y)

    def _mark_drag(self, e, idx):
        if self._drag_idx != idx:
            return
        dx, dy = e.x - self._drag_last[0], e.y - self._drag_last[1]
        oval, txt = self._mark_items[idx]
        self.canvas.move(oval, dx, dy)
        self.canvas.move(txt, dx, dy)
        self._drag_last = (e.x, e.y)

    def _mark_release(self, e, idx):
        if self._drag_idx != idx:
            return
        self._drag_idx = None
        oval, _txt = self._mark_items[idx]
        x0, y0, x1, y1 = self.canvas.coords(oval)
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        st = self._stamps_cache[self._current_path][idx]
        st["x"], st["y"] = cx / self._zoom, cy / self._zoom

    def _mark_edit(self, idx):
        self._open_role_dialog(idx=idx)

    # ---------- Thêm dấu mới ----------
    def _toggle_add_mode(self):
        self._add_mode = not self._add_mode
        self.btn_addmark.config(text="✋ Nhấp vào trang…" if self._add_mode else "➕ Thêm dấu")

    def _canvas_click(self, e):
        if not self._add_mode:
            return
        if "markgrp" in self.canvas.gettags("current"):
            return   # nhấp trúng 1 dấu có sẵn — để binding của dấu đó xử lý
        self._add_mode = False
        self.btn_addmark.config(text="➕ Thêm dấu")
        px, py = self.canvas.canvasx(e.x) / self._zoom, self.canvas.canvasy(e.y) / self._zoom
        self._open_role_dialog(new=True, page=self._current_page, x=px, y=py)

    # ---------- Hộp thoại chọn chức danh/số ----------
    def _open_role_dialog(self, idx=None, new=False, page=None, x=None, y=None):
        stamps = self._stamps_cache.setdefault(self._current_path, [])
        st = None if new else stamps[idx]

        role_map = build_role_number_map(self.cfg.get("flow_nodes_override"))
        role_keys = list(role_map.keys()) or [SIG_CHUYEN_VIEN_ROLE]

        dlg = tk.Toplevel(self)
        dlg.title("Thêm dấu ký" if new else "Sửa dấu ký")
        dlg.transient(self); dlg.grab_set()
        ttk.Label(dlg, text="Chức danh:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        role_var = tk.StringVar(value=(st["role"] if st and st.get("role") in role_keys else role_keys[0]))
        combo = ttk.Combobox(dlg, values=role_keys, textvariable=role_var, state="readonly", width=22)
        combo.grid(row=0, column=1, padx=8, pady=6)

        ttk.Label(dlg, text="Số:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        num_var = tk.StringVar(value=(str(st["number"]) if st else ""))
        ttk.Entry(dlg, textvariable=num_var, width=8).grid(row=1, column=1, sticky="w", padx=8, pady=6)

        def _fill_number(*_):
            n = role_map.get(role_var.get())
            num_var.set(str(n) if n is not None else "")
        combo.bind("<<ComboboxSelected>>", _fill_number)
        if new:
            _fill_number()

        btns = ttk.Frame(dlg); btns.grid(row=2, column=0, columnspan=2, pady=8)

        def _close_dlg():
            # Giải phóng "grab" TƯỜNG MINH trước khi đóng — trên macOS, destroy() không phải
            # lúc nào cũng tự giải phóng grab sạch sẽ, để lại cả ứng dụng bị khoá input (mọi ô
            # nhập khác, kể cả ở màn hình chính, ngừng nhận phím) dù cửa sổ đã đóng từ lâu.
            try:
                dlg.grab_release()
            except Exception:
                pass
            dlg.destroy()

        def _ok():
            try:
                number = int(num_var.get())
            except ValueError:
                messagebox.showwarning("Số không hợp lệ", "Số phải là số nguyên.", parent=dlg)
                return
            role = role_var.get()
            if new:
                stamps.append({"page": page, "x": x, "y": y, "number": number, "role": role,
                                "title": f"{role} (thêm tay)"})
            else:
                st["number"] = number
                st["role"] = role
            _close_dlg()
            self._draw_marks()
        dlg.protocol("WM_DELETE_WINDOW", _close_dlg)   # bấm nút [X] của cửa sổ cũng phải giải phóng grab
        ttk.Button(btns, text="OK", command=_ok).pack(side="left", padx=4)
        if not new:
            def _del():
                del stamps[idx]
                _close_dlg()
                self._draw_marks()
            ttk.Button(btns, text="Xóa", command=_del).pack(side="left", padx=4)
        ttk.Button(btns, text="Hủy", command=_close_dlg).pack(side="left", padx=4)

    # ---------- Tiến trình + kết quả (checklist/banner thân thiện — xem StepChecklist) ----------
    def _toggle_log(self):
        if self.logbox.winfo_ismapped():
            self.logbox.pack_forget()
            self.btn_toggle_log.config(text="▾ Xem chi tiết kỹ thuật")
        else:
            self.logbox.pack(fill="x", side="bottom", padx=10, pady=(0, 10))
            self.btn_toggle_log.config(text="▴ Ẩn chi tiết kỹ thuật")

    def _set_banner(self, kind, lines):
        bg, fg = self.BANNER_STYLE[kind]
        self.banner_var.set("\n".join(lines))
        self.banner_label.config(bg=bg, fg=fg)
        self.banner_label.pack(fill="x", pady=(6, 0))

    def _set_phase(self, key):
        self._current_phase = key
        self.checklist.mark_running(key)

    def _render_result_banner(self, result):
        """`result`: dict trả về từ run_pipeline (xem verify_report_saved) — KHÔNG suy đoán từ
        mã HTTP, chỉ hiện những gì đã tự đọc lại và xác nhận được với server."""
        sign = result.get("submit_sign")
        if not result.get("verified"):
            self._set_banner("ambiguous", [
                "⚠ ĐÃ GỬI YÊU CẦU, NHƯNG CHƯA XÁC MINH ĐƯỢC",
                "Hệ thống chưa phản hồi rõ trong lúc kiểm tra lại — nhiều khả năng vẫn đã thành công.",
                "→ Mở web, vào thùng nháp/đang trình để tự kiểm tra trước khi thử gửi lại (tránh tạo trùng).",
            ])
            return
        if sign != "1":
            self._set_banner("success", [
                "✅ ĐÃ LƯU NHÁP THÀNH CÔNG",
                "Hệ thống đã xác nhận phiếu trình có tồn tại (chưa gửi cho ai).",
                "→ Mở web, vào thùng nháp để trình khi sẵn sàng.",
            ])
            return
        if not result.get("in_process"):
            # Thấy phiếu trình tồn tại (khớp content+thời gian) nhưng KHÔNG thấy trong đúng hộp
            # "đang trình" — với sign=1 lẽ ra phải thấy ở cả 2 nơi. Đáng ngờ riêng, không nên
            # gộp chung với "✅ thành công" — có thể lệnh Trình chưa thật sự vào luồng.
            self._set_banner("ambiguous", [
                "⚠ ĐÃ LƯU, NHƯNG CHƯA THẤY TRONG HỘP \"ĐANG TRÌNH\"",
                "Hệ thống xác nhận phiếu trình tồn tại, nhưng chưa thấy nó ở danh sách đang xử lý "
                "— có thể chỉ là chậm cập nhật, cũng có thể lệnh Trình chưa thật sự vào luồng.",
                "→ Mở web, vào hộp \"đang trình\" kiểm tra lại cho chắc trước khi coi là xong.",
            ])
            return
        line2 = "Hệ thống đã xác nhận phiếu trình đã vào luồng ký duyệt"
        note = result.get("history_note")
        if note and note.get("createAt"):
            line2 += f", lúc {note['createAt'][11:16]}"
        line2 += "."
        lines = ["✅ ĐÃ TRÌNH THÀNH CÔNG", line2]
        pending = result.get("pending")
        if pending:
            who = pending.get("receiveUser") or "?"
            role = pending.get("displayPositionName") or ""
            lines.append(f"Đang chờ: {who}" + (f" ({role})" if role else "") + " xử lý.")
        self._set_banner("success", lines)

    def _render_error_banner(self, err_text):
        phase_label = dict(PIPELINE_PHASES).get(self._current_phase, self._current_phase or "?")
        if self._current_phase:
            self.checklist.mark_error(self._current_phase)
        self._set_banner("error", [
            "❌ CHƯA HOÀN TẤT",
            f"Mắc ở bước: {phase_label}",
            f"Lý do: {err_text[:300]}",
            "→ Xem \"chi tiết kỹ thuật\" bên dưới để rõ hơn. Nếu bước \"Lưu văn bản\" đã xong "
            "trước khi lỗi, kiểm tra kỹ thùng nháp trên web trước khi thử lại, tránh tạo trùng.",
        ])

    # ---------- Gửi thật ----------
    def _plog(self, msg):
        self.logbox.insert("end", msg + "\n"); self.logbox.see("end"); self.update_idletasks()

    def _send(self, sign):
        cfg = dict(self.cfg)
        cfg["submit_sign"] = sign   # "0" = Lưu dự thảo, "1" = Trình văn bản (xem run_pipeline)
        cfg["stamps_report_override"] = self._stamps_cache.get(cfg.get("file_report_main"))
        docs = []
        for doc in (self.cfg.get("documents") or []):
            d = dict(doc)
            main = d.get("file_draft_main")
            if main:
                d["stamps_override"] = self._stamps_cache.get(main)
            docs.append(d)
        cfg["documents"] = docs
        self.btn_send.config(state="disabled")
        self.btn_submit.config(state="disabled")
        self._sending = True
        self._current_phase = None
        self.checklist.reset()
        self.banner_label.pack_forget()
        self.logbox.delete("1.0", "end")
        self._plog("=== BẮT ĐẦU TRÌNH ===" if sign == "1" else "=== BẮT ĐẦU LƯU DỰ THẢO ===")

        def phase_cb(key):
            self.after(0, lambda k=key: self._set_phase(k))

        def worker():
            try:
                result = run_pipeline(self.session, cfg, self._plog, check_only=False, phase_cb=phase_cb)
            except PipelineError as e:
                err_text = str(e)   # tính ngay trong khối except — "e" bị Python xoá khi except kết thúc
                self._plog("\n✖ DỪNG: " + err_text)
                self.after(0, lambda: self._render_error_banner(err_text))
            except Exception as e:
                err_text = repr(e)
                self._plog("\n✖ LỖI KHÔNG NGỜ: " + err_text)
                self.after(0, lambda: self._render_error_banner(err_text))
            else:
                self.after(0, lambda: (self.checklist.complete_all(), self._render_result_banner(result)))
            finally:
                self._sending = False
                self.after(0, lambda: self.btn_send.config(state="normal"))
                self.after(0, lambda: self.btn_submit.config(state="normal"))
        threading.Thread(target=worker, daemon=True).start()

    def _on_close(self):
        if self._sending:
            if not messagebox.askyesno(
                    "Đang gửi",
                    "Đang gửi phiếu trình — đóng cửa sổ này thì việc gửi vẫn tiếp tục chạy ngầm, "
                    "nhưng bạn sẽ không thấy kết quả nữa. Vẫn đóng?", parent=self):
                return
        if self._current_doc is not None:
            try: self._current_doc.close()
            except Exception: pass
        master = self.master
        self.destroy()
        # Trên macOS, đóng 1 Toplevel không tự trả bàn phím về đúng cửa sổ cha — màn hình chính
        # nhìn vẫn "active" nhưng không nhận phím nữa (vd ô "Nơi nhận" gõ không ăn) cho tới khi
        # tự bấm vào. Ép lại tường minh để khỏi phải bấm.
        try:
            master.lift()
            master.focus_force()
        except Exception:
            pass


if __name__ == "__main__":
    App().mainloop()
