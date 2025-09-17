import io, os, uuid, zipfile, tempfile, shutil, sys
from datetime import datetime
import streamlit as st
import subprocess

APP_DIR = os.path.dirname(__file__)

def find_export_script() -> str | None:
    candidates = [
        os.path.join(APP_DIR, "exportword.py"),
        os.path.join(APP_DIR, "data", "exportword.py"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

st.set_page_config(page_title="Đồng bộ bản ảnh - Web", layout="centered")
st.title("🛠 Đồng bộ bản ảnh (gọi exportword.py)")

# --- UI ---
col1, col2 = st.columns(2)
loaibananh = col1.selectbox("Loại bản ảnh", [
    "khám nghiệm hiện trường",
    "khám nghiệm tử thi",
    "khám nghiệm hiện trường và tử thi",
    "khám phương tiện liên quan đến tai nạn giao thông",
    "Tùy chỉnh"
], key="loaibananh_sel")
loaibananh_custom = ""
if loaibananh == "Tùy chỉnh":
    loaibananh_custom = col2.text_input("Nhập loại bản ảnh tùy chỉnh", "", key="loaibananh_custom")

vu_viec = st.selectbox("Tên vụ việc", [
    "Chết người", "Giết người", "Cố ý gây thương tích",
    "Tai nạn giao thông đường bộ", "Hủy hoại tài sản",
    "Trộm cắp tài sản", "Hiếp dâm", "Tùy chỉnh"
], key="vu_viec_sel")
vu_viec_custom = ""
if vu_viec == "Tùy chỉnh":
    vu_viec_custom = st.text_input("Nhập tên vụ việc tùy chỉnh", "", key="vu_viec_custom")

xrph = st.radio("Sự kiện", ["xảy ra", "phát hiện"], horizontal=True, key="xrph_radio")
col3, col4 = st.columns(2)
nkn = col3.date_input("Ngày KN", format="DD/MM/YYYY", key="nkn_date")
nxr = col4.date_input("Ngày xảy ra", format="DD/MM/YYYY", key="nxr_date")
diadiem = st.text_input("Địa điểm xảy ra", "", key="dd_input")

st.divider()
st.subheader("Tệp cần tải lên")
up_tmba = st.file_uploader("TMBA.xlsm", type=["xlsm"], key="tmba_upl")
up_zip = st.file_uploader("Ảnh kèm theo (ZIP)", type=["zip"], key="zip_upl",
                          help="Nén toàn bộ ảnh vào 1 file .zip rồi tải lên")

run_btn = st.button("🚀 Đồng bộ & Xuất file", type="primary", use_container_width=True, key="run_btn")

# --- RUN ---
if run_btn:
    # 0) Kiểm tra script
    script_path = find_export_script()
    if not script_path:
        st.error("Không tìm thấy `exportword.py`. Hãy đặt file này ở **gốc repo** hoặc trong thư mục **data/**.")
        st.stop()

    # 1) Kiểm tra đầu vào
    if up_tmba is None:
        st.error("Vui lòng tải lên **TMBA.xlsm**.")
        st.stop()
    if up_zip is None:
        st.error("Vui lòng tải lên **ZIP ảnh**.")
        st.stop()

    # 2) Tạo thư mục làm việc tạm
    workdir = tempfile.mkdtemp(prefix="dongbo_", dir=None)  # mặc định dùng /tmp trên server
    imgdir = os.path.join(workdir, "images")
    os.makedirs(imgdir, exist_ok=True)

    # 3) Lưu TMBA.xlsm
    tmba_path = os.path.join(workdir, "TMBA.xlsm")
    with open(tmba_path, "wb") as f:
        f.write(up_tmba.read())

    # 4) Giải nén ZIP ảnh
    try:
        with zipfile.ZipFile(io.BytesIO(up_zip.read())) as zf:
            zf.extractall(imgdir)
    except Exception as e:
        st.error(f"Không giải nén được ZIP ảnh: {e}")
        shutil.rmtree(workdir, ignore_errors=True)
        st.stop()

    # 5) Chuẩn bị biến môi trường cho exportword.py
    env = os.environ.copy()
    env["EXPORT_FOLDER"] = workdir
    env["IMAGE_FOLDER"] = imgdir
    env["LOAI_BAN_ANH"] = (loaibananh_custom if loaibananh == "Tùy chỉnh" else loaibananh)
    env["vuviec"] = (vu_viec_custom if vu_viec == "Tùy chỉnh" else vu_viec)
    env["XRPH"] = xrph
    env["diadiem"] = diadiem
    env["NXR"] = nxr.strftime("%d/%m/%Y")
    env["NKN"] = nkn.strftime("%d/%m/%Y")

    st.info("Đang chạy exportword.py — vui lòng đợi…")
    with st.status("Đang xử lý…", expanded=True) as status:
        st.write(f"📁 Thư mục làm việc: `{workdir}`")
        st.write(f"📂 Ảnh bung tại: `{imgdir}`")
        try:
            # 6) Chạy script
            result = subprocess.run(
                [sys.executable, script_path],
                cwd=workdir, env=env,
                capture_output=True, text=True, timeout=600
            )
            st.code(result.stdout or "(không có stdout)", language="bash")
            if result.returncode != 0:
                st.error("exportword.py trả về lỗi:")
                st.code(result.stderr or "(không có stderr)", language="bash")
                status.update(label="❌ Thất bại", state="error")
                # KHÔNG xóa workdir để bạn có thể debug (tuỳ bạn giữ/xoá)
                st.stop()
        except subprocess.TimeoutExpired:
            status.update(label="⏱️ Hết thời gian chờ", state="error")
            st.error("Quá thời gian xử lý. Thử giảm số ảnh hoặc tối ưu script.")
            st.stop()
        except Exception as e:
            status.update(label="❌ Lỗi khi chạy script", state="error")
            st.exception(e)
            st.stop()

        # 7) Thu thập output (docx/pdf/zip…) từ workdir
        outputs = []
        for root, _, files in os.walk(workdir):
            for name in files:
                if name.lower().endswith((".docx", ".pdf", ".zip")):
                    outputs.append(os.path.join(root, name))

        if not outputs:
            status.update(label="⚠️ Không tìm thấy file xuất", state="warning")
            st.warning("Không thấy file đầu ra trong thư mục làm việc. Kiểm tra lại logic trong `exportword.py`.")
        else:
            status.update(label="✅ Hoàn tất", state="complete")
            st.success("Đồng bộ thành công. Tải file bên dưới:")

            # Sắp xếp theo thời gian mới nhất trước
            outputs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            for p in outputs:
                with open(p, "rb") as f:
                    st.download_button(
                        f"⬇️ Tải: {os.path.basename(p)}",
                        data=f.read(),
                        file_name=os.path.basename(p),
                        key=f"dl_{uuid.uuid4().hex}"
                    )

    # 8) (Tuỳ chọn) dọn tạm: comment dòng dưới nếu bạn muốn giữ lại để debug
    # shutil.rmtree(workdir, ignore_errors=True)

st.caption("© PC09 Khánh Hòa • Bản web chỉ đồng bộ bằng exportword.py (không tạo Thuyết minh)")
