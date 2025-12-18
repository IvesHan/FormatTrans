import streamlit as st
import pandas as pd
import io
import zipfile
from PIL import Image
from pypdf import PdfWriter, PdfReader
from pdf2image import convert_from_bytes
import pikepdf
import docx
import pytesseract

# ==========================================
# 页面基础配置
# ==========================================
st.set_page_config(page_title="全能文件处理站 Pro", page_icon="🛠️", layout="wide")

st.title("🛠️ 全能文件处理站 Pro")
st.markdown("""
**功能概览**：
* **📊 表格**：支持 CSV (中/英/法格式)、Excel、JSON 格式互转。
* **📄 文档**：PDF 排序合并、PDF 转高清图、**OCR 文字识别 (支持扫描件)**。
* **🖼️ 图片**：格式互转、修改 DPI、多图拼合转 PDF。
""")

# ==========================================
# 辅助函数定义
# ==========================================

def try_unlock_pdf(file_obj):
    """尝试去除PDF权限限制"""
    try:
        pdf = pikepdf.open(file_obj)
        new_pdf_bytes = io.BytesIO()
        pdf.save(new_pdf_bytes)
        return new_pdf_bytes
    except pikepdf.PasswordError:
        st.error("❌ 此文件设置了【打开密码】，无法强制破除。")
        return None
    except Exception as e:
        st.error(f"❌ 权限处理失败: {e}")
        return None

def convert_df(df, fmt, sep=','):
    """表格导出转换"""
    buffer = io.BytesIO()
    if fmt == "CSV":
        buffer.write(df.to_csv(index=False, sep=sep).encode('utf-8-sig'))
        return buffer, "text/csv", "csv"
    elif fmt == "Excel":
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        return buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"
    elif fmt == "JSON":
        buffer.write(df.to_json(orient='records', force_ascii=False).encode('utf-8'))
        return buffer, "application/json", "json"

# ==========================================
# 侧边栏导航
# ==========================================
category = st.sidebar.selectbox(
    "1️⃣ 选择文件大类",
    ["📊 表格数据 (CSV/Excel)", "📄 文档工具 (PDF/Word)", "🖼️ 图片处理 (Image)"]
)

st.sidebar.markdown("---")

# =========================================================
# 模块 A: 表格数据 (已精简：仅保留转换)
# =========================================================
if category == "📊 表格数据 (CSV/Excel)":
    st.header("表格格式转换")
    
    # CSV 读取设置
    st.markdown("##### ⚙️ CSV 读取设置")
    sep_option = st.selectbox(
        "输入文件分隔符",
        ["逗号 , (标准)", "分号 ; (欧洲)", "Tab (制表符)", "自定义"],
        index=0
    )
    separator = ","
    if "分号" in sep_option: separator = ";"
    elif "Tab" in sep_option: separator = "\t"
    elif "自定义" in sep_option:
        separator = st.text_input("输入自定义分隔符", value="|")

    f = st.file_uploader("上传表格", type=['csv', 'xlsx', 'xls', 'json'])
    
    if f:
        # 读取逻辑
        df = None
        try:
            if f.name.endswith('.csv'): df = pd.read_csv(f, sep=separator)
            elif f.name.endswith('.tsv'): df = pd.read_csv(f, sep='\t')
            elif f.name.endswith(('.xls', '.xlsx')): df = pd.read_excel(f)
            elif f.name.endswith('.json'): df = pd.read_json(f)
        except Exception as e:
            st.error(f"读取错误: {e}")

        if df is not None:
            st.write("### 数据预览 (前5行)")
            st.dataframe(df.head())
            
            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                target_fmt = st.selectbox("目标格式", ["Excel", "CSV", "JSON"])
            with col2:
                export_sep = ","
                if target_fmt == "CSV":
                    export_sep = st.selectbox("导出CSV分隔符", [",", ";", "\t"], index=0)
            
            if st.button("转换并下载"):
                data, mime, ext = convert_df(df, target_fmt, export_sep)
                st.download_button(f"下载 .{ext}", data, f.name.split('.')[0]+f".{ext}", mime)

# =========================================================
# 模块 B: 文档工具 (增强 OCR)
# =========================================================
elif category == "📄 文档工具 (PDF/Word)":
    doc_task = st.sidebar.radio("2️⃣ 选择操作", ["PDF 合并 (带排序)", "PDF 转 图片", "PDF/Word 提取文本 (OCR)", "PDF 权限解除"])

    # --- 1. PDF 合并 ---
    if doc_task == "PDF 合并 (带排序)":
        st.header("PDF 合并 (支持自定义排序)")
        files = st.file_uploader("上传多个 PDF", type=['pdf'], accept_multiple_files=True)
        
        if files:
            file_map = {f.name: f for f in files}
            df_files = pd.DataFrame({"文件名": [f.name for f in files], "排序权重": range(1, len(files)+1)})
            st.info("👇 修改下方数字调整顺序 (1最前)")
            edited_df = st.data_editor(df_files, use_container_width=True)
            
            if st.button("按顺序合并"):
                sorted_names = edited_df.sort_values(by="排序权重")["文件名"].tolist()
                merger = PdfWriter()
                try:
                    for name in sorted_names:
                        f_obj = file_map[name]
                        f_obj.seek(0)
                        try:
                            reader = PdfReader(f_obj)
                            if reader.is_encrypted:
                                f_obj.seek(0)
                                unlocked = try_unlock_pdf(f_obj)
                                if unlocked: reader = PdfReader(unlocked)
                                else: continue
                            merger.append(reader)
                        except Exception: pass
                    
                    out = io.BytesIO()
                    merger.write(out)
                    out.seek(0)
                    st.download_button("下载合并 PDF", out, "merged.pdf", "application/pdf")
                except Exception as e:
                    st.error(f"合并出错: {e}")

    # --- 2. PDF 转图片 ---
    elif doc_task == "PDF 转 图片":
        st.header("PDF 转图片")
        pdf_file = st.file_uploader("上传 PDF", type=['pdf'])
        dpi = st.number_input("清晰度 (DPI)", 72, 600, 200)
        
        if pdf_file and st.button("转换"):
            try:
                images = convert_from_bytes(pdf_file.read(), dpi=dpi)
                st.success(f"共 {len(images)} 页")
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w") as zf:
                    for i, img in enumerate(images):
                        ib = io.BytesIO()
                        img.save(ib, format="JPEG")
                        zf.writestr(f"page_{i+1:03d}.jpg", ib.getvalue())
                st.download_button("下载图片包 (ZIP)", zip_buf.getvalue(), "images.zip", "application/zip")
            except Exception as e:
                st.error(f"错误: {e}")

    # --- 3. 文本提取 (含 OCR) ---
    elif doc_task == "PDF/Word 提取文本 (OCR)":
        st.header("提取文本 (支持扫描件)")
        st.info("如果是图片生成的 PDF (无法选中文本)，请勾选下方的 **'启用 OCR'**。")
        
        f = st.file_uploader("上传文件", type=['docx', 'pdf'])
        use_ocr = st.checkbox("启用 OCR (扫描件/图片模式)", value=False, help="速度较慢，适用于图片型 PDF")
        
        if f:
            txt_output = ""
            
            # Word 处理
            if f.name.endswith('.docx'):
                doc = docx.Document(f)
                txt_output = "\n".join([p.text for p in doc.paragraphs])
            
            # PDF 处理
            elif f.name.endswith('.pdf'):
                if use_ocr:
                    # OCR 模式：PDF -> 图片 -> 文字
                    with st.spinner("正在进行 OCR 识别 (这可能需要几分钟)..."):
                        try:
                            # 1. 也是先解锁
                            f.seek(0)
                            pdf_bytes = f.read()
                            
                            # 2. 转为图片
                            images = convert_from_bytes(pdf_bytes, dpi=300) # 300 DPI 识别率较好
                            
                            # 3. 逐页识别
                            full_text = []
                            progress_bar = st.progress(0)
                            for i, img in enumerate(images):
                                # 这里的 lang='chi_sim+eng' 表示同时识别简体中文和英文
                                text = pytesseract.image_to_string(img, lang='chi_sim+eng')
                                full_text.append(f"--- Page {i+1} ---\n{text}")
                                progress_bar.progress((i + 1) / len(images))
                            
                            txt_output = "\n\n".join(full_text)
                        except Exception as e:
                            st.error(f"OCR 失败: {e} (请检查 packages.txt 是否包含 tesseract-ocr)")
                else:
                    # 普通模式：直接提取
                    reader = PdfReader(f)
                    for p in reader.pages:
                        txt_output += p.extract_text() + "\n\n"
            
            if txt_output:
                st.text_area("提取结果", txt_output, height=400)
                st.download_button("下载 .txt", txt_output, "extracted_text.txt")
            else:
                st.warning("未能提取到文本。如果是扫描件，请勾选 '启用 OCR'。")

    # --- 4. 权限解除 ---
    elif doc_task == "PDF 权限解除":
        st.header("🔒 PDF 权限移除")
        locked = st.file_uploader("上传受限 PDF", type=['pdf'])
        if locked and st.button("解锁"):
            unlocked = try_unlock_pdf(locked)
            if unlocked:
                unlocked.seek(0)
                st.success("解锁成功！")
                st.download_button("下载解锁版 PDF", unlocked, f"unlocked_{locked.name}", "application/pdf")

# =========================================================
# 模块 C: 图片处理 (保持不变)
# =========================================================
elif category == "🖼️ 图片处理 (Image)":
    img_task = st.sidebar.radio("2️⃣ 选择操作", ["格式转换 / 修改PPI", "多图拼合转PDF"])

    if img_task == "格式转换 / 修改PPI":
        st.header("图片处理")
        f = st.file_uploader("上传图片", type=['png', 'jpg', 'jpeg', 'bmp', 'tiff', 'webp'])
        if f:
            img = Image.open(f)
            st.image(img, caption=f"尺寸: {img.size}", width=300)
            c1, c2 = st.columns(2)
            t_fmt = c1.selectbox("目标格式", ["JPEG", "PNG", "PDF", "TIFF"])
            t_dpi = c2.number_input("DPI", 72, 600, 300)
            if st.button("处理"):
                buf = io.BytesIO()
                if t_fmt == "JPEG" and img.mode == "RGBA": img = img.convert("RGB")
                save_args = {} if t_fmt == "WEBP" else {'dpi': (t_dpi, t_dpi)}
                img.save(buf, format=t_fmt, **save_args)
                st.download_button(f"下载 .{t_fmt}", buf.getvalue(), f"processed.{t_fmt.lower()}", "application/octet-stream")

    elif img_task == "多图拼合转PDF":
        st.header("多图转 PDF")
        files = st.file_uploader("按顺序上传", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        if files and st.button("生成 PDF"):
            imgs = [Image.open(f).convert("RGB") for f in files]
            if imgs:
                buf = io.BytesIO()
                imgs[0].save(buf, "PDF", resolution=100.0, save_all=True, append_images=imgs[1:])
                st.download_button("下载 PDF", buf.getvalue(), "images_merged.pdf", "application/pdf")

st.markdown("---")
st.caption("全能文件处理站 Pro | Streamlit")
