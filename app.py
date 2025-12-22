import streamlit as st
import pandas as pd
import io
import zipfile
import os
import subprocess
import tempfile
import shutil
from PIL import Image
from pypdf import PdfWriter, PdfReader
from pdf2image import convert_from_bytes
import pikepdf
import docx
import pytesseract

# ==========================================
# 页面基础配置
# ==========================================
st.set_page_config(page_title="全能文件处理站 Pro Max", page_icon="🚀", layout="wide")

st.title("🚀 全能文件处理站 Pro Max")
st.markdown("""
**新增功能**：
* **📄 Office 转 PDF**：支持 Word (.docx) 和 PowerPoint (.pptx) 转换为 PDF (基于 LibreOffice)。
""")

# ==========================================
# 辅助函数定义
# ==========================================

def try_unlock_pdf(file_obj):
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

def libreoffice_convert_to_pdf(source_file_obj, filename):
    """
    使用 LibreOffice 将 Word/PPT 转为 PDF
    """
    # 创建临时目录
    with tempfile.TemporaryDirectory() as temp_dir:
        # 1. 将上传的文件保存到临时路径 (LibreOffice 需要真实文件路径)
        input_path = os.path.join(temp_dir, filename)
        with open(input_path, "wb") as f:
            f.write(source_file_obj.getbuffer())
        
        # 2. 调用 LibreOffice 命令行进行转换
        # --headless: 无界面模式
        # --convert-to pdf: 转换目标
        # --outdir: 输出目录
        cmd = [
            "libreoffice", "--headless", "--convert-to", "pdf", 
            input_path, "--outdir", temp_dir
        ]
        
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError as e:
            raise Exception(f"LibreOffice 转换失败。请确保 packages.txt 包含 libreoffice。错误: {e}")

        # 3. 读取生成的 PDF
        pdf_filename = filename.rsplit('.', 1)[0] + ".pdf"
        output_path = os.path.join(temp_dir, pdf_filename)
        
        if os.path.exists(output_path):
            with open(output_path, "rb") as f:
                pdf_bytes = f.read()
            return pdf_bytes, pdf_filename
        else:
            raise Exception("转换未生成 PDF 文件，可能是源文件格式不兼容。")

# ==========================================
# 侧边栏导航
# ==========================================
category = st.sidebar.selectbox(
    "1️⃣ 选择文件大类",
    ["📊 表格数据 (CSV/Excel)", "📄 文档工具 (PDF/Word/PPT)", "🖼️ 图片处理 (Image)"]
)

st.sidebar.markdown("---")

# =========================================================
# 模块 A: 表格数据
# =========================================================
if category == "📊 表格数据 (CSV/Excel)":
    st.header("表格格式转换")
    sep_option = st.selectbox("输入文件分隔符", ["逗号 , (标准)", "分号 ; (欧洲)", "Tab", "自定义"], index=0)
    separator = ","
    if "分号" in sep_option: separator = ";"
    elif "Tab" in sep_option: separator = "\t"
    elif "自定义" in sep_option: separator = st.text_input("输入自定义分隔符", value="|")

    f = st.file_uploader("上传表格", type=['csv', 'xlsx', 'xls', 'json'])
    if f:
        df = None
        try:
            if f.name.endswith('.csv'): df = pd.read_csv(f, sep=separator)
            elif f.name.endswith('.tsv'): df = pd.read_csv(f, sep='\t')
            elif f.name.endswith(('.xls', '.xlsx')): df = pd.read_excel(f)
            elif f.name.endswith('.json'): df = pd.read_json(f)
        except Exception as e: st.error(f"读取错误: {e}")

        if df is not None:
            st.dataframe(df.head())
            col1, col2 = st.columns(2)
            with col1: target_fmt = st.selectbox("目标格式", ["Excel", "CSV", "JSON"])
            with col2:
                export_sep = ","
                if target_fmt == "CSV": export_sep = st.selectbox("导出CSV分隔符", [",", ";", "\t"], index=0)
            
            if st.button("转换并下载"):
                data, mime, ext = convert_df(df, target_fmt, export_sep)
                st.download_button(f"下载 .{ext}", data, f.name.split('.')[0]+f".{ext}", mime)

# =========================================================
# 模块 B: 文档工具 (新增 Office 转 PDF)
# =========================================================
elif category == "📄 文档工具 (PDF/Word/PPT)":
    doc_task = st.sidebar.radio(
        "2️⃣ 选择操作", 
        ["Office 转 PDF (Word/PPT)", "PDF 合并 (带排序)", "PDF 转 图片", "文本提取 (OCR)", "PDF 权限解除"]
    )

    # --- 新增功能：Office 转 PDF ---
    if doc_task == "Office 转 PDF (Word/PPT)":
        st.header("Word/PPT 转 PDF")
        st.info("基于 LibreOffice 转换。**注意：** 特殊字体可能会变为标准字体 (如 Arial/文泉驿)。")
        
        files = st.file_uploader("上传 Word (.docx) 或 PPT (.pptx)", type=['docx', 'pptx', 'doc', 'ppt'], accept_multiple_files=True)
        
        if files and st.button("开始转换"):
            # 进度条
            progress_bar = st.progress(0)
            
            for i, f in enumerate(files):
                with st.spinner(f"正在转换 {f.name} ..."):
                    try:
                        pdf_data, pdf_name = libreoffice_convert_to_pdf(f, f.name)
                        st.download_button(
                            label=f"📥 下载 {pdf_name}",
                            data=pdf_data,
                            file_name=pdf_name,
                            mime="application/pdf"
                        )
                        st.success(f"✅ {f.name} 转换成功")
                    except Exception as e:
                        st.error(f"❌ {f.name} 转换失败: {e}")
                
                progress_bar.progress((i + 1) / len(files))

    # --- 1. PDF 合并 ---
    elif doc_task == "PDF 合并 (带排序)":
        st.header("PDF 合并 (支持排序)")
        files = st.file_uploader("上传 PDF", type=['pdf'], accept_multiple_files=True)
        if files:
            file_map = {f.name: f for f in files}
            df_files = pd.DataFrame({"文件名": [f.name for f in files], "排序权重": range(1, len(files)+1)})
            edited_df = st.data_editor(df_files, use_container_width=True)
            if st.button("合并"):
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
                        except: pass
                    out = io.BytesIO()
                    merger.write(out)
                    out.seek(0)
                    st.download_button("下载合并 PDF", out, "merged.pdf", "application/pdf")
                except Exception as e: st.error(f"错误: {e}")

    # --- 2. PDF 转 图片 ---
    elif doc_task == "PDF 转 图片":
        st.header("PDF 转图片")
        pdf_file = st.file_uploader("上传 PDF", type=['pdf'])
        dpi = st.number_input("DPI", 72, 600, 200)
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
                st.download_button("下载 ZIP", zip_buf.getvalue(), "images.zip", "application/zip")
            except Exception as e: st.error(f"错误: {e}")

    # --- 3. 文本提取 (OCR) ---
    elif doc_task == "文本提取 (OCR)":
        st.header("文本提取")
        f = st.file_uploader("上传文件", type=['docx', 'pdf'])
        use_ocr = st.checkbox("启用 OCR (扫描件模式)", value=False)
        if f:
            txt = ""
            if f.name.endswith('.docx'):
                doc = docx.Document(f)
                txt = "\n".join([p.text for p in doc.paragraphs])
            elif f.name.endswith('.pdf'):
                if use_ocr:
                    with st.spinner("正在 OCR..."):
                        try:
                            f.seek(0)
                            images = convert_from_bytes(f.read(), dpi=200)
                            full_text = []
                            for img in images:
                                full_text.append(pytesseract.image_to_string(img, lang='chi_sim+eng'))
                            txt = "\n\n".join(full_text)
                        except Exception as e: st.error(f"OCR 错误: {e}")
                else:
                    reader = PdfReader(f)
                    for p in reader.pages: txt += p.extract_text() + "\n"
            if txt:
                st.text_area("结果", txt, height=300)
                st.download_button("下载 .txt", txt, "extracted.txt")

    # --- 4. 权限解除 ---
    elif doc_task == "PDF 权限解除":
        st.header("PDF 权限移除")
        locked = st.file_uploader("上传 PDF", type=['pdf'])
        if locked and st.button("解锁"):
            unlocked = try_unlock_pdf(locked)
            if unlocked:
                unlocked.seek(0)
                st.download_button("下载解锁版", unlocked, f"unlocked_{locked.name}", "application/pdf")

# =========================================================
# 模块 C: 图片处理
# =========================================================
elif category == "🖼️ 图片处理 (Image)":
    img_task = st.sidebar.radio("2️⃣ 选择操作", ["格式转换 / 修改PPI", "多图拼合转PDF"])
    if img_task == "格式转换 / 修改PPI":
        st.header("图片处理")
        f = st.file_uploader("上传", type=['png', 'jpg', 'jpeg', 'bmp', 'tiff', 'webp'])
        if f:
            img = Image.open(f)
            st.image(img, width=200)
            c1, c2 = st.columns(2)
            t_fmt = c1.selectbox("格式", ["JPEG", "PNG", "PDF", "TIFF"])
            t_dpi = c2.number_input("DPI", 72, 600, 300)
            if st.button("处理"):
                buf = io.BytesIO()
                if t_fmt == "JPEG" and img.mode == "RGBA": img = img.convert("RGB")
                save_args = {} if t_fmt == "WEBP" else {'dpi': (t_dpi, t_dpi)}
                img.save(buf, format=t_fmt, **save_args)
                st.download_button(f"下载 .{t_fmt}", buf.getvalue(), f"processed.{t_fmt.lower()}", "image/octet-stream")
    elif img_task == "多图拼合转PDF":
        st.header("多图转 PDF")
        files = st.file_uploader("上传图片", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        if files and st.button("生成 PDF"):
            imgs = [Image.open(f).convert("RGB") for f in files]
            if imgs:
                buf = io.BytesIO()
                imgs[0].save(buf, "PDF", resolution=100.0, save_all=True, append_images=imgs[1:])
                st.download_button("下载 PDF", buf.getvalue(), "images.pdf", "application/pdf")
