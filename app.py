import streamlit as st
import pandas as pd
import io
import zipfile
from PIL import Image
from pypdf import PdfWriter, PdfReader
from pdf2image import convert_from_bytes
import pikepdf
import docx

# ==========================================
# 页面基础配置
# ==========================================
st.set_page_config(page_title="Ives全能文件处理站 Pro", page_icon="🛠️", layout="wide")

st.title("🛠️ Ives全能文件处理站 Pro")
st.markdown("""
**功能概览**：
* **📊 表格**：支持 CSV (中/英/法格式)、Excel、JSON 互转与合并。
* **📄 文档**：PDF 排序合并、PDF 转高清图、**解除 PDF 打印/复制限制**、Word/PDF 转纯文本。
* **🖼️ 图片**：格式互转、修改 DPI (PPI)、多图拼合转 PDF。
""")

# ==========================================
# 辅助函数定义
# ==========================================

def try_unlock_pdf(file_obj):
    """尝试去除PDF权限限制 (Owner Password)"""
    try:
        # pikepdf 可以在不知道 owner password 的情况下移除编辑/打印限制
        pdf = pikepdf.open(file_obj)
        new_pdf_bytes = io.BytesIO()
        pdf.save(new_pdf_bytes)
        return new_pdf_bytes
    except pikepdf.PasswordError:
        st.error("❌ 此文件设置了【打开密码】(User Password)，无法强制破除。")
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
# 模块 A: 表格数据 (CSV/Excel/JSON)
# =========================================================
if category == "📊 表格数据 (CSV/Excel)":
    task = st.sidebar.radio("2️⃣ 选择操作", ["格式互转/读取", "多表合并", "数据排序"])
    
    # --- 全局设置：CSV 分隔符 ---
    st.sidebar.markdown("### ⚙️ CSV 读取设置")
    sep_option = st.sidebar.selectbox(
        "输入文件分隔符",
        ["逗号 , (标准/英语系统)", "分号 ; (法语/欧洲系统)", "Tab (制表符)", "自定义"],
        index=0
    )
    separator = ","
    if "分号" in sep_option: separator = ";"
    elif "Tab" in sep_option: separator = "\t"
    elif "自定义" in sep_option:
        separator = st.sidebar.text_input("输入自定义分隔符", value="|")

    def load_table(file, sep):
        try:
            name = file.name
            if name.endswith('.csv'): return pd.read_csv(file, sep=sep)
            elif name.endswith('.tsv'): return pd.read_csv(file, sep='\t')
            elif name.endswith(('.xls', '.xlsx')): return pd.read_excel(file)
            elif name.endswith('.json'): return pd.read_json(file)
        except Exception as e:
            st.error(f"读取错误 ({file.name}): {e}")
            return None

    # --- 子功能：格式转换 ---
    if task == "格式互转/读取":
        st.header("表格读取与转换")
        f = st.file_uploader("上传表格", type=['csv', 'xlsx', 'xls', 'json'])
        if f:
            df = load_table(f, separator)
            if df is not None:
                st.write("### 数据预览 (前5行)")
                st.dataframe(df.head())
                
                col1, col2 = st.columns(2)
                with col1:
                    target_fmt = st.selectbox("目标格式", ["Excel", "CSV", "JSON"])
                with col2:
                    export_sep = ","
                    if target_fmt == "CSV":
                        export_sep = st.selectbox("导出CSV分隔符", [",", ";", "\t"], index=0, help="法语系统建议选分号")
                
                if st.button("转换并下载"):
                    data, mime, ext = convert_df(df, target_fmt, export_sep)
                    st.download_button(f"下载 .{ext}", data, f.name.split('.')[0]+f".{ext}", mime)

    # --- 子功能：多表合并 ---
    elif task == "多表合并":
        st.header("合并多个表格")
        files = st.file_uploader("上传多个结构相同的表格", type=['csv', 'xlsx', 'json'], accept_multiple_files=True)
        if files and st.button("开始合并"):
            dfs = []
            for f in files:
                d = load_table(f, separator)
                if d is not None: dfs.append(d)
            
            if dfs:
                merged = pd.concat(dfs, ignore_index=True)
                st.success(f"成功合并 {len(dfs)} 个文件，共 {len(merged)} 行。")
                st.dataframe(merged.head())
                
                data, mime, ext = convert_df(merged, "Excel")
                st.download_button("下载合并结果 (Excel)", data, "merged_data.xlsx", mime)

    # --- 子功能：排序 ---
    elif task == "数据排序":
        st.header("数据排序")
        f = st.file_uploader("上传表格", type=['csv', 'xlsx'])
        if f:
            df = load_table(f, separator)
            if df is not None:
                col = st.selectbox("选择排序列", df.columns)
                asc = st.radio("排序方式", ["升序 (A-Z)", "降序 (Z-A)"]) == "升序 (A-Z)"
                
                if st.button("执行排序"):
                    res = df.sort_values(by=col, ascending=asc)
                    st.dataframe(res.head())
                    data, mime, ext = convert_df(res, "Excel")
                    st.download_button("下载排序结果", data, "sorted_data.xlsx", mime)

# =========================================================
# 模块 B: 文档工具 (PDF/Word)
# =========================================================
elif category == "📄 文档工具 (PDF/Word)":
    doc_task = st.sidebar.radio("2️⃣ 选择操作", ["PDF 合并 (带排序)", "PDF 转 图片 (含DPI)", "PDF 权限解除", "提取文本 (Word/PDF)"])

    # --- 子功能：PDF 合并 (带排序) ---
    if doc_task == "PDF 合并 (带排序)":
        st.header("PDF 合并 (支持自定义排序)")
        files = st.file_uploader("上传多个 PDF", type=['pdf'], accept_multiple_files=True)
        
        if files:
            # 创建排序界面
            file_map = {f.name: f for f in files}
            df_files = pd.DataFrame({"文件名": [f.name for f in files], "排序权重": range(1, len(files)+1)})
            st.info("👇 在下方表格修改数字以调整顺序 (1排最前)")
            edited_df = st.data_editor(df_files, use_container_width=True)
            
            if st.button("按顺序合并"):
                sorted_names = edited_df.sort_values(by="排序权重")["文件名"].tolist()
                merger = PdfWriter()
                
                try:
                    for name in sorted_names:
                        f_obj = file_map[name]
                        f_obj.seek(0)
                        
                        # 尝试处理加密文件
                        try:
                            reader = PdfReader(f_obj)
                            if reader.is_encrypted:
                                f_obj.seek(0)
                                unlocked = try_unlock_pdf(f_obj)
                                if unlocked: reader = PdfReader(unlocked)
                                else: continue # 无法解密则跳过
                            merger.append(reader)
                        except Exception as e:
                            st.error(f"跳过文件 {name}: {e}")
                    
                    out = io.BytesIO()
                    merger.write(out)
                    out.seek(0)
                    st.success("合并完成！")
                    st.download_button("下载合并 PDF", out, "merged_sorted.pdf", "application/pdf")
                except Exception as e:
                    st.error(f"合并出错: {e}")

    # --- 子功能：PDF 转图片 ---
    elif doc_task == "PDF 转 图片 (含DPI)":
        st.header("PDF 转图片")
        pdf_file = st.file_uploader("上传 PDF", type=['pdf'])
        
        col1, col2 = st.columns(2)
        with col1:
            dpi = st.number_input("DPI (清晰度)", 72, 600, 200, step=50)
        with col2:
            fmt = st.selectbox("输出格式", ["JPEG", "PNG"])
            
        if pdf_file and st.button("开始转换"):
            try:
                # 预处理：解锁
                pdf_reader = PdfReader(pdf_file)
                if pdf_reader.is_encrypted:
                    pdf_file.seek(0)
                    pdf_stream = try_unlock_pdf(pdf_file)
                    if not pdf_stream: st.stop()
                    bytes_data = pdf_stream.read()
                else:
                    pdf_file.seek(0)
                    bytes_data = pdf_file.read()

                # 转换
                images = convert_from_bytes(bytes_data, dpi=dpi)
                st.success(f"成功转换 {len(images)} 页。")
                
                if len(images) == 1:
                    buf = io.BytesIO()
                    images[0].save(buf, format=fmt)
                    st.download_button("下载图片", buf.getvalue(), f"page.1.{fmt.lower()}", f"image/{fmt.lower()}")
                else:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as zf:
                        for i, img in enumerate(images):
                            ib = io.BytesIO()
                            img.save(ib, format=fmt)
                            zf.writestr(f"page_{i+1:03d}.{fmt.lower()}", ib.getvalue())
                    st.download_button("下载所有图片 (ZIP)", zip_buf.getvalue(), "pdf_images.zip", "application/zip")
            except Exception as e:
                st.error(f"转换失败 (请确保服务器安装了 poppler): {e}")

    # --- 子功能：PDF 权限解除 ---
    elif doc_task == "PDF 权限解除":
        st.header("🔒 PDF 权限移除")
        st.markdown("移除 **禁止打印、禁止复制** 等限制 (需无打开密码)。")
        locked = st.file_uploader("上传受限 PDF", type=['pdf'])
        if locked and st.button("解锁"):
            unlocked = try_unlock_pdf(locked)
            if unlocked:
                unlocked.seek(0)
                st.success("解锁成功！")
                st.download_button("下载解锁版 PDF", unlocked, f"unlocked_{locked.name}", "application/pdf")

    # --- 子功能：提取文本 ---
    elif doc_task == "提取文本 (Word/PDF)":
        st.header("提取纯文本")
        f = st.file_uploader("上传 Word 或 PDF", type=['docx', 'pdf'])
        if f:
            txt = ""
            if f.name.endswith('.docx'):
                doc = docx.Document(f)
                txt = "\n".join([p.text for p in doc.paragraphs])
            elif f.name.endswith('.pdf'):
                reader = PdfReader(f)
                for p in reader.pages:
                    txt += p.extract_text() + "\n\n"
            
            st.text_area("内容预览", txt, height=300)
            st.download_button("下载 .txt", txt, "extracted.txt")

# =========================================================
# 模块 C: 图片处理
# =========================================================
elif category == "🖼️ 图片处理 (Image)":
    img_task = st.sidebar.radio("2️⃣ 选择操作", ["格式转换 / 修改PPI", "多图拼合转PDF"])

    # --- 子功能：图片转换 ---
    if img_task == "格式转换 / 修改PPI":
        st.header("图片处理")
        f = st.file_uploader("上传图片", type=['png', 'jpg', 'jpeg', 'bmp', 'tiff', 'webp'])
        if f:
            img = Image.open(f)
            st.image(img, caption=f"原尺寸: {img.size}", width=300)
            
            c1, c2 = st.columns(2)
            t_fmt = c1.selectbox("目标格式", ["JPEG", "PNG", "PDF", "TIFF", "BMP"])
            t_dpi = c2.number_input("设置 DPI/PPI", 72, 600, 300)
            
            if st.button("处理"):
                buf = io.BytesIO()
                if t_fmt == "JPEG" and img.mode == "RGBA": img = img.convert("RGB")
                
                save_args = {}
                if t_fmt != "WEBP": save_args['dpi'] = (t_dpi, t_dpi)
                
                img.save(buf, format=t_fmt, **save_args)
                mime_map = {"JPEG": "image/jpeg", "PNG": "image/png", "PDF": "application/pdf"}
                st.download_button(f"下载 .{t_fmt.lower()}", buf.getvalue(), f"processed.{t_fmt.lower()}", mime_map.get(t_fmt))

    # --- 子功能：图片转 PDF ---
    elif img_task == "多图拼合转PDF":
        st.header("多图转 PDF")
        files = st.file_uploader("按顺序上传图片", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        if files and st.button("生成 PDF"):
            imgs = []
            for f in files:
                i = Image.open(f)
                if i.mode == "RGBA": i = i.convert("RGB")
                imgs.append(i)
            
            if imgs:
                buf = io.BytesIO()
                imgs[0].save(buf, "PDF", resolution=100.0, save_all=True, append_images=imgs[1:])
                st.download_button("下载 PDF", buf.getvalue(), "images_merged.pdf", "application/pdf")

st.markdown("---")
st.caption("全能文件处理站 Pro | Powered by Streamlit")

