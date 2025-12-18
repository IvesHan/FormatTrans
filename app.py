import streamlit as st
import pandas as pd
import io
from PIL import Image
from PyPDF2 import PdfMerger
import docx

# 设置页面配置
st.set_page_config(page_title="全能文件处理站", page_icon="🛠️", layout="wide")

st.title("🛠️ 全能文件处理站")
st.markdown("按 **文件类型** 分类处理：表格、图片、文档")

# --- 侧边栏：一级导航 ---
category = st.sidebar.selectbox(
    "1️⃣ 选择文件大类",
    ["📊 表格数据 (Excel/CSV/JSON)", "🖼️ 图片处理 (Image)", "📄 文档工具 (PDF/Word)"]
)

# =========================================================
# 模块 A: 表格数据 (保持原有逻辑，优化结构)
# =========================================================
if category == "📊 表格数据 (Excel/CSV/JSON)":
    st.sidebar.markdown("---")
    task = st.sidebar.radio("2️⃣ 选择操作", ["格式互转", "多表合并", "数据排序"])

    # 辅助函数
    def load_table(file):
        try:
            name = file.name
            if name.endswith('.csv'): return pd.read_csv(file)
            elif name.endswith('.tsv'): return pd.read_csv(file, sep='\t')
            elif name.endswith(('.xls', '.xlsx')): return pd.read_excel(file)
            elif name.endswith('.json'): return pd.read_json(file)
        except Exception as e:
            st.error(f"读取错误: {e}")
            return None

    def convert_table(df, fmt):
        buf = io.BytesIO()
        if fmt == "CSV":
            buf.write(df.to_csv(index=False).encode('utf-8-sig'))
            return buf, "text/csv", "csv"
        elif fmt == "Excel":
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            return buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"
        elif fmt == "JSON":
            buf.write(df.to_json(orient='records', force_ascii=False).encode('utf-8'))
            return buf, "application/json", "json"

    if task == "格式互转":
        st.header("表格格式转换")
        f = st.file_uploader("上传表格", type=['csv', 'xlsx', 'xls', 'json'])
        if f:
            df = load_table(f)
            if df is not None:
                st.dataframe(df.head(3))
                fmt = st.selectbox("转为:", ["Excel", "CSV", "JSON"])
                if st.button("转换并下载"):
                    data, mime, ext = convert_table(df, fmt)
                    st.download_button(f"下载 .{ext}", data, f.name.split('.')[0]+f".{ext}", mime)

    elif task == "多表合并":
        st.header("合并多个表格")
        files = st.file_uploader("上传多个结构相同的表格", type=['csv', 'xlsx', 'json'], accept_multiple_files=True)
        if files and st.button("合并"):
            dfs = [load_table(f) for f in files]
            merged = pd.concat(dfs, ignore_index=True)
            st.success(f"合并了 {len(dfs)} 个文件，共 {len(merged)} 行")
            st.dataframe(merged.head())
            data, mime, ext = convert_table(merged, "Excel")
            st.download_button("下载合并后的 Excel", data, "merged.xlsx", mime)

    elif task == "数据排序":
        st.header("表格排序")
        f = st.file_uploader("上传表格", type=['csv', 'xlsx'])
        if f:
            df = load_table(f)
            if df is not None:
                col = st.selectbox("排序列", df.columns)
                asc = st.checkbox("升序 (A->Z)", value=True)
                if st.button("排序"):
                    res = df.sort_values(by=col, ascending=asc)
                    st.dataframe(res.head())
                    data, mime, ext = convert_table(res, "Excel")
                    st.download_button("下载结果", data, "sorted.xlsx", mime)

# =========================================================
# 模块 B: 图片处理 (新增功能)
# =========================================================
elif category == "🖼️ 图片处理 (Image)":
    st.sidebar.markdown("---")
    img_task = st.sidebar.radio("2️⃣ 选择操作", ["格式转换 / 修改PPI", "多图拼合转PDF"])

    if img_task == "格式转换 / 修改PPI":
        st.header("图片格式转换 & DPI 设置")
        st.info("支持 JPG, PNG, BMP, TIFF, WEBP 等互转。")
        
        uploaded_img = st.file_uploader("上传图片", type=['png', 'jpg', 'jpeg', 'bmp', 'tiff', 'webp'])
        
        if uploaded_img:
            image = Image.open(uploaded_img)
            st.image(image, caption=f"原图: {image.size} | 模式: {image.mode}", width=300)
            
            col1, col2 = st.columns(2)
            with col1:
                target_format = st.selectbox("目标格式", ["JPEG", "PNG", "PDF", "TIFF", "BMP", "WEBP"])
            with col2:
                # 默认 DPI 通常是 72 或 96，打印常用 300
                target_dpi = st.number_input("设置 DPI/PPI (像素/英寸)", min_value=72, max_value=600, value=300, step=1)
            
            if st.button("处理图片"):
                buf = io.BytesIO()
                
                # 兼容性处理：JPEG 不支持透明度 (RGBA)，需转为 RGB
                if target_format == "JPEG" and image.mode == "RGBA":
                    image = image.convert("RGB")
                
                # 保存图片，设置 DPI
                try:
                    save_kwargs = {}
                    if target_format != "WEBP": # WEBP saving doesn't always support dpi kwarg consistently in older versions
                        save_kwargs['dpi'] = (target_dpi, target_dpi)
                        
                    image.save(buf, format=target_format, **save_kwargs)
                    buf.seek(0)
                    
                    mime_map = {"JPEG": "image/jpeg", "PNG": "image/png", "PDF": "application/pdf", "TIFF": "image/tiff"}
                    mime = mime_map.get(target_format, "application/octet-stream")
                    ext = target_format.lower()
                    
                    st.success(f"转换成功！DPI 已设为 {target_dpi}")
                    st.download_button(
                        label=f"下载 .{ext}",
                        data=buf,
                        file_name=f"processed_image.{ext}",
                        mime=mime
                    )
                except Exception as e:
                    st.error(f"转换失败: {e}")

    elif img_task == "多图拼合转PDF":
        st.header("多图合并为一个 PDF")
        img_files = st.file_uploader("按顺序上传图片", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        
        if img_files and st.button("生成 PDF"):
            pil_images = []
            for f in img_files:
                img = Image.open(f)
                if img.mode == "RGBA":
                    img = img.convert("RGB")
                pil_images.append(img)
            
            if pil_images:
                pdf_buf = io.BytesIO()
                # 第一张图作为基准，保存其他图为 append
                pil_images[0].save(
                    pdf_buf, "PDF", resolution=100.0, save_all=True, append_images=pil_images[1:]
                )
                pdf_buf.seek(0)
                st.download_button("下载 PDF", pdf_buf, "images_merged.pdf", "application/pdf")

# =========================================================
# 模块 C: 文档工具 (新增功能)
# =========================================================
elif category == "📄 文档工具 (PDF/Word)":
    st.sidebar.markdown("---")
    doc_task = st.sidebar.radio("2️⃣ 选择操作", ["PDF 合并", "Word 转 纯文本", "PDF 提取文本"])

    if doc_task == "PDF 合并":
        st.header("PDF 文件合并")
        pdfs = st.file_uploader("上传多个 PDF", type=['pdf'], accept_multiple_files=True)
        
        if pdfs and st.button("开始合并"):
            merger = PdfMerger()
            for pdf in pdfs:
                merger.append(pdf)
            
            output = io.BytesIO()
            merger.write(output)
            output.seek(0)
            
            st.success("合并完成！")
            st.download_button("下载合并后的 PDF", output, "merged_document.pdf", "application/pdf")

    elif doc_task == "Word 转 纯文本":
        st.header("提取 Word (.docx) 内容")
        st.info("将 Word 文档中的文字快速提取为 TXT 文件。")
        word_file = st.file_uploader("上传 Word 文件", type=['docx'])
        
        if word_file:
            doc = docx.Document(word_file)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            
            text_str = "\n".join(full_text)
            st.text_area("内容预览", text_str, height=300)
            
            st.download_button(
                "下载 .txt 文件",
                text_str,
                word_file.name.replace(".docx", ".txt")
            )

    elif doc_task == "PDF 提取文本":
        st.header("提取 PDF 文本")
        # 注意：这只能提取可选中的文字，扫描件无法提取（需要OCR，那是另一个庞大的库）
        pdf_file = st.file_uploader("上传 PDF", type=['pdf'])
        
        if pdf_file:
            from PyPDF2 import PdfReader
            reader = PdfReader(pdf_file)
            text_content = ""
            for page in reader.pages:
                text_content += page.extract_text() + "\n\n"
            
            st.text_area("提取结果", text_content, height=300)
            st.download_button("下载文本", text_content, "extracted_from_pdf.txt")

# 页脚
st.markdown("---")
st.caption("多功能文件处理站 | 基于 Python Streamlit 构建")