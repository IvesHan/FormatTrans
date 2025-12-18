import streamlit as st
import pandas as pd
import io
import zipfile
from PIL import Image
from pypdf import PdfWriter, PdfReader
from pdf2image import convert_from_bytes
import pikepdf
import docx

# 设置页面配置
st.set_page_config(page_title="Ives全能文件处理站 Pro", page_icon="🛠️", layout="wide")

st.title("🛠️ Ives全能文件处理站 Pro")
st.markdown("增强功能：**CSV多分隔符支持、PDF排序合并、PDF转图、权限解除**")

# --- 侧边栏：一级导航 ---
category = st.sidebar.selectbox(
    "1️⃣ 选择文件大类",
    ["📊 表格数据 (CSV/Excel)", "📄 文档工具 (PDF/Word)", "🖼️ 图片处理 (Image)"]
)

# =========================================================
# 辅助函数区
# =========================================================

def try_unlock_pdf(file_obj):
    """尝试去除PDF权限限制"""
    try:
        # pikepdf 可以在不知道 owner password 的情况下移除编辑/打印限制
        pdf = pikepdf.open(file_obj)
        # 如果能打开，说明没有 user password (打开密码)，或者密码为空
        # 创建一个新的流
        new_pdf_bytes = io.BytesIO()
        pdf.save(new_pdf_bytes)
        return new_pdf_bytes
    except pikepdf.PasswordError:
        st.error("此文件设置了【打开密码】(User Password)，无法强制破除。请输入密码解密（暂不支持前端输入密码解密）。")
        return None
    except Exception as e:
        st.error(f"权限处理失败: {e}")
        return None

# =========================================================
# 模块 A: 表格数据 (增强 CSV 分隔符支持)
# =========================================================
if category == "📊 表格数据 (CSV/Excel)":
    st.sidebar.markdown("---")
    task = st.sidebar.radio("2️⃣ 选择操作", ["格式互转/读取", "多表合并"])

    # --- CSV 分隔符设置 ---
    st.sidebar.markdown("### ⚙️ CSV 读取设置")
    sep_option = st.sidebar.selectbox(
        "选择 CSV 分隔符",
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
            if name.endswith('.csv'): 
                return pd.read_csv(file, sep=sep)
            elif name.endswith('.tsv'): 
                return pd.read_csv(file, sep='\t')
            elif name.endswith(('.xls', '.xlsx')): 
                return pd.read_excel(file)
            elif name.endswith('.json'): 
                return pd.read_json(file)
        except Exception as e:
            st.error(f"读取错误 ({name}): {e}")
            return None

    if task == "格式互转/读取":
        st.header("表格读取与转换")
        st.info(f"当前使用的 CSV 分隔符为: `{separator}` (可在侧边栏修改)")
        
        f = st.file_uploader("上传表格", type=['csv', 'xlsx', 'xls', 'json'])
        if f:
            df = load_table(f, separator)
            if df is not None:
                st.write("### 数据预览")
                st.dataframe(df.head())
                
                target_fmt = st.selectbox("转为:", ["Excel", "CSV", "JSON"])
                if st.button("转换并下载"):
                    buf = io.BytesIO()
                    if target_fmt == "Excel":
                        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                            df.to_excel(writer, index=False)
                        mime, ext = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"
                    elif target_fmt == "CSV":
                        # 欧洲系统通常习惯用分号，这里可以给个选项，或者默认标准逗号
                        export_sep = st.selectbox("导出 CSV 分隔符", [",", ";", "\t"], index=0)
                        buf.write(df.to_csv(index=False, sep=export_sep).encode('utf-8-sig'))
                        mime, ext = "text/csv", "csv"
                    else: # JSON
                        buf.write(df.to_json(orient='records', force_ascii=False).encode('utf-8'))
                        mime, ext = "application/json", "json"
                    
                    buf.seek(0)
                    st.download_button(f"下载 .{ext}", buf, f.name.split('.')[0]+f".{ext}", mime)

# =========================================================
# 模块 B: 文档工具 (增强 PDF 排序、转图、权限)
# =========================================================
elif category == "📄 文档工具 (PDF/Word)":
    st.sidebar.markdown("---")
    doc_task = st.sidebar.radio("2️⃣ 选择操作", ["PDF 合并 (支持排序)", "PDF 转 图片", "权限解除 (Unlock)"])

    # --- 功能 1: PDF 合并 (带排序) ---
    if doc_task == "PDF 合并 (支持排序)":
        st.header("PDF 合并")
        files = st.file_uploader("上传多个 PDF", type=['pdf'], accept_multiple_files=True)
        
        if files:
            st.write("### 📂 文件排序")
            st.info("在下方表格中修改 **'排序权重'** 数字来调整合并顺序（数字越小越靠前）。")
            
            # 创建一个 DataFrame 来管理文件顺序
            file_map = {f.name: f for f in files}
            df_files = pd.DataFrame({
                "文件名": [f.name for f in files],
                "排序权重": range(1, len(files) + 1)
            })
            
            # 使用 st.data_editor 允许用户修改
            edited_df = st.data_editor(df_files, use_container_width=True)
            
            if st.button("按指定顺序合并"):
                # 根据用户编辑后的权重排序
                sorted_files_names = edited_df.sort_values(by="排序权重")["文件名"].tolist()
                
                merger = PdfWriter()
                
                try:
                    for name in sorted_files_names:
                        f_obj = file_map[name]
                        f_obj.seek(0) # 重置指针
                        
                        # 尝试读取，如果加密则尝试解密
                        try:
                            reader = PdfReader(f_obj)
                            if reader.is_encrypted:
                                st.warning(f"检测到 {name} 有加密，尝试去除权限...")
                                # 使用 pikepdf 处理后的流
                                f_obj.seek(0)
                                unlocked_stream = try_unlock_pdf(f_obj)
                                if unlocked_stream:
                                    reader = PdfReader(unlocked_stream)
                                else:
                                    st.stop() # 无法解密则停止
                            
                            merger.append(reader)
                            
                        except Exception as e:
                            st.error(f"处理文件 {name} 时出错: {e}")
                    
                    output = io.BytesIO()
                    merger.write(output)
                    output.seek(0)
                    st.success("合并成功！")
                    st.download_button("下载合并 PDF", output, "merged_sorted.pdf", "application/pdf")
                    
                except Exception as e:
                    st.error(f"合并失败: {e}")

    # --- 功能 2: PDF 转图片 ---
    elif doc_task == "PDF 转 图片":
        st.header("PDF 转图片 (JPG/PNG)")
        st.warning("注意：此功能需消耗较多内存，大文件请耐心等待。")
        
        pdf_file = st.file_uploader("上传 PDF", type=['pdf'])
        
        col1, col2 = st.columns(2)
        with col1:
            dpi_val = st.number_input("设置 DPI (清晰度)", min_value=72, max_value=600, value=200, step=50, help="屏幕查看72-150，打印建议300以上")
        with col2:
            img_fmt = st.selectbox("输出格式", ["JPEG", "PNG"])
            
        if pdf_file and st.button("开始转换"):
            try:
                # 检查加密
                pdf_reader = PdfReader(pdf_file)
                if pdf_reader.is_encrypted:
                    st.warning("检测到加密，正在尝试解除权限...")
                    pdf_file.seek(0)
                    pdf_stream = try_unlock_pdf(pdf_file)
                    if not pdf_stream: st.stop()
                    bytes_data = pdf_stream.read()
                else:
                    pdf_file.seek(0)
                    bytes_data = pdf_file.read()

                # 使用 pdf2image 转换
                images = convert_from_bytes(bytes_data, dpi=dpi_val)
                
                st.success(f"转换成功，共 {len(images)} 页。")
                
                # 如果只有1页，直接下载图片
                if len(images) == 1:
                    img_buf = io.BytesIO()
                    images[0].save(img_buf, format=img_fmt)
                    img_buf.seek(0)
                    st.download_button(f"下载图片", img_buf, f"page_1.{img_fmt.lower()}", f"image/{img_fmt.lower()}")
                
                # 如果有多页，打包成 ZIP
                else:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as zf:
                        for i, img in enumerate(images):
                            img_byte_arr = io.BytesIO()
                            img.save(img_byte_arr, format=img_fmt)
                            zf.writestr(f"page_{i+1:03d}.{img_fmt.lower()}", img_byte_arr.getvalue())
                    
                    zip_buf.seek(0)
                    st.download_button("下载所有图片 (ZIP)", zip_buf, "pdf_images.zip", "application/zip")
                    
            except Exception as e:
                st.error(f"转换失败 (请检查是否安装了 poppler): {e}")

    # --- 功能 3: 纯权限解除 ---
    elif doc_task == "权限解除 (Unlock)":
        st.header("🔒 PDF 权限/密码移除")
        st.markdown("""
        此功能用于去除 PDF 的 **Owner Password** (如禁止打印、禁止复制)。
        *如果文件有 **User Password** (打开即需密码)，则无法在此强制破除。*
        """)
        
        locked_file = st.file_uploader("上传受限 PDF", type=['pdf'])
        
        if locked_file:
            if st.button("尝试破除限制"):
                result_stream = try_unlock_pdf(locked_file)
                if result_stream:
                    result_stream.seek(0)
                    st.success("成功！权限限制已移除。")
                    st.download_button("下载解锁版 PDF", result_stream, f"unlocked_{locked_file.name}", "application/pdf")

# =========================================================
# 模块 C: 图片处理 (保持不变)
# =========================================================
elif category == "🖼️ 图片处理 (Image)":
    st.info("图片功能参考上一版代码，此处从略以节省篇幅...")
    # 这里可以保留上一版本的图片处理代码

