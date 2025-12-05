import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

st.title("通报表扬生成工具")

uploaded_file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df_first = pd.read_excel(uploaded_file)
        
        has_name_in_first = any("姓名" in str(col) or "名字" in str(col) for col in df_first.columns)
        
        if has_name_in_first:
            df = df_first
            st.info("✅ 第一行找到姓名列")
        else:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=1)
            st.info("✅ 第二行找到姓名列")
        
        name_column = None
        for col in df.columns:
            col_str = str(col)
            if "姓名" in col_str or "名字" in col_str:
                name_column = col
                break
        
        if not name_column:
            st.warning("请手动选择姓名列")
            name_column = st.selectbox("选择姓名列：", df.columns)
        else:
            st.success(f"自动识别到姓名列：'{name_column}'")
        
        if name_column:
            names = df[name_column].dropna().astype(str).str.strip().tolist()
            
            names = [name for name in names if name and name != 'nan' and name != 'None']
            
            if not names:
                st.error("没有找到有效的姓名数据")
                st.stop()
            
            st.success(f"✅ 提取到 {len(names)} 个姓名")
            
            with st.expander("查看姓名预览"):
                cols = st.columns(3)
                for i, name in enumerate(names[:15]):
                    with cols[i % 3]:
                        st.write(f"{i+1}. {name}")
                if len(names) > 15:
                    st.write(f"... 等共 {len(names)} 个姓名")
            
            st.subheader("文档设置")
            
            col1, col2 = st.columns(2)
            with col1:
                per_row = st.selectbox("每行姓名数", [2, 3, 4, 5, 6, 7, 8, 9, 10], index=2)
            with col2:
                font_size = st.selectbox("姓名字体大小", [12, 14, 16], index=1)
            
            st.subheader("活动信息")
            col1, col2, col3 = st.columns(3)
            with col1:
                year = st.text_input("年份", "2024")
            with col2:
                month = st.text_input("月份", "10")
            with col3:
                day = st.text_input("日期", "25")
            
            activity = st.text_input("活动名称", "校园文化节")
            
            def to_chinese(num):
                chinese = {'0':'〇','1':'一','2':'二','3':'三','4':'四','5':'五','6':'六','7':'七','8':'八','9':'九'}
                if num == '10': return '十'
                if num == '11': return '十一'
                if num == '12': return '十二'
                return ''.join(chinese[char] for char in num)
            
            if st.button("生成通报表扬"):
                with st.spinner("生成中..."):
                    doc = Document()
                    
                    style = doc.styles['Normal']
                    style.font.name = '宋体'
                    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    style.font.size = Pt(12)
                    
                    title = doc.add_paragraph()
                    title_run = title.add_run("通报表扬")
                    title_run.font.name = '黑体'
                    title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                    title_run.font.size = Pt(28)
                    title_run.bold = True
                    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    doc.add_paragraph()
                    
                    line1 = doc.add_paragraph()
                    line1_run = line1.add_run("各学院团委及学生会：")
                    line1_run.font.name = '宋体'
                    line1_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    line1_run.font.size = Pt(14)
                    line1_run.bold = True
                    
                    line2 = doc.add_paragraph()
                    line2.paragraph_format.first_line_indent = Inches(0.5)
                    line2_text = f"兹有 {year}年 {month}月 {day}日温州理工学院 {activity}活动，在以下同学的共同努力下，本次活动取得了圆满成功，经研究决定，特给予以下同学通报表扬一次："
                    line2_run = line2.add_run(line2_text)
                    line2_run.font.name = '宋体'
                    line2_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    line2_run.font.size = Pt(14)
                    
                    line3 = doc.add_paragraph()
                    line3_run = line3.add_run("具体名单如下：")
                    line3_run.font.name = '宋体'
                    line3_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    line3_run.font.size = Pt(14)
                    line3_run.bold = True
                    
                    doc.add_paragraph()
                    
                    total = len(names)
                    rows = (total + per_row - 1) // per_row
                    table = doc.add_table(rows=rows, cols=per_row)
                    
                    idx = 0
                    for row in table.rows:
                        for cell in row.cells:
                            if idx < total:
                                cell.text = names[idx]
                                paragraph = cell.paragraphs[0]
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                
                                for run in paragraph.runs:
                                    run.font.name = '宋体'
                                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                                    run.font.size = Pt(font_size)
                                
                                idx += 1
                    
                    doc.add_paragraph()
                    
                    footer = doc.add_paragraph()
                    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    
                    footer_run1 = footer.add_run("共青团温州理工学院委员会")
                    footer_run1.font.name = '宋体'
                    footer_run1._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    footer_run1.font.size = Pt(16)
                    footer_run1.bold = True
                    footer.add_run("\n")
                    
                    footer_run2 = footer.add_run(f"{to_chinese(year)}年{to_chinese(month)}月{to_chinese(day)}日")
                    footer_run2.font.name = '宋体'
                    footer_run2._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    footer_run2.font.size = Pt(15)
                    
                    bio = io.BytesIO()
                    doc.save(bio)
                    bio.seek(0)
                    
                    st.success("✅ 通报表扬文档生成成功！")
                    
                    st.download_button(
                        "📥 下载Word文档",
                        bio,
                        f"通报表扬_{activity}.docx",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
    except Exception as e:
        st.error(f"处理文件时出错：{str(e)}")
        st.write("请确保上传的是正确的Excel文件")

else:
    st.info("请上传Excel文件开始使用")
