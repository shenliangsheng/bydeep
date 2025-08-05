import os
import re
import pdfplumber
import streamlit as st
from docx import Document
from openpyxl import load_workbook
from collections import defaultdict
import tempfile
import io
import base64
import time
import traceback

# 设置页面标题和布局
st.set_page_config(page_title="商标请款单生成系统", layout="wide")
st.title("商标请款单生成系统")

# 初始化session状态
if 'processing_stage' not in st.session_state:
    st.session_state.processing_stage = 0  # 0: 未开始, 1: 提取完成, 2: 生成完成
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = None
if 'manual_input_needed' not in st.session_state:
    st.session_state.manual_input_needed = {}
if 'agent_fees' not in st.session_state:
    st.session_state.agent_fees = {}
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = ""
if 'templates_uploaded' not in st.session_state:
    st.session_state.templates_uploaded = False

# PDF处理函数
def extract_pdf_data(pdf_path):
    """从PDF提取数据"""
    applicant = "N/A"
    unified_credit_code = "N/A"
    final_date = "N/A"
    trademarks_with_categories = []
    pending_categories = []
    
    with pdfplumber.open(pdf_path) as pdf:
        all_texts = [page.extract_text().replace("　", " ").replace("\xa0", " ").strip() 
                     if page.extract_text() else "" for page in pdf.pages]
        all_text_combined = "\n---PAGE_BREAK---\n".join(all_texts)
        pages = all_text_combined.split("\n---PAGE_BREAK---\n")
        
        for page_num, page_text in enumerate(pages):
            # 第一页：提取申请人和统一社会信用代码
            if page_num == 0:
                applicant_match = re.search(r"申请人名称\(中文\)：\s*(.*?)\s*\(\s*英文\)", page_text)
                applicant = applicant_match.group(1).strip() if applicant_match else "N/A"
                
                unified_credit_code_match = re.search(r"统一社会信用代码：\s*([0-9A-Z]+)", page_text)
                unified_credit_code = unified_credit_code_match.group(1).strip() if unified_credit_code_match else "N/A"
                
                # 尝试从第一页提取日期
                if final_date == "N/A":
                    date_match = re.search(r"(\d{4}年\s*\d{1,2}月\s*\d{1,2}日)", page_text)
                    final_date = date_match.group(1).replace(" ", "") if date_match else "N/A"
                continue
            
            # 后续页面：提取类别或商标名
            # 检查是否包含类别信息
            if re.search(r'类别：\d+', page_text):
                categories_found = re.findall(r'类别：(\d+)', page_text)
                pending_categories.extend(categories_found)
            
            # 检查是否包含委托书
            elif '商 标 代 理 委 托 书' in page_text:
                tm_name_match = re.search(r'商标代理委托书.*?代理\s+(.*?)商标\s*的\s*如下.*?事宜', 
                                         page_text, re.DOTALL)
                tm_name = tm_name_match.group(1).strip() if tm_name_match else ""
                
                if not tm_name:
                    fallback_match = re.search(r'代理\s+(.*?)\s*商标', page_text)
                    tm_name = fallback_match.group(1).strip() if fallback_match else ""
                
                if not tm_name:
                    st.warning(f"警告：在文件 {os.path.basename(pdf_path)} 的第 {page_num + 1} 页委托书中未找到商标名称。")
                
                # 提取委托书日期
                date_match = re.search(r"(\d{4}年\s*\d{1,2}月\s*\d{1,2}日)", page_text)
                if date_match:
                    final_date = date_match.group(1).replace(" ", "")
                
                # 关联类别与商标名
                if pending_categories:
                    for category in pending_categories:
                        trademarks_with_categories.append({
                            "商标名称": tm_name,
                            "类别": category
                        })
                    pending_categories.clear()
                else:
                    trademarks_with_categories.append({
                        "商标名称": tm_name,
                        "类别": "MANUAL_INPUT_REQUIRED"
                    })
                    st.warning(f"提示：文件 {os.path.basename(pdf_path)} 中的商标 '{tm_name}' 未找到自动关联的类别，需要手动输入。")
        
        # 检查是否还有未关联的类别
        if pending_categories:
            st.warning(f"警告：文件 {os.path.basename(pdf_path)} 处理完毕，但仍有未关联的类别 {pending_categories}。这些类别将被忽略。")
    
    return {
        "申请人": applicant,
        "统一社会信用代码": unified_credit_code,
        "日期": final_date,
        "商标列表": trademarks_with_categories,
        "事宜类型": "商标注册申请"
    }

# 金额转大写函数
def number_to_upper(amount):
    """金额转大写（支持万、千等单位）"""
    CN_NUM = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    CN_UNIT = ['元', '拾', '佰', '仟', '万', '拾万', '佰万', '仟万', '亿']
    
    s = str(int(amount))[::-1]
    result = []
    
    for i in range(len(s)):
        digit = int(s[i])
        unit = CN_UNIT[i] if i < len(CN_UNIT) else ''
        
        if digit != 0:
            result.append(f"{CN_NUM[digit]}{unit}")
        else:
            if i == 0 and not result:
                result.append("零")
    
    formatted = ''.join(reversed(result))
    return formatted + "元整"

# 生成Word文档函数
def create_word_doc(data, agent_fee, categories, output_dir):
    """生成Word请款单"""
    # 使用临时目录中的模板文件
    template_path = os.path.join(st.session_state.temp_dir, "请款单模板.docx")
    
    if not os.path.exists(template_path):
        st.error(f"找不到模板文件: {template_path}")
        return None
    
    doc = Document(template_path)
    
    num_items = len(categories)
    total_official = num_items * 270
    total_agent = num_items * agent_fee
    total_subtotal = total_official + total_agent
    total_upper = number_to_upper(total_subtotal)

    # 替换段落占位符
    for para in doc.paragraphs:
        if "{申请人}" in para.text:
            para.text = para.text.replace("{申请人}", data["申请人"])
        if "{事宜类型}" in para.text:
            para.text = para.text.replace("{事宜类型}", "商标注册申请")
        if "{日期}" in para.text:
            para.text = para.text.replace("{日期}", data["日期"])
        if "合计：" in para.text:
            para.text = para.text.replace("{总官费}", str(total_official))
            para.text = para.text.replace("{总代理费}", str(total_agent))
            para.text = para.text.replace("{总计}", str(total_subtotal))
            para.text = para.text.replace("{大写}", total_upper)

    # 处理表格
    if doc.tables:
        table = doc.tables[0]
        if len(table.rows) > 2:
            # 删除模板中的示例行
            row_to_delete = table.rows[1]
            tbl = row_to_delete._element
            tbl.getparent().remove(tbl)

        # 添加商标信息行
        for idx, item in enumerate(categories, 1):
            row = table.add_row().cells
            row[0].text = str(idx)  # 序号
            row[1].text = "商标注册申请"  # 商标名称
            row[2].text = item['商标名称']  # 事宜
            row[3].text = item['类别']  # 类别
            row[4].text = f"{270}"  # 官费
            row[5].text = f"{agent_fee}"  # 代理费
            row[6].text = f"{270 + agent_fee}"  # 小计

        # 添加合计行
        total_row = table.add_row().cells
        total_row[0].merge(total_row[3])  # 合并前四个单元格
        total_row[0].text = "合计"
        total_row[0].paragraphs[0].alignment = 1  # 居中对齐
        total_row[4].text = f"{total_official}"  # 总官费
        total_row[5].text = f"{total_agent}"  # 总代理费
        total_row[6].text = f"{total_subtotal}"  # 总计

    # 生成文件名并保存
    filename = f"请款单（{data['申请人']}-商标注册申请-{total_subtotal}-{data['日期']}）.docx"
    output_path = os.path.join(output_dir, filename)
    doc.save(output_path)
    
    return filename

# 生成Excel汇总函数
def create_excel_summary(all_applicants_summary, output_dir):
    """生成Excel汇总表"""
    # 使用临时目录中的模板文件
    template_path = os.path.join(st.session_state.temp_dir, "发票申请表模板.xlsx")
    
    if not os.path.exists(template_path):
        st.error(f"找不到模板文件: {template_path}")
        return None
    
    try:
        wb = load_workbook(template_path)
        ws = wb.active
        row_num = 2
        
        for applicant_data in all_applicants_summary:
            # 官费行
            ws[f'B{row_num}'] = applicant_data["申请人"]
            ws[f'C{row_num}'] = applicant_data["统一社会信用代码"]
            ws[f'G{row_num}'] = applicant_data["总官费"]
            ws[f'H{row_num}'] = applicant_data["总官费"]
            ws[f'I{row_num}'] = applicant_data["总计"]
            ws[f'Q{row_num}'] = applicant_data["日期"]
            row_num += 1
            
            # 代理费行
            ws[f'B{row_num}'] = applicant_data["申请人"]
            ws[f'C{row_num}'] = applicant_data["统一社会信用代码"]
            ws[f'G{row_num}'] = applicant_data["总代理费"]
            ws[f'H{row_num}'] = applicant_data["总代理费"]
            ws[f'I{row_num}'] = applicant_data["总计"]
            ws[f'Q{row_num}'] = applicant_data["日期"]
            row_num += 1
        
        summary_date = all_applicants_summary[0]["日期"] if all_applicants_summary else "N/A"
        excel_filename = f"发票申请表-{summary_date}.xlsx"
        excel_path = os.path.join(output_dir, excel_filename)
        wb.save(excel_path)
        
        return excel_filename
    except Exception as e:
        st.error(f"生成Excel汇总时出错: {str(e)}")
        return None

# 模板文件上传区域
st.sidebar.header("模板文件上传")
st.sidebar.info("请上传以下模板文件以继续操作")

# 请款单模板上传
payment_template = st.sidebar.file_uploader("请款单模板 (Word)", type=["docx"])
if payment_template:
    # 创建临时目录（如果尚未创建）
    if not st.session_state.temp_dir:
        st.session_state.temp_dir = tempfile.mkdtemp()
    
    template_path = os.path.join(st.session_state.temp_dir, "请款单模板.docx")
    with open(template_path, "wb") as f:
        f.write(payment_template.getbuffer())
    st.sidebar.success("请款单模板上传成功！")

# 发票申请表模板上传
invoice_template = st.sidebar.file_uploader("发票申请表模板 (Excel)", type=["xlsx"])
if invoice_template:
    # 创建临时目录（如果尚未创建）
    if not st.session_state.temp_dir:
        st.session_state.temp_dir = tempfile.mkdtemp()
    
    template_path = os.path.join(st.session_state.temp_dir, "发票申请表模板.xlsx")
    with open(template_path, "wb") as f:
        f.write(invoice_template.getbuffer())
    st.sidebar.success("发票申请表模板上传成功！")

# 检查模板是否已上传
if payment_template and invoice_template:
    st.session_state.templates_uploaded = True
    st.sidebar.success("所有模板文件已就绪！")
elif payment_template or invoice_template:
    st.sidebar.warning("请上传所有必需的模板文件")
else:
    st.sidebar.info("请上传所有必需的模板文件以开始")

# 文件上传和处理区域
if st.session_state.templates_uploaded:
    st.header("1. 上传PDF文件")
    uploaded_files = st.file_uploader("请选择PDF文件", type="pdf", accept_multiple_files=True)

    if uploaded_files and st.button("处理PDF文件"):
        with st.spinner("正在处理PDF文件..."):
            try:
                # 创建临时目录（如果尚未创建）
                if not st.session_state.temp_dir:
                    st.session_state.temp_dir = tempfile.mkdtemp()
                    
                pdf_dir = os.path.join(st.session_state.temp_dir, "pdf_files")
                output_dir = os.path.join(st.session_state.temp_dir, "output")
                os.makedirs(pdf_dir, exist_ok=True)
                os.makedirs(output_dir, exist_ok=True)
                
                # 保存上传的文件
                for uploaded_file in uploaded_files:
                    file_path = os.path.join(pdf_dir, uploaded_file.name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                
                # 使用原有逻辑处理PDF
                applicant_data_groups = defaultdict(list)
                manual_input_needed = {}
                
                for filename in os.listdir(pdf_dir):
                    if filename.endswith(".pdf"):
                        try:
                            pdf_path = os.path.join(pdf_dir, filename)
                            data = extract_pdf_data(pdf_path)
                            applicant = data["申请人"]
                            applicant_data_groups[applicant].append(data)
                            
                            # 检查需要手动输入的商标
                            for tm_item in data["商标列表"]:
                                if tm_item["类别"] == "MANUAL_INPUT_REQUIRED":
                                    if applicant not in manual_input_needed:
                                        manual_input_needed[applicant] = []
                                    manual_input_needed[applicant].append(tm_item["商标名称"])
                                    
                        except Exception as e:
                            st.error(f"处理文件 {filename} 时出错: {str(e)}")
                            st.text(traceback.format_exc())
                
                # 保存处理结果到session
                st.session_state.extracted_data = dict(applicant_data_groups)
                st.session_state.manual_input_needed = manual_input_needed
                st.session_state.processing_stage = 1
                
                st.success(f"成功处理 {len(uploaded_files)} 个PDF文件！")
                st.info(f"共发现 {len(applicant_data_groups)} 个申请人")
            except Exception as e:
                st.error(f"处理过程中发生错误: {str(e)}")
                st.text(traceback.format_exc())

# 显示提取结果
if st.session_state.processing_stage >= 1 and st.session_state.extracted_data:
    st.header("2. 提取结果")
    
    for applicant, data_list in st.session_state.extracted_data.items():
        with st.expander(f"申请人: {applicant}"):
            total_trademarks = 0
            for data in data_list:
                total_trademarks += len(data["商标列表"])
                st.write(f"统一社会信用代码: {data.get('统一社会信用代码', 'N/A')}")
                st.write(f"日期: {data.get('日期', 'N/A')}")
                st.write(f"商标数量: {len(data['商标列表'])}")
            
            # 显示需要手动输入的商标
            if applicant in st.session_state.manual_input_needed:
                st.warning("以下商标需要手动输入类别:")
                for tm_name in st.session_state.manual_input_needed[applicant]:
                    st.write(f"- {tm_name}")

# 手动输入区域
if st.session_state.processing_stage >= 1 and st.session_state.extracted_data:
    st.header("3. 设置参数")
    
    # 为每个申请人设置代理费
    st.subheader("代理费设置")
    for applicant in st.session_state.extracted_data.keys():
        default_fee = st.session_state.agent_fees.get(applicant, 1000)
        fee = st.number_input(
            f"{applicant}的代理费(元/项)", 
            min_value=0, 
            value=default_fee,
            key=f"fee_{applicant}"
        )
        st.session_state.agent_fees[applicant] = fee
    
    # 为需要手动输入的商标提供输入框
    if any(st.session_state.manual_input_needed.values()):
        st.subheader("商标类别设置")
        for applicant, tm_list in st.session_state.manual_input_needed.items():
            if tm_list:
                st.markdown(f"**{applicant}**")
                for tm_name in tm_list:
                    categories = st.text_input(
                        f"商标 '{tm_name}' 的类别(多个类别用逗号分隔)", 
                        key=f"manual_{applicant}_{tm_name}",
                        placeholder="例如: 9,35,42"
                    )
    else:
        st.info("没有需要手动输入类别的商标")

# 生成文档按钮
if st.session_state.processing_stage >= 1 and st.session_state.extracted_data and st.button("生成请款单"):
    with st.spinner("正在生成请款单和汇总表..."):
        try:
            output_dir = os.path.join(st.session_state.temp_dir, "output")
            os.makedirs(output_dir, exist_ok=True)
            
            generated_files = []
            all_applicants_summary = []
            
            for applicant, data_list in st.session_state.extracted_data.items():
                try:
                    # 合并商标数据
                    merged_trademarks = []
                    latest_date = "N/A"
                    unified_credit_code = "N/A"
                    
                    for data in data_list:
                        for tm_item in data["商标列表"]:
                            # 处理需要手动输入的商标
                            if tm_item["类别"] == "MANUAL_INPUT_REQUIRED":
                                key = (applicant, tm_item["商标名称"])
                                categories_input = st.session_state.get(f"manual_{applicant}_{tm_item['商标名称']}", "")
                                if categories_input:
                                    categories = [cat.strip() for cat in categories_input.split(",") if cat.strip()]
                                    for cat in categories:
                                        merged_trademarks.append({
                                            "商标名称": tm_item["商标名称"],
                                            "类别": cat
                                        })
                            else:
                                merged_trademarks.append(tm_item)
                        
                        # 更新最新日期和统一社会信用代码
                        if data["日期"] != "N/A":
                            latest_date = data["日期"]
                        if data["统一社会信用代码"] != "N/A":
                            unified_credit_code = data["统一社会信用代码"]
                    
                    # 准备数据
                    merged_data = {
                        "申请人": applicant,
                        "统一社会信用代码": unified_credit_code,
                        "日期": latest_date,
                        "商标列表": merged_trademarks,
                        "事宜类型": "商标注册申请"
                    }
                    
                    # 获取代理费
                    agent_fee = st.session_state.agent_fees.get(applicant, 1000)
                    
                    # 生成Word文档
                    if merged_trademarks:
                        word_filename = create_word_doc(
                            merged_data, 
                            agent_fee, 
                            merged_trademarks,
                            output_dir
                        )
                        
                        if word_filename:
                            word_path = os.path.join(output_dir, word_filename)
                            with open(word_path, "rb") as f:
                                word_data = f.read()
                            
                            generated_files.append({
                                "name": word_filename,
                                "data": word_data,
                                "type": "word"
                            })
                            
                            # 收集汇总数据
                            num_items = len(merged_trademarks)
                            total_official = num_items * 270
                            total_agent = num_items * agent_fee
                            total_subtotal = total_official + total_agent
                            
                            all_applicants_summary.append({
                                "申请人": applicant,
                                "统一社会信用代码": unified_credit_code,
                                "日期": latest_date,
                                "总官费": total_official,
                                "总代理费": total_agent,
                                "总计": total_subtotal
                            })
                
                except Exception as e:
                    st.error(f"为申请人 '{applicant}' 生成请款单时出错: {str(e)}")
                    st.text(traceback.format_exc())
            
            # 生成Excel汇总
            if all_applicants_summary:
                excel_filename = create_excel_summary(all_applicants_summary, output_dir)
                if excel_filename:
                    excel_path = os.path.join(output_dir, excel_filename)
                    with open(excel_path, "rb") as f:
                        excel_data = f.read()
                    
                    generated_files.append({
                        "name": excel_filename,
                        "data": excel_data,
                        "type": "excel"
                    })
            
            # 保存生成的文件到session
            st.session_state.generated_files = generated_files
            st.session_state.processing_stage = 2
            st.success("文档生成完成！")
        except Exception as e:
            st.error(f"生成过程中发生错误: {str(e)}")
            st.text(traceback.format_exc())

# 下载区域
if st.session_state.processing_stage == 2 and st.session_state.generated_files:
    st.header("4. 下载生成的文件")
    
    # 显示所有生成的文件
    st.subheader("生成的文件列表")
    
    word_files = [f for f in st.session_state.generated_files if f["type"] == "word"]
    excel_files = [f for f in st.session_state.generated_files if f["type"] == "excel"]
    
    if word_files:
        st.subheader("请款单")
        for file in word_files:
            st.download_button(
                label=f"下载 {file['name']}",
                data=file["data"],
                file_name=file["name"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    
    if excel_files:
        st.subheader("汇总表")
        for file in excel_files:
            st.download_button(
                label=f"下载 {file['name']}",
                data=file["data"],
                file_name=file["name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# 重置按钮
if st.button("重置所有数据"):
    # 清除所有session状态
    keys_to_clear = list(st.session_state.keys())
    for key in keys_to_clear:
        del st.session_state[key]
    
    # 重新初始化必要的状态
    st.session_state.processing_stage = 0
    st.session_state.extracted_data = None
    st.session_state.manual_input_needed = {}
    st.session_state.agent_fees = {}
    st.session_state.generated_files = []
    st.session_state.templates_uploaded = False
    
    # 重新创建临时目录
    st.session_state.temp_dir = tempfile.mkdtemp()
    
    st.success("系统已重置，可以开始新的处理流程！")
