import fitz  # PyMuPDF
import json
import os
import csv
import re

# 坐标匹配误差范围（单位：点）
COORDINATE_TOLERANCE = 8
# 发票号码和开票日期的特殊容差
DATE_NUMBER_TOLERANCE = 2

def get_text_coordinates(pdf_path):
    """获取PDF中文本的坐标信息，使用迭代方式处理"""
    doc = fitz.open(pdf_path)
    page = doc[0]
    coordinates = []
    
    # 获取所有文本块
    text_dict = page.get_text("dict")
    
    # 使用列表来存储待处理的项目，而不是递归
    items_to_process = [(block, "block") for block in text_dict["blocks"]]
    
    while items_to_process:
        item, item_type = items_to_process.pop(0)
        
        if item_type == "block" and "lines" in item:
            # 将lines添加到待处理列表
            items_to_process.extend([(line, "line") for line in item["lines"]])
        
        elif item_type == "line":
            # 将spans添加到待处理列表
            items_to_process.extend([(span, "span") for span in item["spans"]])
        
        elif item_type == "span":
            # 处理文本span
            text = ''.join(item["text"].split())
            if text:
                # 如果文本是数字，在前面加上单引号
                if text.replace('.', '').replace('E', '').replace('+', '').replace('-', '').isdigit():
                    text = "'" + text
                coordinates.append({
                    "text": text,
                    "left": float(round(item["bbox"][0], 2)),
                    "top": float(round(item["bbox"][1], 2)),
                    "right": float(round(item["bbox"][2], 2)),
                    "bottom": float(round(item["bbox"][3], 2))
                })
    
    doc.close()
    return coordinates

def extract_invoice_fields(coordinates):
    """提取发票字段信息"""
    fields = {
        "invoice_type": "",
        "invoice_number": "",
        "invoice_date": "",
        "buyer_name": "",
        "buyer_tax_id": "",
        "seller_name": "",
        "seller_tax_id": "",
        "net_amount": "",
        "tax_amount": "",
        "total_amount": ""
    }
    
    # 获取文件的最大right值
    max_right = max(coord["right"] for coord in coordinates)
    
    # 1. 发票类型提取规则
    if coordinates:
        # 直接选择 top 值最小的文本作为发票类型
        top_coord = min(coordinates, key=lambda x: x["top"])
        fields["invoice_type"] = top_coord["text"]
    
    # 2. 发票号码和开票日期提取规则
    # 按 top 值排序所有坐标
    sorted_coords = sorted(coordinates, key=lambda x: x["top"])
    
    # 找到"购"字的位置
    gou_index = len(sorted_coords)
    for i, coord in enumerate(sorted_coords):
        if coord["text"] == "购":
            gou_index = i
            break
    
    # 在"购"字之前的文本中查找发票号码和开票日期
    for coord in sorted_coords[:gou_index]:
        text = coord["text"].replace("'", "")  # 移除可能的单引号前缀
        
        # 检查发票号码：长度超过6位的纯数字
        if text.isdigit() and len(text) > 6 and not fields["invoice_number"]:
            fields["invoice_number"] = text
        
        # 检查开票日期：长度为11的包含数字但不是纯数字字符串
        if len(text) == 11 and any(c.isdigit() for c in text) and not text.isdigit() and not fields["invoice_date"]:
            fields["invoice_date"] = text
    
    # 4. 购买方和销售方信息提取规则
    gou = None
    xiao = None
    xin_xi_list = []
    
    for coord in coordinates:
        if coord["text"] == "购":
            gou = coord
        elif coord["text"] == "销":
            xiao = coord
        elif coord["text"] == "息":
            xin_xi_list.append(coord)
    
    if gou and xiao and len(xin_xi_list) >= 2:
        # 确定信息区域边界
        bottom = max(x["bottom"] for x in xin_xi_list)
        right_min = min(x["right"] for x in xin_xi_list)
        right_max = max(x["right"] for x in xin_xi_list)
        
        # 提取购买方信息
        buyer_info = []
        for coord in coordinates:
            if (coord["top"] >= gou["top"] and 
                coord["bottom"] <= bottom and 
                coord["left"] >= gou["right"] and 
                coord["right"] <= xiao["left"]):
                if coord["text"] not in ["名称", "统一社会信用代码/纳税人识别号:"]:
                    buyer_info.append(coord)
        
        # 分离名称和税号
        for info in buyer_info:
            if any(c.isdigit() for c in info["text"]):
                fields["buyer_tax_id"] = info["text"]
            else:
                fields["buyer_name"] = info["text"]
        
        # 提取销售方信息
        seller_info = []
        for coord in coordinates:
            if (coord["top"] >= xiao["top"] and 
                coord["bottom"] <= bottom and 
                coord["left"] >= xiao["right"] and 
                coord["right"] <= max_right):
                if coord["text"] not in ["名称", "统一社会信用代码/纳税人识别号:"]:
                    seller_info.append(coord)
        
        # 分离销售方名称和税号
        for info in seller_info:
            if any(c.isdigit() for c in info["text"]):
                fields["seller_tax_id"] = info["text"].replace("'", "")
            else:
                fields["seller_name"] = info["text"]
    
    # 4. 金额和税额提取规则
    he_ji = None
    he = None
    ji = None
    
    # 先尝试找完整的"合计"
    for coord in coordinates:
        if coord["text"] == "合计":
            he_ji = coord
            break
    
    # 如果没找到完整的"合计"，尝试找分开的"合"和"计"
    if not he_ji:
        for coord in coordinates:
            if coord["text"] == "合":
                he = coord
            elif coord["text"] == "计":
                ji = coord
            # 如果都找到了，检查是否在同一水平线上
            if he and ji and abs(he["top"] - ji["top"]) < COORDINATE_TOLERANCE:
                # 创建一个虚拟的合计坐标对象
                he_ji = {
                    "text": "合计",
                    "left": he["left"],
                    "top": he["top"],
                    "right": ji["right"],
                    "bottom": ji["bottom"]
                }
                break
    
    if he_ji:
        amounts = []
        for coord in coordinates:
            if (abs(coord["top"] - he_ji["top"]) < COORDINATE_TOLERANCE and 
                any(c.isdigit() for c in coord["text"])):
                text = coord["text"].replace("¥", "").replace("'", "")
                amounts.append((float(coord["left"]), text))
        
        if len(amounts) >= 2:
            amounts.sort(key=lambda x: x[0])
            fields["net_amount"] = amounts[0][1]
            fields["tax_amount"] = amounts[1][1]
    
    # 5. 价税合计提取规则
    jia_shui_he_ji = None
    for coord in coordinates:
        if "价税合计" in coord["text"]:
            jia_shui_he_ji = coord
            break
    
    if jia_shui_he_ji:
        for coord in coordinates:
            if (abs(coord["top"] - jia_shui_he_ji["top"]) < COORDINATE_TOLERANCE and 
                any(c.isdigit() for c in coord["text"])):
                fields["total_amount"] = coord["text"].replace("¥", "").replace("'", "")
    
    return fields

def create_annotation(pdf_path, coordinates):
    """创建标注文件"""
    # 获取PDF尺寸
    doc = fitz.open(pdf_path)
    page = doc[0]
    width = page.rect.width
    height = page.rect.height
    doc.close()
    
    # 提取字段信息
    fields = extract_invoice_fields(coordinates)
    
    # 创建标注结构
    annotation = {
        "invoice_type": fields["invoice_type"],
        "invoice_number": fields["invoice_number"],
        "invoice_date": fields["invoice_date"],
        "buyer_name": fields["buyer_name"],
        "buyer_tax_id": fields["buyer_tax_id"],
        "seller_name": fields["seller_name"],
        "seller_tax_id": fields["seller_tax_id"],
        "net_amount": fields["net_amount"],
        "tax_amount": fields["tax_amount"],
        "total_amount": fields["total_amount"],
        "nGrams": [],
        "height": round(height, 2),
        "width": round(width, 2),
        "filename": os.path.basename(pdf_path)
    }
    
    # 将坐标信息添加到nGrams中
    for coord in coordinates:
        ngram = {
            "words": [{
                "text": coord["text"],
                "left": str(coord["left"]),
                "top": str(coord["top"]),
                "right": str(coord["right"]),
                "bottom": str(coord["bottom"])
            }],
            "parses": {}
        }
        annotation["nGrams"].append(ngram)
    
    return annotation

def process_pdf(pdf_path):
    try:
        # 获取PDF文件的目录和文件名（不含扩展名）
        pdf_dir = os.path.dirname(pdf_path)
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        
        # 获取文本和坐标
        coordinates = get_text_coordinates(pdf_path)
        
        # 提取发票字段
        invoice_data = extract_invoice_fields(coordinates)
        
        # 创建JSON文件路径
        json_path = os.path.join(pdf_dir, f"{pdf_name}.json")
        
        # 保存JSON数据
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(invoice_data, f, ensure_ascii=False, indent=4)
        print(f"Created {json_path}")
        
        # 创建CSV文件路径
        csv_path = os.path.join(pdf_dir, f"{pdf_name}.csv")
        
        # 保存CSV数据
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['字段', '值'])
            for key, value in invoice_data.items():
                if key != 'nGrams':  # 排除nGrams字段
                    writer.writerow([key, value])
        print(f"Created {csv_path}")
        
    except Exception as e:
        print(f"处理文件时出错 {pdf_path}: {str(e)}")

def main():
    """主函数，用于测试单个PDF文件处理"""
    # 设置数据目录
    data_dir = "data"
    
    # 确保数据目录存在
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
    
    # 处理所有PDF文件
    for filename in os.listdir(data_dir):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(data_dir, filename)
            print(f"\nProcessing {pdf_path}...")
            process_pdf(pdf_path)

if __name__ == "__main__":
    main() 