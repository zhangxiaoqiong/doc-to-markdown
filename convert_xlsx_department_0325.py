"""
Excel转MD - 适配Dify父子分段模式
子段用完整key-value形式，不用表格，避免丢失表头
"""

import pandas as pd
import os
import re
from collections import defaultdict

SEPARATOR = "\n\n---SPLIT---\n\n"


def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).strip()
    if text.startswith('"') and text.endswith('"'):
        text = text[1:-1]
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    return text.strip()


def parse_manager_info(manager_text):
    text = clean_text(manager_text)
    if not text:
        return {"name": "", "id": ""}
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if len(lines) >= 2:
        return {"name": lines[0], "id": lines[1]}
    elif len(lines) == 1:
        return {"name": lines[0], "id": ""}
    return {"name": "", "id": ""}


def format_manager(manager):
    if not manager["name"]:
        return ""
    s = manager["name"]
    if manager["id"]:
        s += f"（工号：{manager['id']}）"
    return s


def read_excel(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception:
        try:
            df = pd.read_excel(file_path, engine='xlrd')
        except Exception:
            df = pd.read_csv(file_path, sep='\t')
    df.columns = [str(c).strip() for c in df.columns]
    column_mapping = {
        '公司': 'company',
        '部门': 'department',
        '业务类别': 'business_category',
        '业务描述': 'business_desc',
        '业务负责人/对接人姓名': 'person_name',
        '业务负责人/对接人工号': 'person_id',
        '部门负责人': 'dept_manager'
    }
    df = df.rename(columns=column_mapping)
    for col in df.columns:
        df[col] = df[col].apply(clean_text)
    return df


# =============================================
# L0 公司级
# =============================================
def generate_company_md(df):
    company_map = defaultdict(lambda: {
        "departments": defaultdict(lambda: {
            "dept_manager": "", "biz_categories": [], "person_count": 0
        })
    })
    person_seen = defaultdict(set)

    for _, row in df.iterrows():
        company = row.get('company', '')
        dept = row.get('department', '')
        d = company_map[company]["departments"][dept]
        if not d["dept_manager"]:
            d["dept_manager"] = row.get('dept_manager', '')
        biz_cat = row.get('business_category', '').replace('\n', '、')
        if biz_cat and biz_cat not in d["biz_categories"]:
            d["biz_categories"].append(biz_cat)
        person_name = re.sub(r'兼$', '', row.get('person_name', ''))
        key = (company, dept)
        if person_name and person_name not in person_seen[key]:
            person_seen[key].add(person_name)
            d["person_count"] += 1

    segments = []
    for company, info in company_map.items():
        lines = []
        lines.append(f"# {company} - 公司组织架构概览")
        lines.append(f"公司名称：{company}")
        lines.append(f"下设部门数量：{len(info['departments'])}个")
        lines.append("")
        for i, (dept, di) in enumerate(info["departments"].items(), 1):
            mgr = format_manager(parse_manager_info(di["dept_manager"]))
            lines.append(f"部门{i}：{dept}，部门负责人：{mgr}，团队规模：{di['person_count']}人，业务范围：{'、'.join(di['biz_categories'])}")
        segments.append('\n'.join(lines))

    return SEPARATOR.join(segments)


# =============================================
# L1 部门级 - 核心改造：不用表格，每人一段完整描述
# =============================================
def generate_dept_md(df):
    dept_map = defaultdict(lambda: {
        "company": "", "dept_manager": "",
        "biz_categories": [], "persons": []
    })
    seen_persons = defaultdict(set)

    for _, row in df.iterrows():
        company = row.get('company', '')
        dept = row.get('department', '')
        key = (company, dept)
        d = dept_map[key]
        if not d["company"]:
            d["company"] = company
        if not d["dept_manager"]:
            d["dept_manager"] = row.get('dept_manager', '')
        biz_cat = row.get('business_category', '').replace('\n', '、')
        if biz_cat and biz_cat not in d["biz_categories"]:
            d["biz_categories"].append(biz_cat)
        person_name = re.sub(r'兼$', '', row.get('person_name', ''))
        person_id = row.get('person_id', '')
        if person_name and person_name not in seen_persons[key]:
            seen_persons[key].add(person_name)
            d["persons"].append((person_name, person_id, biz_cat))

    segments = []
    for (company, dept), info in dept_map.items():
        mgr = format_manager(parse_manager_info(info["dept_manager"]))
        lines = []

        # === 父段头部：部门概要信息 ===
        lines.append(f"# {company} - {dept}")
        lines.append(f"所属公司：{company}")
        lines.append(f"部门名称：{dept}")
        if mgr:
            lines.append(f"部门负责人：{mgr}")
        lines.append(f"团队规模：{len(info['persons'])}人")
        lines.append(f"部门业务范围：{'、'.join(info['biz_categories'])}")
        lines.append("")

        # === 每个人一行完整信息（子段按\n拆分后每条都自包含） ===
        lines.append("部门人员信息如下：")
        for pname, pid, pcat in info["persons"]:
            # 每一行都是完整的、自包含的描述
            pid_str = f"，工号：{pid}" if pid else ""
            lines.append(
                f"{dept}成员：姓名：{pname}{pid_str}，所属部门：{dept}，部门负责人：{mgr}，负责业务类别：{pcat}"
            )

        segments.append('\n'.join(lines))

    return SEPARATOR.join(segments)


# =============================================
# L2 业务级 - 同样不用表格
# =============================================
def generate_business_md(df):
    biz_map = defaultdict(lambda: {
        "company": "", "department": "",
        "dept_manager": "", "persons": []
    })

    for _, row in df.iterrows():
        company = row.get('company', '')
        dept = row.get('department', '')
        biz_cat = row.get('business_category', '')
        key = (company, dept, biz_cat)
        b = biz_map[key]
        if not b["company"]:
            b["company"] = company
        if not b["department"]:
            b["department"] = dept
        if not b["dept_manager"]:
            b["dept_manager"] = row.get('dept_manager', '')
        person_name = re.sub(r'兼$', '', row.get('person_name', ''))
        person_id = row.get('person_id', '')
        biz_desc = row.get('business_desc', '')
        if person_name:
            b["persons"].append((person_name, person_id, biz_desc))

    segments = []
    for (company, dept, biz_cat), info in biz_map.items():
        mgr = format_manager(parse_manager_info(info["dept_manager"]))
        title_cat = biz_cat.replace('\n', '、') if biz_cat else "未分类"
        lines = []

        # === 父段头部 ===
        lines.append(f"# {company} - {dept} - {title_cat}")
        lines.append(f"所属公司：{company}")
        lines.append(f"所属部门：{dept}")
        if mgr:
            lines.append(f"部门负责人：{mgr}")
        lines.append(f"业务类别：{title_cat}")
        lines.append("")

        # === 每人一行完整信息 ===
        lines.append("该业务相关人员信息如下：")
        for pname, pid, pdesc in info["persons"]:
            pid_str = f"，工号：{pid}" if pid else ""
            pdesc_oneline = pdesc.replace('\n', '；') if pdesc else ""
            desc_str = f"，业务描述：{pdesc_oneline}" if pdesc_oneline else ""
            lines.append(
                f"业务类别【{title_cat}】负责人：姓名：{pname}{pid_str}，所属部门：{dept}{desc_str}"
            )

        segments.append('\n'.join(lines))

    return SEPARATOR.join(segments)


# =============================================
# L3 人员级
# =============================================
def generate_person_md(df):
    person_map = defaultdict(lambda: {
        "person_id": "", "company": "",
        "department": "", "dept_manager": "",
        "businesses": []
    })

    for _, row in df.iterrows():
        name = row.get('person_name', '')
        if not name:
            continue
        clean_name = re.sub(r'兼$', '', name)
        p = person_map[clean_name]
        if not p["person_id"] and row.get('person_id', ''):
            p["person_id"] = row['person_id']
        if not p["company"]:
            p["company"] = row.get('company', '')
        if not p["department"]:
            p["department"] = row.get('department', '')
        if not p["dept_manager"]:
            p["dept_manager"] = row.get('dept_manager', '')
        biz_cat = row.get('business_category', '').replace('\n', '、')
        biz_desc = row.get('business_desc', '').replace('\n', '；')
        if biz_cat or biz_desc:
            p["businesses"].append((biz_cat, biz_desc))

    segments = []
    for name, info in person_map.items():
        mgr = format_manager(parse_manager_info(info["dept_manager"]))
        lines = []

        # === 基本信息（父段头部）===
        lines.append(f"# 人员信息 - {name}")
        basic = f"姓名：{name}"
        if info["person_id"]:
            basic += f"，工号：{info['person_id']}"
        if info["company"]:
            basic += f"，所属公司：{info['company']}"
        if info["department"]:
            basic += f"，所属部门：{info['department']}"
        if mgr:
            basic += f"，部门负责人：{mgr}"
        lines.append(basic)

        # === 业务信息（每条业务一行完整描述）===
        if len(info["businesses"]) == 1:
            biz_cat, biz_desc = info["businesses"][0]
            biz_line = f"{name}负责的业务类别：{biz_cat}"
            if biz_desc:
                biz_line += f"，业务描述：{biz_desc}"
            lines.append(biz_line)
        else:
            lines.append(f"{name}负责多项业务，详情如下：")
            for biz_cat, biz_desc in info["businesses"]:
                cat_label = biz_cat if biz_cat else "其他"
                biz_line = f"{name}负责的业务类别：{cat_label}"
                if biz_desc:
                    biz_line += f"，业务描述：{biz_desc}"
                lines.append(biz_line)

        segments.append('\n'.join(lines))

    return SEPARATOR.join(segments)


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Excel转MD - 适配Dify父子分段')
    parser.add_argument('input', help='输入Excel文件路径')
    parser.add_argument('-o', '--output', default='./knowledge_base', help='输出目录')
    parser.add_argument('--layers', default='0,1,2,3', help='生成层级')
    parser.add_argument('--separator', default='---SPLIT---',
                        help='父段分隔符（需与Dify设置一致）')

    args = parser.parse_args()

    global SEPARATOR
    SEPARATOR = f"\n\n{args.separator}\n\n"

    print(f"正在读取: {args.input}")
    df = read_excel(args.input)
    print(f"共读取 {len(df)} 行数据")

    os.makedirs(args.output, exist_ok=True)
    layers = [int(x) for x in args.layers.split(',')]

    generators = {
        0: ("L0_公司级.md", generate_company_md),
        1: ("L1_部门级.md", generate_dept_md),
        2: ("L2_业务级.md", generate_business_md),
        3: ("L3_人员级.md", generate_person_md),
    }

    for layer in layers:
        filename, func = generators[layer]
        content = func(df)
        filepath = os.path.join(args.output, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        seg_count = content.count(args.separator) + 1
        print(f"  {filename} -> {seg_count} 个父段")

    print(f"\n输出目录: {os.path.abspath(args.output)}")


if __name__ == '__main__':
    main()
