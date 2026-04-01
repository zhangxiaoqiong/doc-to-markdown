"""
Excel转MD - 适配Dify父子分段模式
严格按照Excel原始表头字段，各层不遗漏任何信息
原始字段：公司、部门、业务类别、业务描述、业务负责人/对接人姓名、业务负责人/对接人工号、部门负责人
"""

import pandas as pd
import os
import re
from collections import defaultdict

SEPARATOR = "\n\n---SPLIT---\n\n"

# 原始Excel表头（所有层严格使用这些字段名）
F_COMPANY = "公司"
F_DEPT = "部门"
F_BIZ_CAT = "业务类别"
F_BIZ_DESC = "业务描述"
F_PERSON_NAME = "业务负责人/对接人姓名"
F_PERSON_ID = "业务负责人/对接人工号"
F_DEPT_MGR = "部门负责人"


def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).strip()
    if text.startswith('"') and text.endswith('"'):
        text = text[1:-1]
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    return text.strip()


def oneline(text):
    """多行文本转单行，用；连接"""
    return text.replace('\n', '；') if text else ""


def parse_manager_info(manager_text):
    text = clean_text(manager_text)
    if not text:
        return ""
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if len(lines) >= 2:
        return f"{lines[0]}（工号：{lines[1]}）"
    return lines[0] if lines else ""


def read_excel(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception:
        try:
            df = pd.read_excel(file_path, engine='xlrd')
        except Exception:
            df = pd.read_csv(file_path, sep='\t')
    df.columns = [str(c).strip() for c in df.columns]

    # 确保所有必需列存在
    required = [F_COMPANY, F_DEPT, F_BIZ_CAT, F_BIZ_DESC, F_PERSON_NAME, F_PERSON_ID, F_DEPT_MGR]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"缺少必需列: {col}，当前列: {list(df.columns)}")

    for col in df.columns:
        df[col] = df[col].apply(clean_text)
    return df


def clean_person_name(name):
    """去掉姓名后的'兼'字"""
    return re.sub(r'兼$', '', name) if name else ""


# =============================================
# L0 公司级
# =============================================
def generate_company_md(df):
    company_map = defaultdict(lambda: {
        "departments": defaultdict(lambda: {
            "dept_manager_raw": "",
            "biz_categories": [],
            "person_count": 0,
            "persons": []
        })
    })
    person_seen = defaultdict(set)

    for _, row in df.iterrows():
        company = row[F_COMPANY]
        dept = row[F_DEPT]
        d = company_map[company]["departments"][dept]
        if not d["dept_manager_raw"]:
            d["dept_manager_raw"] = row[F_DEPT_MGR]
        biz_cat = oneline(row[F_BIZ_CAT])
        if biz_cat and biz_cat not in d["biz_categories"]:
            d["biz_categories"].append(biz_cat)
        person_name = clean_person_name(row[F_PERSON_NAME])
        key = (company, dept)
        if person_name and person_name not in person_seen[key]:
            person_seen[key].add(person_name)
            d["person_count"] += 1
            d["persons"].append(person_name)

    segments = []
    for company, info in company_map.items():
        lines = []
        lines.append(f"# {company} - 公司组织架构概览")
        lines.append(f"{F_COMPANY}：{company}")
        lines.append(f"下设{F_DEPT}数量：{len(info['departments'])}个")
        lines.append(f"下设{F_DEPT}列表：{'、'.join(info['departments'].keys())}")
        lines.append("")
        for dept, di in info["departments"].items():
            mgr = parse_manager_info(di["dept_manager_raw"])
            lines.append(
                f"{F_COMPANY}：{company}，{F_DEPT}：{dept}，{F_DEPT_MGR}：{mgr}，"
                f"团队规模：{di['person_count']}人，"
                f"{F_BIZ_CAT}范围：{'、'.join(di['biz_categories'])}，"
                f"成员列表：{'、'.join(di['persons'])}"
            )
        segments.append('\n'.join(lines))

    return SEPARATOR.join(segments)


# =============================================
# L1 部门级
# =============================================
def generate_dept_md(df):
    dept_map = defaultdict(lambda: {
        "company": "",
        "dept_manager_raw": "",
        "biz_categories": [],
        "persons": []  # [(姓名, 工号, 业务类别, 业务描述)]
    })
    seen_persons = defaultdict(set)

    for _, row in df.iterrows():
        company = row[F_COMPANY]
        dept = row[F_DEPT]
        key = (company, dept)
        d = dept_map[key]
        if not d["company"]:
            d["company"] = company
        if not d["dept_manager_raw"]:
            d["dept_manager_raw"] = row[F_DEPT_MGR]
        biz_cat = oneline(row[F_BIZ_CAT])
        if biz_cat and biz_cat not in d["biz_categories"]:
            d["biz_categories"].append(biz_cat)

        person_name = clean_person_name(row[F_PERSON_NAME])
        person_id = row[F_PERSON_ID]
        biz_desc = oneline(row[F_BIZ_DESC])

        # 同一人可能负责多个业务，都要记录
        person_key = (person_name, biz_cat)
        if person_name and person_key not in seen_persons[key]:
            seen_persons[key].add(person_key)
            d["persons"].append((person_name, person_id, biz_cat, biz_desc))

    segments = []
    for (company, dept), info in dept_map.items():
        mgr = parse_manager_info(info["dept_manager_raw"])
        lines = []

        # 父段头部
        lines.append(f"# {company} - {dept}")
        lines.append(f"{F_COMPANY}：{company}")
        lines.append(f"{F_DEPT}：{dept}")
        lines.append(f"{F_DEPT_MGR}：{mgr}")
        lines.append(f"团队规模：{len(set(p[0] for p in info['persons']))}人")
        lines.append(f"{F_BIZ_CAT}范围：{'、'.join(info['biz_categories'])}")
        lines.append("")

        # 每人每业务一行，7个字段全带上
        lines.append(f"{dept}人员信息如下：")
        for pname, pid, pcat, pdesc in info["persons"]:
            lines.append(
                f"{F_COMPANY}：{company}，{F_DEPT}：{dept}，{F_DEPT_MGR}：{mgr}，"
                f"{F_PERSON_NAME}：{pname}，{F_PERSON_ID}：{pid}，"
                f"{F_BIZ_CAT}：{pcat}，{F_BIZ_DESC}：{pdesc}"
            )

        segments.append('\n'.join(lines))

    return SEPARATOR.join(segments)


# =============================================
# L2 业务级
# =============================================
def generate_business_md(df):
    biz_map = defaultdict(lambda: {
        "company": "",
        "department": "",
        "dept_manager_raw": "",
        "persons": []  # [(姓名, 工号, 业务描述)]
    })

    for _, row in df.iterrows():
        company = row[F_COMPANY]
        dept = row[F_DEPT]
        biz_cat = row[F_BIZ_CAT]
        key = (company, dept, biz_cat)
        b = biz_map[key]
        if not b["company"]:
            b["company"] = company
        if not b["department"]:
            b["department"] = dept
        if not b["dept_manager_raw"]:
            b["dept_manager_raw"] = row[F_DEPT_MGR]

        person_name = clean_person_name(row[F_PERSON_NAME])
        person_id = row[F_PERSON_ID]
        biz_desc = oneline(row[F_BIZ_DESC])
        if person_name:
            b["persons"].append((person_name, person_id, biz_desc))

    segments = []
    for (company, dept, biz_cat), info in biz_map.items():
        mgr = parse_manager_info(info["dept_manager_raw"])
        title_cat = oneline(biz_cat) if biz_cat else "未分类"
        lines = []

        # 父段头部
        lines.append(f"# {company} - {dept} - {title_cat}")
        lines.append(f"{F_COMPANY}：{company}")
        lines.append(f"{F_DEPT}：{dept}")
        lines.append(f"{F_DEPT_MGR}：{mgr}")
        lines.append(f"{F_BIZ_CAT}：{title_cat}")
        lines.append("")

        # 每人一行，7个字段全带上
        lines.append(f"该{F_BIZ_CAT}【{title_cat}】相关人员信息如下：")
        for pname, pid, pdesc in info["persons"]:
            lines.append(
                f"{F_COMPANY}：{company}，{F_DEPT}：{dept}，{F_DEPT_MGR}：{mgr}，"
                f"{F_BIZ_CAT}：{title_cat}，"
                f"{F_PERSON_NAME}：{pname}，{F_PERSON_ID}：{pid}，"
                f"{F_BIZ_DESC}：{pdesc}"
            )

        segments.append('\n'.join(lines))

    return SEPARATOR.join(segments)


# =============================================
# L3 人员级
# =============================================
def generate_person_md(df):
    person_map = defaultdict(lambda: {
        "person_id": "",
        "company": "",
        "department": "",
        "dept_manager_raw": "",
        "rows": []  # [(业务类别, 业务描述)] - 保留原始每行
    })

    for _, row in df.iterrows():
        name = row[F_PERSON_NAME]
        if not name:
            continue
        clean_name = clean_person_name(name)
        p = person_map[clean_name]
        if not p["person_id"] and row[F_PERSON_ID]:
            p["person_id"] = row[F_PERSON_ID]
        if not p["company"]:
            p["company"] = row[F_COMPANY]
        if not p["department"]:
            p["department"] = row[F_DEPT]
        if not p["dept_manager_raw"]:
            p["dept_manager_raw"] = row[F_DEPT_MGR]

        biz_cat = oneline(row[F_BIZ_CAT])
        biz_desc = oneline(row[F_BIZ_DESC])
        p["rows"].append((biz_cat, biz_desc))

    segments = []
    for name, info in person_map.items():
        mgr = parse_manager_info(info["dept_manager_raw"])
        lines = []

        # 父段头部：基本信息，每个字段独占一行
        lines.append(f"# 人员信息 - {name}")
        lines.append(f"{F_PERSON_NAME}：{name}")
        lines.append(f"{F_PERSON_ID}：{info['person_id']}")
        lines.append(f"{F_COMPANY}：{info['company']}")
        lines.append(f"{F_DEPT}：{info['department']}")
        lines.append(f"{F_DEPT_MGR}：{mgr}")
        lines.append("")

        # 业务信息：每条业务一行，带全部7个字段
        if len(info["rows"]) == 1:
            biz_cat, biz_desc = info["rows"][0]
            lines.append(
                f"{F_COMPANY}：{info['company']}，{F_DEPT}：{info['department']}，{F_DEPT_MGR}：{mgr}，"
                f"{F_PERSON_NAME}：{name}，{F_PERSON_ID}：{info['person_id']}，"
                f"{F_BIZ_CAT}：{biz_cat}，{F_BIZ_DESC}：{biz_desc}"
            )
        else:
            lines.append(f"{name}负责多项业务，详情如下：")
            for biz_cat, biz_desc in info["rows"]:
                lines.append(
                    f"{F_COMPANY}：{info['company']}，{F_DEPT}：{info['department']}，{F_DEPT_MGR}：{mgr}，"
                    f"{F_PERSON_NAME}：{name}，{F_PERSON_ID}：{info['person_id']}，"
                    f"{F_BIZ_CAT}：{biz_cat}，{F_BIZ_DESC}：{biz_desc}"
                )

        segments.append('\n'.join(lines))

    return SEPARATOR.join(segments)


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Excel转MD - 适配Dify父子分段')
    parser.add_argument('input', help='输入Excel文件路径')
    parser.add_argument('-o', '--output', default='./knowledge_base', help='输出目录')
    parser.add_argument('--layers', default='0,1,2,3', help='生成层级(0=公司,1=部门,2=业务,3=人员)')
    parser.add_argument('--separator', default='---SPLIT---',
                        help='父段分隔符（需与Dify设置一致）')

    args = parser.parse_args()

    global SEPARATOR
    SEPARATOR = f"\n\n{args.separator}\n\n"

    print(f"正在读取: {args.input}")
    df = read_excel(args.input)
    print(f"共读取 {len(df)} 行数据")
    print(f"列名: {list(df.columns)}")

    os.makedirs(args.output, exist_ok=True)
    layers = [int(x) for x in args.layers.split(',')]

    generators = {
        0: ("L0_公司级.md", generate_company_md),
        1: ("L1_部门级.md", generate_dept_md),
        2: ("L2_业务级.md", generate_business_md),
        3: ("L3_人员级.md", generate_person_md),
    }

    total = 0
    for layer in layers:
        filename, func = generators[layer]
        content = func(df)
        filepath = os.path.join(args.output, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        seg_count = content.count(args.separator) + 1
        total += seg_count
        print(f"  {filename} -> {seg_count} 个父段")

    print(f"\n共生成 {total} 个父段")
    print(f"输出目录: {os.path.abspath(args.output)}")
    print(f"\nDify配置：父段分隔符={args.separator}，子段分隔符=\\n")


if __name__ == '__main__':
    main()
