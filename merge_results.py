"""
合并工具：读取所有 LLM chunk 输出 JSON，本地做 小类→大类 和 国家→代码 映射。
输出：
  output/词频统计结果.txt  — 层级文本（大类→国家→年份→关键词:频次）
  output/原始数据映射表.xlsx — 逐条映射表
"""

import os, json, re, argparse
from collections import defaultdict

# ── 154 WoS Research Area → 5 大类（本地映射）──
AREA_TO_BROAD = {
    "Architecture": "Arts & Humanities","Art": "Arts & Humanities","Arts & Humanities Other Topics": "Arts & Humanities","Asian Studies": "Arts & Humanities","Classics": "Arts & Humanities","Dance": "Arts & Humanities","Film, Radio & Television": "Arts & Humanities","History": "Arts & Humanities","History & Philosophy of Science": "Arts & Humanities","Literature": "Arts & Humanities","Music": "Arts & Humanities","Philosophy": "Arts & Humanities","Religion": "Arts & Humanities","Theater": "Arts & Humanities",
    "Agriculture": "Life Sciences & Biomedicine","Allergy": "Life Sciences & Biomedicine","Anatomy & Morphology": "Life Sciences & Biomedicine","Anesthesiology": "Life Sciences & Biomedicine","Anthropology": "Life Sciences & Biomedicine","Audiology & Speech-Language Pathology": "Life Sciences & Biomedicine","Behavioral Sciences": "Life Sciences & Biomedicine","Biochemistry & Molecular Biology": "Life Sciences & Biomedicine","Biodiversity & Conservation": "Life Sciences & Biomedicine","Biophysics": "Life Sciences & Biomedicine","Biotechnology & Applied Microbiology": "Life Sciences & Biomedicine","Cardiovascular System & Cardiology": "Life Sciences & Biomedicine","Cell Biology": "Life Sciences & Biomedicine","Critical Care Medicine": "Life Sciences & Biomedicine","Dentistry, Oral Surgery & Medicine": "Life Sciences & Biomedicine","Dermatology": "Life Sciences & Biomedicine","Developmental Biology": "Life Sciences & Biomedicine","Emergency Medicine": "Life Sciences & Biomedicine","Endocrinology & Metabolism": "Life Sciences & Biomedicine","Entomology": "Life Sciences & Biomedicine","Environmental Sciences & Ecology": "Life Sciences & Biomedicine","Evolutionary Biology": "Life Sciences & Biomedicine","Fisheries": "Life Sciences & Biomedicine","Food Science & Technology": "Life Sciences & Biomedicine","Forestry": "Life Sciences & Biomedicine","Gastroenterology & Hepatology": "Life Sciences & Biomedicine","General & Internal Medicine": "Life Sciences & Biomedicine","Genetics & Heredity": "Life Sciences & Biomedicine","Geriatrics & Gerontology": "Life Sciences & Biomedicine","Health Care Sciences & Services": "Life Sciences & Biomedicine","Hematology": "Life Sciences & Biomedicine","Immunology": "Life Sciences & Biomedicine","Infectious Diseases": "Life Sciences & Biomedicine","Integrative & Complementary Medicine": "Life Sciences & Biomedicine","Legal Medicine": "Life Sciences & Biomedicine","Life Sciences Biomedicine Other Topics": "Life Sciences & Biomedicine","Marine & Freshwater Biology": "Life Sciences & Biomedicine","Mathematical & Computational Biology": "Life Sciences & Biomedicine","Medical Ethics": "Life Sciences & Biomedicine","Medical Informatics": "Life Sciences & Biomedicine","Medical Laboratory Technology": "Life Sciences & Biomedicine","Microbiology": "Life Sciences & Biomedicine","Mycology": "Life Sciences & Biomedicine","Neurosciences & Neurology": "Life Sciences & Biomedicine","Nursing": "Life Sciences & Biomedicine","Nutrition & Dietetics": "Life Sciences & Biomedicine","Obstetrics & Gynecology": "Life Sciences & Biomedicine","Oncology": "Life Sciences & Biomedicine","Ophthalmology": "Life Sciences & Biomedicine","Orthopedics": "Life Sciences & Biomedicine","Otorhinolaryngology": "Life Sciences & Biomedicine","Paleontology": "Life Sciences & Biomedicine","Parasitology": "Life Sciences & Biomedicine","Pathology": "Life Sciences & Biomedicine","Pediatrics": "Life Sciences & Biomedicine","Pharmacology & Pharmacy": "Life Sciences & Biomedicine","Physiology": "Life Sciences & Biomedicine","Plant Sciences": "Life Sciences & Biomedicine","Psychiatry": "Life Sciences & Biomedicine","Public, Environmental & Occupational Health": "Life Sciences & Biomedicine","Radiology, Nuclear Medicine & Medical Imaging": "Life Sciences & Biomedicine","Rehabilitation": "Life Sciences & Biomedicine","Reproductive Biology": "Life Sciences & Biomedicine","Research & Experimental Medicine": "Life Sciences & Biomedicine","Respiratory System": "Life Sciences & Biomedicine","Rheumatology": "Life Sciences & Biomedicine","Sport Sciences": "Life Sciences & Biomedicine","Substance Abuse": "Life Sciences & Biomedicine","Surgery": "Life Sciences & Biomedicine","Toxicology": "Life Sciences & Biomedicine","Transplantation": "Life Sciences & Biomedicine","Tropical Medicine": "Life Sciences & Biomedicine","Urology & Nephrology": "Life Sciences & Biomedicine","Veterinary Sciences": "Life Sciences & Biomedicine","Virology": "Life Sciences & Biomedicine","Zoology": "Life Sciences & Biomedicine",
    "Astronomy & Astrophysics": "Physical Sciences","Chemistry": "Physical Sciences","Crystallography": "Physical Sciences","Electrochemistry": "Physical Sciences","Geochemistry & Geophysics": "Physical Sciences","Geology": "Physical Sciences","Mathematics": "Physical Sciences","Meteorology & Atmospheric Sciences": "Physical Sciences","Mineralogy": "Physical Sciences","Mining & Mineral Processing": "Physical Sciences","Oceanography": "Physical Sciences","Optics": "Physical Sciences","Physical Geography": "Physical Sciences","Physics": "Physical Sciences","Polymer Science": "Physical Sciences","Thermodynamics": "Physical Sciences","Water Resources": "Physical Sciences",
    "Archaeology": "Social Sciences","Area Studies": "Social Sciences","Biomedical Social Sciences": "Social Sciences","Business & Economics": "Social Sciences","Communication": "Social Sciences","Criminology & Penology": "Social Sciences","Cultural Studies": "Social Sciences","Demography": "Social Sciences","Development Studies": "Social Sciences","Education & Educational Research": "Social Sciences","Ethnic Studies": "Social Sciences","Family Studies": "Social Sciences","Geography": "Social Sciences","Government & Law": "Social Sciences","International Relations": "Social Sciences","Linguistics": "Social Sciences","Mathematical Methods In Social Sciences": "Social Sciences","Psychology": "Social Sciences","Public Administration": "Social Sciences","Social Issues": "Social Sciences","Social Sciences Other Topics": "Social Sciences","Social Work": "Social Sciences","Sociology": "Social Sciences","Urban Studies": "Social Sciences","Women's Studies": "Social Sciences",
    "Acoustics": "Technology","Automation & Control Systems": "Technology","Computer Science": "Technology","Construction & Building Technology": "Technology","Energy & Fuels": "Technology","Engineering": "Technology","Imaging Science & Photographic Technology": "Technology","Information Science & Library Science": "Technology","Instruments & Instrumentation": "Technology","Materials Science": "Technology","Mechanics": "Technology","Metallurgy & Metallurgical Engineering": "Technology","Microscopy": "Technology","Nuclear Science & Technology": "Technology","Operations Research & Management Science": "Technology","Remote Sensing": "Technology","Robotics": "Technology","Science & Technology Other Topics": "Technology","Spectroscopy": "Technology","Telecommunications": "Technology","Transportation": "Technology",
}

# 加载国家名→代码映射
def load_country_map(tsv_path):
    cmap = {}
    with open(tsv_path, "r", encoding="utf-8") as f:
        f.readline()
        for line in f:
            parts = line.strip().split("\t")
            if len(parts) >= 5:
                n, c = parts[3].strip(), parts[4].strip()
                if n and c:
                    cmap[n] = c
    return cmap


def load_chunk_results(chunks_dir):
    results = []
    files = sorted(f for f in os.listdir(chunks_dir) if f.endswith("_result.json"))
    for fname in files:
        with open(os.path.join(chunks_dir, fname), "r", encoding="utf-8") as f:
            data = json.load(f)
        records = data.get("r") or data.get("records", [])
        results.append({"file": fname, "data": data})
        print(f"  Loaded {fname}: {len(records)} records")
    return results


def area_to_broad(area_name):
    """研究领域名→大类，处理逗号差异。"""
    # 直接匹配
    if area_name in AREA_TO_BROAD:
        return AREA_TO_BROAD[area_name]
    # 去逗号匹配
    flat = area_name.replace(",", "").replace("  ", " ")
    for orig, broad in AREA_TO_BROAD.items():
        if orig.replace(",", "").replace("  ", " ") == flat:
            return broad
    # 连字符变体（Data Citation Index: "Arts & Humanities - Other Topics"）
    if " - Other Topics" in area_name:
        fixed = area_name.replace(" - Other Topics", " Other Topics")
        if fixed in AREA_TO_BROAD:
            return AREA_TO_BROAD[fixed]
    if "Science & Technology" in area_name or "Science Technology" in area_name:
        return "Technology"
    return None


def merge_and_aggregate(results, country_map):
    all_records = []
    freq_map = defaultdict(int)
    unknown_areas = defaultdict(int)

    for r in results:
        data = r["data"]
        records = data.get("r") or data.get("records", [])

        for rec in records:
            i  = rec.get("i") or rec.get("row", "")
            y  = rec.get("y") or rec.get("year", "")
            n  = rec.get("n") or rec.get("country", "")
            s  = rec.get("s") or rec.get("sheet", "")
            a  = rec.get("a") or rec.get("research_areas", [])
            k  = rec.get("k") or rec.get("keywords", [])
            cc = country_map.get(n, "")

            # 本地做小类→大类映射
            broad_cats = []
            for area in a:
                broad = area_to_broad(area)
                if broad:
                    broad_cats.append(broad)
                else:
                    unknown_areas[area] += 1
            broad_cats = list(dict.fromkeys(broad_cats))  # 去重保序

            rec_norm = {
                "row": i, "year": y, "country": n, "country_code": cc,
                "sheet": s, "broad_categories": broad_cats,
                "research_areas": a, "keywords": k,
            }
            all_records.append(rec_norm)

            # 统计频率
            if n and broad_cats:
                for cat in broad_cats:
                    for kw in k:
                        kw_clean = kw.strip()
                        if kw_clean:
                            freq_map[(cat, n, str(y), kw_clean)] += 1

    if unknown_areas:
        print(f"\n  [WARN] 未匹配的研究领域 ({len(unknown_areas)}种):")
        for area, cnt in sorted(unknown_areas.items(), key=lambda x: -x[1])[:10]:
            print(f"    \"{area}\": {cnt}次")

    return all_records, freq_map


def write_txt_output(freq_map, output_path):
    tree = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
    for (cat, country, year, kw), count in freq_map.items():
        tree[cat][country][year][kw] = count

    category_order = [
        "Arts & Humanities", "Life Sciences & Biomedicine",
        "Physical Sciences", "Social Sciences", "Technology",
    ]

    lines = []
    for cat in category_order:
        if cat not in tree:
            matched = [c for c in tree if cat.lower() in c.lower()]
            if not matched:
                continue
            cat = matched[0]
        lines.append(f"【{cat}】")
        for country in sorted(tree[cat].keys()):
            if not country:
                continue
            lines.append(f"({country})")
            years = tree[cat][country]
            for year in sorted(years.keys()):
                lines.append(f"({year})")
                for kw in sorted(years[year].keys()):
                    lines.append(f"{kw}: {years[year][kw]}")
            lines.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"TXT output: {output_path} ({len(lines)} lines)")


def write_mapping_excel(all_records, output_path):
    try:
        import openpyxl
    except ImportError:
        print("  [WARN] openpyxl not available, skipping Excel")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "映射表"
    ws.append(["行号","年份","国家","国家代码","Sheet","大类","研究领域(WoS)","关键词(拆分后)"])
    for rec in all_records:
        ws.append([
            rec["row"], rec["year"], rec["country"], rec["country_code"], rec["sheet"],
            "; ".join(rec["broad_categories"]), "; ".join(rec["research_areas"]),
            "; ".join(rec["keywords"]),
        ])
    for col, w in [("A",8),("B",8),("C",22),("D",10),("E",12),("F",28),("G",50),("H",70)]:
        ws.column_dimensions[col].width = w
    wb.save(output_path)
    print(f"Excel mapping: {output_path} ({len(all_records)} rows)")


def main():
    p = argparse.ArgumentParser(description="合并 LLM chunk 输出")
    p.add_argument("--chunks-dir", default=None); p.add_argument("--output-dir", default=None)
    args = p.parse_args()
    base = os.path.dirname(os.path.abspath(__file__))
    chunks_dir = args.chunks_dir or os.path.join(base, "chunks")
    out_dir = args.output_dir or os.path.join(base, "output")
    os.makedirs(out_dir, exist_ok=True)

    tsv = os.path.join(base, "chunks", "all_rows.tsv")
    country_map = load_country_map(tsv)
    print(f"国家映射: {len(country_map)} 条")

    results = load_chunk_results(chunks_dir)
    if not results:
        print("[ERROR] No *_result.json found.")
        return

    all_records, freq_map = merge_and_aggregate(results, country_map)
    print(f"合并: {len(all_records)} 条记录, {len(freq_map)} 个词频条目")

    write_txt_output(freq_map, os.path.join(out_dir, "词频统计结果.txt"))
    write_mapping_excel(all_records, os.path.join(out_dir, "原始数据映射表.xlsx"))
    print("Done.")


if __name__ == "__main__":
    main()
