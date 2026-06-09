"""
上下文/记忆测试工具。
测试 LLM 在不同块大小下的表现：
  1. 行覆盖率（是否所有行都被处理）
  2. 关键词拆分一致性（相同行在不同块大小下的结果是否一致）
  3. 研究领域分类准确率
  4. needle-in-haystack 记忆测试
"""

import json
import os
import argparse
from collections import defaultdict


def load_result(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_record_index(data):
    """构建 {row_number: record} 索引。"""
    return {rec["row"]: rec for rec in data.get("records", [])}


def compare_two_results(ref_data, test_data, label):
    """对比测试结果与参考结果的差异。"""
    ref_idx = build_record_index(ref_data)
    test_idx = build_record_index(test_data)

    ref_rows = set(ref_idx.keys())
    test_rows = set(test_idx.keys())

    missing = ref_rows - test_rows  # 参考中有但测试中缺失的
    extra = test_rows - ref_rows    # 测试中多出的

    report = [f"\n=== {label} ==="]
    report.append(f"参考行数: {len(ref_rows)}, 测试行数: {len(test_rows)}")
    report.append(f"缺失行: {len(missing)} ({sorted(missing)[:20]}{'...' if len(missing) > 20 else ''})")
    report.append(f"多余行: {len(extra)}")

    # 逐行比较关键词
    kw_diff = 0
    kw_total = 0
    cat_diff = 0
    cat_total = 0
    detail = []

    for row_id in sorted(ref_rows & test_rows):
        ref_rec = ref_idx[row_id]
        test_rec = test_idx[row_id]

        ref_kw = set(ref_rec.get("keywords", []))
        test_kw = set(test_rec.get("keywords", []))
        kw_total += len(ref_kw)
        kw_delta = len(ref_kw - test_kw) + len(test_kw - ref_kw)
        if kw_delta > 0:
            kw_diff += kw_delta
            if len(detail) < 10:
                detail.append(
                    f"  行{row_id} 关键词差异: ref={ref_kw - test_kw} extra={test_kw - ref_kw}"
                )

        ref_cat = set(ref_rec.get("broad_categories", []))
        test_cat = set(test_rec.get("broad_categories", []))
        cat_total += 1
        if ref_cat != test_cat:
            cat_diff += 1
            if len(detail) < 10:
                detail.append(
                    f"  行{row_id} 大类差异: ref={ref_cat} test={test_cat}"
                )

    kw_acc = 1 - (kw_diff / max(kw_total, 1))
    cat_acc = 1 - (cat_diff / max(cat_total, 1))
    coverage = len(ref_rows & test_rows) / max(len(ref_rows), 1)

    report.append(f"行覆盖率: {coverage:.1%}")
    report.append(f"关键词准确率: {kw_acc:.1%} (差异 {kw_diff}/{kw_total})")
    report.append(f"大类准确率: {cat_acc:.1%} (差异 {cat_diff}/{cat_total})")
    report.append(f"\n差异详情 (最多10条):")
    report.extend(detail)

    return {
        "coverage": coverage,
        "keyword_accuracy": kw_acc,
        "category_accuracy": cat_acc,
        "report": "\n".join(report),
    }


def needle_in_haystack_test(template_path, output_dir):
    """生成 needle-in-haystack 测试块：在块开头放独特 token，验证 LLM 是否记得。"""
    needle = "NEEDLE_TOKEN_XY7ZQ_MUST_BE_IN_OUTPUT"
    # 这里只输出提示，实际测试需要配合 LLM
    print(f"""
=== Needle-in-Haystack 测试说明 ===
1. 在测试块的开头添加一条特殊记录：
   {needle}
2. 将该记录的关键词字段设为 "{needle}"
3. 将 block 大小设为不同值（如 50、100、200、400 行）
4. 检查 LLM 输出中是否包含该 needle token
5. 如果包含 → 该块大小下 LLM 记忆完好
   如果不包含 → 已超出有效上下文/记忆范围
""")


def main():
    parser = argparse.ArgumentParser(description="LLM 上下文长度与准确性测试")
    parser.add_argument("--ref", default=None, help="参考结果 JSON（小块的黄金标准）")
    parser.add_argument("--test", default=None, help="测试结果 JSON（大块的结果）")
    parser.add_argument("--ref-dir", default=None, help="批量比较：参考结果目录")
    parser.add_argument("--test-dir", default=None, help="批量比较：测试结果目录")
    parser.add_argument("--gen-needle", action="store_true", help="生成 needle-in-haystack 测试提示")
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))

    if args.gen_needle:
        needle_in_haystack_test(None, None)
        return

    if args.ref and args.test:
        ref = load_result(args.ref)
        test = load_result(args.test)
        result = compare_two_results(ref, test, f"{args.ref} vs {args.test}")
        print(result["report"])
        return

    if args.ref_dir and args.test_dir:
        ref_files = sorted(f for f in os.listdir(args.ref_dir) if f.endswith(".json"))
        test_files = sorted(f for f in os.listdir(args.test_dir) if f.endswith(".json"))
        # 按 chunk_id 匹配
        summary = []
        for rf in ref_files:
            # 找对应 test file（按行范围匹配）
            cid = rf.replace("_result.json", "")
            tf = None
            for t in test_files:
                if cid in t:
                    tf = t
                    break
            if not tf:
                # 尝试行号匹配
                ref_data = load_result(os.path.join(args.ref_dir, rf))
                ref_rows = {r["row"] for r in ref_data.get("records", [])}
                for t in test_files:
                    test_data = load_result(os.path.join(args.test_dir, t))
                    test_rows = {r["row"] for r in test_data.get("records", [])}
                    if ref_rows & test_rows:
                        tf = t
                        break
            if tf:
                ref_data = load_result(os.path.join(args.ref_dir, rf))
                test_data = load_result(os.path.join(args.test_dir, tf))
                result = compare_two_results(ref_data, test_data, f"{rf} vs {tf}")
                summary.append(result)
                print(result["report"])
            else:
                print(f"[SKIP] {rf}: no matching test file found")

        # 汇总
        if summary:
            print("\n=== 汇总 ===")
            avg_cov = sum(s["coverage"] for s in summary) / len(summary)
            avg_kw = sum(s["keyword_accuracy"] for s in summary) / len(summary)
            avg_cat = sum(s["category_accuracy"] for s in summary) / len(summary)
            print(f"平均行覆盖率: {avg_cov:.1%}")
            print(f"平均关键词准确率: {avg_kw:.1%}")
            print(f"平均大类准确率: {avg_cat:.1%}")
        return

    parser.print_help()


if __name__ == "__main__":
    main()
