"""
分块工具：把 all_rows.tsv 按指定行数分割成 LLM 可读的纯文本文件。
用法：python chunk_split.py [--rows-per-chunk 200] [--max-chunks 5]
"""

import os
import argparse


def format_chunk(chunk_rows, chunk_id, total_chunks, sources):
    """将一组行格式化为一个 chunk 文本。"""
    lines = []
    lines.append("=" * 60)
    lines.append(f"=== 块编号: {chunk_id}/{total_chunks} ===")
    start_row = (chunk_id - 1) * len(chunk_rows) + 1
    end_row = start_row + len(chunk_rows) - 1
    lines.append(f"行范围: {start_row}-{end_row} (共 {len(chunk_rows)} 行)")
    # 汇总来源
    src_desc = ", ".join(f"{year}_{sheet}({cnt}行)" for (year, sheet), cnt in sorted(sources.items()))
    lines.append(f"包含来源: {src_desc}")
    lines.append("=" * 60)
    lines.append("")

    for i, row in enumerate(chunk_rows):
        lines.append(f"--- [行 {start_row + i}] ---")
        lines.append(f"年份: {row['year']}")
        lines.append(f"国家: {row['country']}  国家代码: {row['country_code']}  Sheet: {row['sheet']}")
        lines.append(f"研究领域(原始): {row['research_raw']}")
        lines.append(f"关键词(原始): {row['keywords_raw']}")
        lines.append("")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="将 TSV 分割为 LLM 可读的文本块")
    parser.add_argument("--rows-per-chunk", type=int, default=200, help="每块行数 (默认 200)")
    parser.add_argument("--max-chunks", type=int, default=0, help="最多生成多少块 (0=全部)")
    parser.add_argument("--tsv", default=None, help="输入 TSV 路径")
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    tsv_path = args.tsv or os.path.join(base_dir, "chunks", "all_rows.tsv")
    out_dir = os.path.join(base_dir, "chunks")
    os.makedirs(out_dir, exist_ok=True)

    # 读取 TSV
    rows = []
    with open(tsv_path, "r", encoding="utf-8") as f:
        header = f.readline().strip().split("\t")
        for line in f:
            parts = line.strip().split("\t")
            if len(parts) < 6:
                continue
            rows.append({
                "year": parts[0],
                "keywords_raw": parts[1],
                "research_raw": parts[2],
                "country": parts[3],
                "country_code": parts[4],
                "sheet": parts[5],
            })

    total_rows = len(rows)
    chunk_size = args.rows_per_chunk
    # 计算需要多少块
    total_chunks = (total_rows + chunk_size - 1) // chunk_size
    if args.max_chunks > 0:
        total_chunks = min(total_chunks, args.max_chunks)

    print(f"总行数: {total_rows}")
    print(f"每块行数: {chunk_size}")
    print(f"块数: {total_chunks}")

    for chunk_id in range(1, total_chunks + 1):
        start = (chunk_id - 1) * chunk_size
        end = min(start + chunk_size, total_rows)
        chunk_rows = rows[start:end]

        # 统计来源
        sources = {}
        for r in chunk_rows:
            key = (r["year"], r["sheet"])
            sources[key] = sources.get(key, 0) + 1

        text = format_chunk(chunk_rows, chunk_id, total_chunks, sources)

        fname = f"chunk_{chunk_id:04d}.txt"
        fpath = os.path.join(out_dir, fname)
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(text)

    print(f"Chunks written to: {out_dir}")
    print(f"命名格式: chunk_0001.txt ~ chunk_{total_chunks:04d}.txt")


if __name__ == "__main__":
    main()
