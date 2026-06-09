"""
DeepSeek API 调用工具。
读取 prompt_template.txt + chunk 数据 --> 调 DeepSeek API --> 保存 JSON 结果。

用法：
  python call_deepseek.py --api-key sk-xxx --start 1 --end 5
  或设置环境变量 DEEPSEEK_API_KEY

可选参数：
  --model deepseek-chat       (默认)
  --base-url https://api.deepseek.com  (默认)
  --chunks-dir chunks         (数据块目录)
  --prompt prompt_template.txt (提示词文件)
"""
import json, os, time, argparse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def load_prompt(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def extract_json(text):
    """从 LLM 返回文本中提取 JSON，自动修复常见截断。"""
    text = text.strip()
    start = text.find('{')
    end = text.rfind('}')
    if start >= 0 and end > start:
        text = text[start:end+1]
    text = text.strip()
    if text.endswith('"'):
        text += ']}'
    if not text.endswith('}'):
        depth = sum(1 for ch in text if ch == '{') - sum(1 for ch in text if ch == '}')
        text += ']' if depth > 0 else ''
        while depth > 0:
            text += '}'
            depth -= 1
    return text


def call_deepseek(chunk_text, chunk_id, api_key, base_url, model, chunks_dir):
    """调用 DeepSeek API，返回解析后的 JSON 和性能指标。"""
    import requests

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": chunk_text},
        ],
        "temperature": 0.1,
        "max_tokens": 32768,
    }

    t0 = time.time()
    metrics = {"chunk_id": chunk_id, "success": False, "attempts": 0,
               "time_seconds": 0, "input_tokens": 0, "output_tokens": 0,
               "output_chars": 0, "finish_reason": ""}

    for attempt in range(3):
        metrics["attempts"] = attempt + 1
        try:
            resp = requests.post(
                f"{base_url}/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=300,
            )
            elapsed = time.time() - t0
            metrics["time_seconds"] = round(elapsed, 1)

            if resp.status_code != 200:
                print(f"  HTTP {resp.status_code}: {resp.text[:200]}")
                if attempt < 2:
                    time.sleep(3)
                continue

            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            finish = data["choices"][0].get("finish_reason", "?")
            usage = data.get("usage", {})

            metrics["finish_reason"] = finish
            metrics["input_tokens"] = usage.get("prompt_tokens", 0)
            metrics["output_tokens"] = usage.get("completion_tokens", 0)
            metrics["output_chars"] = len(content)

            # 保存原始响应用于调试
            debug_path = os.path.join(chunks_dir, f"chunk_{chunk_id:04d}_raw.txt")
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(content)

            json_str = extract_json(content)
            try:
                result = json.loads(json_str)
                metrics["success"] = True
                n_recs = len(result.get("r") or result.get("records", []))
                print(f"  OK | {elapsed:.0f}s | in:{usage.get('prompt_tokens',0)} out:{usage.get('completion_tokens',0)} | finish:{finish} | {n_recs} rows")
                return result, metrics
            except json.JSONDecodeError as e:
                err_path = os.path.join(chunks_dir, f"chunk_{chunk_id:04d}_error.json")
                with open(err_path, "w", encoding="utf-8") as f:
                    f.write(json_str)
                print(f"  JSON解析失败: {e} (line {e.lineno}) | finish:{finish}")
                loc = e.pos if hasattr(e, 'pos') else 0
                print(f"  错误位置前后: ...{json_str[max(0,loc-50):loc+50]}...")
                if attempt < 2:
                    time.sleep(2)

        except requests.exceptions.Timeout:
            print(f"  超时，重试...")
            time.sleep(5)
        except Exception as e:
            print(f"  异常: {e}")
            if attempt < 2:
                time.sleep(3)

    return None, metrics


def main():
    parser = argparse.ArgumentParser(description="DeepSeek API 批量处理 WoS 数据块")
    parser.add_argument("--api-key", default=os.environ.get("DEEPSEEK_API_KEY", ""),
                        help="DeepSeek API key（也可设置环境变量 DEEPSEEK_API_KEY）")
    parser.add_argument("--base-url", default="https://api.deepseek.com")
    parser.add_argument("--model", default="deepseek-chat")
    parser.add_argument("--prompt", default="prompt_template.txt", help="提示词模板文件")
    parser.add_argument("--chunks-dir", default="chunks", help="数据块目录")
    parser.add_argument("--start", type=int, required=True, help="起始块号")
    parser.add_argument("--end", type=int, required=True, help="结束块号")
    args = parser.parse_args()

    if not args.api_key:
        print("错误: 请设置 DEEPSEEK_API_KEY 环境变量或用 --api-key 参数")
        print("用法: python call_deepseek.py --api-key sk-xxx --start 1 --end 5")
        return

    global SYSTEM_PROMPT
    SYSTEM_PROMPT = load_prompt(os.path.join(BASE_DIR, args.prompt))
    chunks_dir = os.path.join(BASE_DIR, args.chunks_dir)
    all_metrics = []
    total_rows = 0

    for cid in range(args.start, args.end + 1):
        chunk_file = os.path.join(chunks_dir, f"chunk_{cid:04d}.txt")
        result_file = os.path.join(chunks_dir, f"chunk_{cid:04d}_result.json")

        if not os.path.exists(chunk_file):
            continue

        if os.path.exists(result_file):
            with open(result_file, "r", encoding="utf-8") as f:
                existing = json.load(f)
            n = existing.get("total_rows", len(existing.get("r", existing.get("records", []))))
            total_rows += n
            print(f"chunk_{cid:04d}: [EXISTS] {n} rows (跳过)")
            continue

        print(f"\n--- chunk_{cid:04d} ({cid}/{args.end}) ---")
        with open(chunk_file, "r", encoding="utf-8") as f:
            chunk_text = f.read()

        result, metrics = call_deepseek(chunk_text, cid, args.api_key, args.base_url, args.model, chunks_dir)
        all_metrics.append(metrics)

        if result is None:
            print(f"  FAILED: chunk_{cid:04d}")
            continue

        records = result.get("r") or result.get("records", [])
        result["total_rows"] = len(records)

        with open(result_file, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        total_rows += len(records)

    # 效率报告
    print(f"\n{'='*60}")
    print(f"效率报告: {args.start}-{args.end} 块")
    print(f"{'='*60}")
    successful = [m for m in all_metrics if m["success"]]
    failed = [m for m in all_metrics if not m["success"]]
    print(f"成功: {len(successful)} 块, 失败: {len(failed)} 块")
    print(f"累计行数: {total_rows}")
    if successful:
        times = [m["time_seconds"] for m in successful]
        ins = [m["input_tokens"] for m in successful]
        outs = [m["output_tokens"] for m in successful]
        truns = sum(1 for m in successful if m["finish_reason"] == "length")
        print(f"平均耗时: {sum(times)/len(times):.0f}s (最快:{min(times):.0f}s, 最慢:{max(times):.0f}s)")
        print(f"平均输入: {sum(ins)/len(ins):.0f} tokens")
        print(f"平均输出: {sum(outs)/len(outs):.0f} tokens")
        print(f"truncated (finish_reason=length): {truns}/{len(successful)}")
    for m in all_metrics:
        status = "OK" if m["success"] else "FAIL"
        print(f"  chunk_{m['chunk_id']:04d}: {status} {m['time_seconds']}s in:{m['input_tokens']} out:{m['output_tokens']} finish:{m['finish_reason']}")


if __name__ == "__main__":
    main()
