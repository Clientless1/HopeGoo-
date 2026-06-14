# -*- coding: utf-8 -*-
"""
自动化脚本：监控 Reqable 抓包 → 自动刷新 token → 批量抓取 → 出文档。
用法: .venv/bin/python3 auto_run.py 福州,上海
     在手机上 App 搜任意城市，工具自动检测到新鲜 token 后立刻开始抓。
"""
import sys, time, json, os, re, base64
from datetime import datetime
import hopegoo_api as api
import hopegoo_gcoins as g

REQABLE_DB = os.path.expanduser("~/Library/Application Support/com.reqable.macosx/box/data.mdb")

def get_latest_ts():
    """获取 LMDB 中最新的 hotel.reshg.com list 请求的时间戳。"""
    last = 0
    try:
        raw = open(REQABLE_DB, "rb").read()
        for blk in re.findall(rb"H4sI[A-Za-z0-9+/=]{80,}", raw):
            try:
                d = json.loads(gzip.decompress(base64.b64decode(blk + b"==" * (-len(blk) % 4))).decode("utf-8", "replace"))
            except Exception:
                continue
            if not isinstance(d, dict):
                continue
            s = d.get("session", {})
            if s.get("connection", {}).get("originHost", "") != "hotel.reshg.com":
                continue
            p = s.get("request", {}).get("requestLine", {}).get("path", "")
            if "/tapi/v2/list?" not in p:
                continue
            ts = s.get("timestamp", 0)
            if ts > last:
                last = ts
    except Exception:
        pass
    return last

def main():
    if len(sys.argv) < 2:
        print("用法: .venv/bin/python3 auto_run.py 城市1,城市2")
        print("例如: .venv/bin/python3 auto_run.py 福州,上海")
        sys.exit(1)

    cities = [c.strip() for c in sys.argv[1].split(",")]
    cities_map = api.load_city_map()

    print("=" * 50)
    print("等待手机 App 搜索…")
    print("请在手机 App 里搜索任意一个城市（看到酒店列表即可）")
    print("=" * 50)

    # 记录当前最新时间戳
    prev_ts = get_latest_ts()
    start = time.time()

    while True:
        time.sleep(2)
        latest = get_latest_ts()
        if latest > prev_ts:
            age = (datetime.now().timestamp() - latest / 1e6)
            print(f"\n✅ 检测到新的 App 请求!（{age:.0f}秒前）")
            break
        elapsed = time.time() - start
        if elapsed % 10 < 2:
            print(f"  等待中...（{elapsed:.0f}秒）", end="\r")
        if elapsed > 300:
            print("\n超时 5 分钟，请确认手机 Reqable 正在抓包、App 已搜索")
            sys.exit(1)

    # 刷新 token
    print("刷新 token...")
    api.extract_token_from_reqable(log=print)

    # 批量抓取
    all_res = []
    for name in cities:
        cid = cities_map.get(name) or cities_map.get(name.replace("市", ""))
        if cid is None:
            print(f"⚠️ 未知城市: {name}，跳过")
            continue
        print(f"\n=== {name} ===")
        res = api.scrape_city(cid, datetime.now().strftime("%Y-%m-%d"),
                              (datetime.now() + __import__("datetime").timedelta(days=1)).strftime("%Y-%m-%d"),
                              max_pages=5, log=print)
        for r in res:
            r["_city_label"] = name
        all_res += res
        print(f"  → {len(res)} 家")

    # 出文档
    h = {"ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
         "city": " / ".join(cities), "in": datetime.now().strftime("%Y-%m-%d"),
         "out": (datetime.now() + __import__("datetime").timedelta(days=1)).strftime("%Y-%m-%d"),
         "adults": 1, "n": len(all_res), "gamt": None}
    os.makedirs("output", exist_ok=True)
    safe = "".join(c for c in h["city"] if c not in '\\/:*?"<>| ')[:30] or "酒店"
    path = f"output/{safe}_{datetime.now():%Y-%m-%d_%H%M%S}.docx"
    g.write_docx(all_res, path, h)
    print(f"\n📄 文档: {path}（{len(all_res)} 家）")
    g.open_file(path)


if __name__ == "__main__":
    import gzip
    main()
