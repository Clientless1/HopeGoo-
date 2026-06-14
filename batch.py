# -*- coding: utf-8 -*-
"""批量抓取多城市 G-Coins 酒店"""
import sys, os, time, json
from datetime import datetime
import hopegoo_api as api
import hopegoo_gcoins as g

OUT = "/Users/congwenjie/Desktop/未命名文件夹"
os.makedirs(OUT, exist_ok=True)

CITIES = """
北京 上海 广州 深圳
成都 杭州 重庆 武汉 苏州 西安 南京 长沙 郑州 天津 合肥 青岛 东莞 宁波 佛山
厦门 大连 沈阳 昆明 南昌 哈尔滨 泉州 常州 南通 烟台 温州 长春 南宁 贵阳
石家庄 太原 珠海 嘉兴 金华 绍兴 潍坊 徐州 惠州 台州 呼和浩特 乌鲁木齐
扬州 中山 保定 兰州
""".split()

cmap = api.load_city_map()
missing = [c for c in CITIES if not cmap.get(c)]
if missing:
    print(f"未知城市(需要App搜一次后刷新城市表): {missing}")
    cmap = api.refresh_city_map_from_reqable(log=print)

total = len(CITIES)
all_res = []
failed = []

for i, name in enumerate(CITIES, 1):
    cid = cmap.get(name)
    if not cid:
        print(f"[{i}/{total}] ⚠️ {name} 无ID，跳过")
        failed.append(name)
        continue

    print(f"\n{'='*40}")
    print(f"[{i}/{total}] {name}(city={cid})")

    # 失败重试（token过期自动刷新）
    for attempt in range(3):
        try:
            res = api.scrape_city(
                cid,
                datetime.now().strftime("%Y-%m-%d"),
                (datetime.now() + __import__("datetime").timedelta(days=1)).strftime("%Y-%m-%d"),
                max_pages=3, log=print
            )
            for r in res: r["_city_label"] = name
            all_res += res
            print(f"  → {len(res)} 家 (累计 {len(all_res)})")
            break
        except Exception as e:
            print(f"  第{attempt+1}次失败: {e}")
            if attempt < 2:
                print("  等2秒重试...")
                time.sleep(2)
            else:
                failed.append(name)
                print(f"  ❌ {name} 失败")

    # 每10个城市存一次中间结果，防止丢失
    if i % 10 == 0:
        h = {"ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             "city": f"批量{i}城", "in": datetime.now().strftime("%Y-%m-%d"),
             "out": (datetime.now() + __import__("datetime").timedelta(days=1)).strftime("%Y-%m-%d"),
             "adults": 1, "n": len(all_res), "gamt": None}
        mid = os.path.join(OUT, f"中间结果_{i}城_{datetime.now():%H%M%S}.docx")
        g.write_docx(all_res, mid, h)
        print(f"  💾 中间存档: {mid}")

# 最终输出
h = {"ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
     "city": f"全国{len(CITIES)-len(failed)}城",
     "in": datetime.now().strftime("%Y-%m-%d"),
     "out": (datetime.now() + __import__("datetime").timedelta(days=1)).strftime("%Y-%m-%d"),
     "adults": 1, "n": len(all_res), "gamt": None}
final = os.path.join(OUT, f"全国G-Coins酒店_{datetime.now():%Y%m%d_%H%M%S}.docx")
g.write_docx(all_res, final, h)
print(f"\n{'='*40}")
print(f"✅ 完成! {len(all_res)} 家酒店")
print(f"📄 文档: {final}")
if failed:
    print(f"⚠️ 失败城市: {failed}")
g.open_file(final)
