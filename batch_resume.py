import sys, os, time, json
from datetime import datetime
import hopegoo_api as api
import hopegoo_gcoins as g

OUT = "/Users/congwenjie/Desktop/未命名文件夹"
os.makedirs(OUT, exist_ok=True)

CITIES = """北京 上海 广州 深圳
成都 杭州 重庆 武汉 苏州 西安 南京 长沙 郑州 天津 合肥 青岛 东莞 宁波 佛山
厦门 大连 沈阳 昆明 南昌 哈尔滨 泉州 常州 南通 烟台 温州 长春 南宁 贵阳
石家庄 太原 珠海 嘉兴 金华 绍兴 潍坊 徐州 惠州 台州 呼和浩特 乌鲁木齐
扬州 中山 保定 兰州""".split()

cmap = api.load_city_map()
# 从中间文件恢复已抓结果
import glob
mid_files = sorted(glob.glob(os.path.join(OUT, "中间结果_*_*.docx")))
done_cities = set()
all_res = []
if mid_files:
    print(f"发现中间存档，已抓过 {len(mid_files)} 轮")
    # 简化：重新跑全部（数据去重）

start_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 0
total = len(CITIES)
failed = []

print(f"\n从第 {start_idx+1} 个城市开始 ({CITIES[start_idx] if start_idx < total else 'done'})")
print("=" * 50)

for i in range(start_idx, total):
    name = CITIES[i]
    cid = cmap.get(name)
    if not cid:
        print(f"[{i+1}/{total}] ⚠️ {name} 无ID")
        continue
    
    print(f"\n[{i+1}/{total}] {name}(city={cid})")
    
    ok = False
    for attempt in range(5):
        try:
            res = api.scrape_city(cid,
                datetime.now().strftime("%Y-%m-%d"),
                (datetime.now() + __import__("datetime").timedelta(days=1)).strftime("%Y-%m-%d"),
                max_pages=3, log=lambda x: print(f"  {x}"))
            for r in res: r["_city_label"] = name
            all_res += res
            print(f"  ✅ {len(res)} 家")
            ok = True
            break
        except Exception as e:
            err = str(e)
            print(f"  第{attempt+1}次: {err[:60]}")
            if "token" in err.lower() or "anti" in err.lower() or attempt >= 2:
                print("\n🔴 Token 过期! 请去手机 App 搜一下任意城市，然后回车继续...")
                input()
                api.extract_token_from_reqable(log=print)
            else:
                time.sleep(2)
    if not ok:
        failed.append(name)

    if (i+1) % 10 == 0:
        h = {"ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"city":f"批量{i+1}城",
             "in": datetime.now().strftime("%Y-%m-%d"),"out":(datetime.now()+__import__("datetime").timedelta(days=1)).strftime("%Y-%m-%d"),
             "adults":1,"n":len(all_res),"gamt":None}
        mid = os.path.join(OUT, f"中间结果_{i+1}城_{datetime.now():%H%M%S}.docx")
        try: g.write_docx(all_res, mid, h); print(f"  💾 {mid}")
        except: pass

# 最终
h = {"ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"city":f"全国{len(CITIES)}城",
     "in": datetime.now().strftime("%Y-%m-%d"),"out":(datetime.now()+__import__("datetime").timedelta(days=1)).strftime("%Y-%m-%d"),
     "adults":1,"n":len(all_res),"gamt":None}
final = os.path.join(OUT, f"全国GCoins酒店_{datetime.now():%Y%m%d_%H%M%S}.docx")
g.write_docx(all_res, final, h)
print(f"\n✅ {len(all_res)} 家 | {final}")
if failed: print(f"⚠️ 失败: {failed}")
g.open_file(final)
