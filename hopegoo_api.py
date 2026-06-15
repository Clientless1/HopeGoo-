# -*- coding: utf-8 -*-
"""
HopeGoo App 接口抓取核心（hotel.reshg.com/tapi/v2/list）
通过抓包得到的 token 模板(reshg_req.json) 直接调用 App 后端，
拿到带 G-Coins 抵扣的真实价格（网页版没有）。

字段映射：
  price          原价
  couponPrice    抵扣金额(G-Coins 等优惠券抵扣)
  discountPrice  付款金额(= price - couponPrice)
  hotelName / hotelAddress / cityName
"""
import os, re, json, gzip, time, urllib.parse
import requests as _requests

BASE = "https://hotel.reshg.com"
import sys as _sys
_APP_DIR = os.path.dirname(_sys.executable) if getattr(_sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
REQ_TEMPLATE = os.path.join(_APP_DIR, "reshg_req.json")
_SESSION = _requests.Session()
_SESSION.verify = False
import urllib3; urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 已知城市 ID（可继续补充；运行时会从 Reqable 抓包自动扩充并缓存到 cities.json）
CITY_IDS = {
    "上海": 321, "福州": 54, "北京": 53,
}
CITIES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cities.json")


def load_city_map():
    m = dict(CITY_IDS)
    try:
        with open(CITIES_FILE, encoding="utf-8") as f:
            m.update({k: int(v) for k, v in json.load(f).items()})
    except Exception:
        pass
    return m


def refresh_city_map_from_reqable(db_path=None, log=print):
    """从抓包响应里抽取 cityName→cityId，扩充 cities.json。"""
    import base64
    db = db_path or os.path.expanduser(
        "~/Library/Application Support/com.reqable.macosx/box/data.mdb")
    import glob
    found = load_city_map()

    def harvest(t):
        for nm, cid in re.findall(r'"cityName"\s*:\s*"([^"]{1,12})"[^{}]{0,400}?"cityId"\s*:\s*"?(\d+)', t):
            found[nm.replace("市", "")] = int(cid)
        for cid, nm in re.findall(r'"cityId"\s*:\s*"?(\d+)"?[^{}]{0,400}?"cityName"\s*:\s*"([^"]{1,12})"', t):
            found[nm.replace("市", "")] = int(cid)

    # 1) 抓包响应文件（明文 JSON）
    capdir = os.path.join(os.path.dirname(db), "..", "capture")
    for f in glob.glob(os.path.join(capdir, "*-res-extract-body.reqable")):
        try:
            b = open(f, encoding="utf-8", errors="replace").read()
            if '"cityName"' in b and '"cityId"' in b:
                harvest(b)
        except Exception:
            pass
    # 2) LMDB 里的元数据块
    try:
        raw = open(db, "rb").read()
        for blk in re.findall(rb"H4sI[A-Za-z0-9+/=]{80,}", raw):
            try:
                harvest(gzip.decompress(base64.b64decode(blk + b"==" * (-len(blk) % 4))).decode("utf-8", "replace"))
            except Exception:
                continue
    except Exception:
        pass
    json.dump(found, open(CITIES_FILE, "w"), ensure_ascii=False, indent=1)
    log(f"城市映射已更新，共 {len(found)} 个城市")
    return found


REQABLE_DB = os.path.expanduser(
    "~/Library/Application Support/com.reqable.macosx/box/data.mdb")


def extract_token_from_reqable(db_path=REQABLE_DB, log=print):
    """
    从 Reqable 的抓包库里提取最新一条 hotel.reshg.com /tapi/v2/list 请求
    的 headers+cookies(含 sectoken)，写入 reshg_req.json。
    用户在 App 里搜一次后调用即可刷新 token。
    """
    import base64
    raw = open(db_path, "rb").read()
    blocks = re.findall(rb"H4sI[A-Za-z0-9+/=]{80,}", raw)
    best = None
    for blk in blocks:
        try:
            d = json.loads(gzip.decompress(
                base64.b64decode(blk + b"==" * (-len(blk) % 4))).decode("utf-8", "replace"))
        except Exception:
            continue
        sess = d.get("session", {}) if isinstance(d, dict) else {}
        if sess.get("connection", {}).get("originHost", "") != "hotel.reshg.com":
            continue
        req = sess.get("request", {})
        rl = req.get("requestLine", {})
        path = rl.get("path", "") if isinstance(rl, dict) else ""
        hs = req.get("headers", [])
        if "/tapi/v2/list?" not in path:
            continue
        low = [str(h).lower() for h in hs]
        has_sectoken = any(l.startswith("sectoken:") and len(l.split(":", 1)[1].strip()) > 20 for l in low)
        has_dun = any(l.startswith("dun-token:") for l in low)
        cookie_count = sum(1 for l in low if l.startswith("cookie:"))
        if not (has_sectoken and has_dun and cookie_count > 0):
            continue
        ts = sess.get("timestamp", 0)
        # 优先 cookie 最全的（完整鉴权），其次最新的
        score = (cookie_count, ts)
        if best is None or score > best[0]:
            best = (score, path, hs)
    if not best:
        raise RuntimeError("没在 Reqable 抓包里找到完整鉴权的列表请求；请在 App 里搜一次酒店后重试")
    (_, _), path, hs = best
    hd, cookies = {}, []
    for h in hs:
        s = str(h); i = s.find(":")
        if i < 0:
            continue
        nm, v = s[:i].strip(), s[i + 1:].strip()
        if nm.lower() == "cookie":
            cookies.append(v)
        else:
            hd[nm] = v
    json.dump({"path": path, "headers": hd, "cookies": cookies},
              open(REQ_TEMPLATE, "w"), ensure_ascii=False, indent=1)
    sectok_key = [k for k in hd if k.lower() == "sectoken"]
    log(f"已刷新 token（sectoken 长度 {len(hd.get(sectok_key[0],'')) if sectok_key else 0}，cookie {len(cookies)} 条）")
    return True


def load_template():
    """读取抓包得到的请求模板（headers + cookies）。"""
    with open(REQ_TEMPLATE, encoding="utf-8") as f:
        return json.load(f)


def _build_headers(tpl):
    """原样重建请求头，去掉 br 编码（requests 不支持自动解压 brotli）。"""
    hd = {}
    for k, v in tpl.get("headers", {}).items():
        if k.lower() == "accept-encoding":
            hd[k] = "gzip, deflate"
        else:
            hd[k] = v
    hd["Cookie"] = "; ".join(tpl.get("cookies", []))
    return hd


def query_list(city_id, in_date, out_date, page_index, tpl, adults=1, log=print):
    """调用 /tapi/v2/list，返回该页 hotelList。低价优先 + G-Coins 筛选。"""
    params = {
        "city": str(city_id),
        "filterList": "8888_4,391010_1",   # 8888_4=低价优先, 391010_1=G-Coins
        "inDate": in_date, "outDate": out_date,
        "pageIndex": str(page_index), "pageSize": "20",
        "adultsNumber": str(adults), "currency": "CNY",
        "ref": "hopegooh5", "needTitleBar": "0", "scriptVersion": "0.2.15",
    }
    url = BASE + "/tapi/v2/list?" + urllib.parse.urlencode(params)
    hd = _build_headers(tpl)
    hd["Referer"] = (BASE + "/m/zh-hk/hotel/hotellist?city=%s&inDate=%s&outDate=%s&currency=CNY"
                     % (city_id, in_date, out_date))
    resp = _SESSION.get(url, headers=hd, timeout=30)
    d = resp.json()
    data = d.get("data")
    if not data and str(d.get("errorCode", "0")) not in ("0", "None"):
        raise TokenExpired(f"接口返回 errorCode={d.get('errorCode')}（anti-bot 令牌已失效，需在 App 里刷新）")
    return (data or {}).get("hotelList") or []


class TokenExpired(Exception):
    pass


def ensure_fresh_template(auto=True, log=print):
    """直接读取 reshg_req.json（由内置代理自动更新）。"""
    if not os.path.exists(REQ_TEMPLATE):
        raise RuntimeError("缺少 token，请先启动代理并在 App 搜索")
    return load_template()


def get_gcoin_amount(hotel):
    """从 productLabelList 取 G-Coins(345) 的 amount，没有则返回 None。"""
    for lb in (hotel.get("productLabelList") or []):
        if isinstance(lb, dict) and lb.get("productLabelId") == 345:
            return lb.get("amount")
    return None


def scrape_city(city_id, in_date, out_date, adults=1, max_pages=50,
                gcoin_amount=None,
                log=print, auto_token=True):
    """
    分页抓取一个城市所有带 G-Coins 标签的酒店（接口已用 391010_1 筛选）。
    返回 [{name,address,price,discountPrice,couponPrice,city}]
    """
    tpl = ensure_fresh_template(auto_token, log)
    results, seen = [], set()
    retried = False
    for pi in range(max_pages):
        try:
            hl = query_list(city_id, in_date, out_date, pi, tpl, adults, log)
        except TokenExpired:
            if not retried:
                retried = True
                log("  token过期，等待App搜索…")
                for _ in range(30):
                    time.sleep(2)
                    try: tpl=load_template();hl=query_list(city_id,in_date,out_date,pi,tpl,adults,log);break
                    except:pass
                else:log("  超时");break
            else:break
        except Exception as e:
            log(f"  第{pi+1}页请求失败: {e}")
            break
        if not hl:
            break
        new = 0
        for h in hl:
            hid = h.get("hotelId") or h.get("hotelName")
            if hid in seen:
                continue
            seen.add(hid); new += 1
            g_amt = get_gcoin_amount(h)
            if gcoin_amount is not None:
                if g_amt != gcoin_amount:
                    continue  # 只要指定金额的（如20=可抵CNY20）
            results.append({
                "name": (h.get("hotelName") or "").strip(),
                "address": (h.get("hotelAddress") or "").strip() or "（未返回地址）",
                "price": h.get("price"),
                "discountPrice": h.get("discountPrice"),
                "couponPrice": h.get("couponPrice"),
                "gcoinAmount": g_amt,
                "city": h.get("cityName") or "",
            })
        log(f"  第{pi+1}页：{new}家新，累计 {len(results)} 家")
        if new == 0:
            break
    return results


if __name__ == "__main__":
    import sys
    cid = sys.argv[1] if len(sys.argv) > 1 else "321"
    res = scrape_city(cid, "2026-06-14", "2026-06-15", threshold=25)
    print(f"\n命中 {len(res)} 家：")
    for r in res:
        print(f"  付款CNY{r['pay']} 抵扣CNY{r['coupon']} 原价CNY{r['price']} | {r['name']} | {r['address'][:30]}")
