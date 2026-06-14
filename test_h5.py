# -*- coding: utf-8 -*-
"""测试：Playwright 模拟手机浏览器，用桌面登录 cookie 打开 hotel.reshg.com H5 列表页，
   拦截 /tapi/v2/list XHR 响应，看能否绕过 token 问题。"""
import time, json
from playwright.sync_api import sync_playwright
from hopegoo_hotel_filter import do_login

ACC="xekdyqmc@hotmail.com"; PWD="ljta4227"
UA=("Mozilla/5.0 (Linux; Android 16; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/148.0.7778.96 Mobile Safari/537.36")

def L(m): print(m, flush=True)
captured=[]

with sync_playwright() as p:
    # 1. 桌面登录拿 cookie
    br=p.chromium.launch(channel="chrome", headless=False)
    d=br.new_context(viewport={"width":1440,"height":900}, locale="zh-Hant")
    dp=d.new_page(); dp.set_default_timeout(15000)
    ok=do_login(dp, ACC, PWD, L)
    L(f"登录: {ok}")
    cookies=d.cookies()
    d.close()

    if not ok:
        print("登录失败"); exit()

    # 2. 手机 H5 context，注入登录 cookie
    m=br.new_context(viewport={"width":390,"height":844}, user_agent=UA,
                     is_mobile=True, has_touch=True, device_scale_factor=2, locale="zh-Hant")
    page=m.new_page(); page.set_default_timeout(20000)

    # 先访问 hopegoo 设 cookie，再跳转到 hotel.reshg.com
    page.goto("https://www.hopegoo.com/zh-HK", wait_until="commit", timeout=90000)
    m.add_cookies(cookies)
    L("已注入 cookie")
    time.sleep(2)

    # 3. 拦截所有响应，看页面实际发什么请求
    all_resp=[]
    page.on("response", lambda r: all_resp.append(r.url))

    # 4. 打开 H5 列表
    url="https://hotel.reshg.com/m/zh-hk/hotel/hotellist?city=321&filterList=8888_4,391010_1&inDate=2026-06-14&outDate=2026-06-15&adultsNumber=1&currency=CNY"
    L(f"打开: {url[:90]}")
    page.goto(url, wait_until="commit", timeout=90000)

    # 5. 结果
    L(f"所有响应({len(all_resp)}条):")
    reshg_resp=[u for u in all_resp if "reshg" in u or "tapi" in u or "hopegoo" in u]
    for u in reshg_resp[:30]: L(f"  {u[:120]}")
    # 看页面内容
    body=page.evaluate("()=>document.body.innerText")
    L(f"页面文本前500字: {body[:500]}")
    br.close()
