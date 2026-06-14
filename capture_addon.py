"""mitmproxy addon：自动捕获 hopegoo App token"""
import json,os,time

BASE=os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE=os.path.join(BASE,"reshg_req.json")
STATUS_FILE=os.path.join(BASE,"token_updated.txt")
LOG_FILE=os.path.join(BASE,"proxy_log.txt")

def plog(msg):
    t=time.strftime("%H:%M:%S")
    with open(LOG_FILE,"a") as f:f.write(f"[{t}] {msg}\n")

def request(flow):
    if "reshg" in flow.request.pretty_host or "hopegoo" in flow.request.pretty_host:
        plog(f"REQ {flow.request.pretty_host}{flow.request.path[:60]}")

def response(flow):
    host=flow.request.pretty_host
    if "hotel.reshg.com" not in host and "hopegoo" not in host:return
    path=flow.request.path
    plog(f"RES {flow.response.status_code} {host}{path[:80]}")

    if "hotel.reshg.com" not in host:return
    if "/tapi/v2/list" not in path:return

    hd={};cookies=[]
    for k,v in flow.request.headers.items():
        if k.lower()=="cookie":cookies.append(v)
        else:hd[k]=v
    has_stk=any(k.lower()=="sectoken" and len(v)>20 for k,v in hd.items())
    has_dun=any(k.lower()=="dun-token" for k,v in hd.items())

    # mitmproxy 会把多个Cookie合并成一个header，所以不能数header数量
    if has_stk and has_dun:
        json.dump({"path":path,"headers":hd,"cookies":cookies},open(TOKEN_FILE,"w"),ensure_ascii=False)
        open(STATUS_FILE,"w").write(str(time.time()))
        plog(f"✅ TOKEN捕获成功! sectoken={len(hd.get('sectoken',''))} cookies={len(cookies)}")
    else:
        plog(f"❌ 请求不完整 stk={has_stk} dun={has_dun} ck={len(cookies)}")
