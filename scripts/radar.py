#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
必赢 · 标讯雷达  查询引擎  radar.py
====================================
只用 Python 标准库，Windows / macOS 通用（Python ≥ 3.8）。

确定性的活全在代码里，模型只负责听懂人话、传语义参数、把 JSON 结果讲给用户：
  · 日期换算   —— 一律按北京时间计算，不由模型推断「今天几号」
  · 参数锁死   —— 精准匹配、每页 ≤50
  · 地区筛选   —— 人话地名 → 6 位地区码
  · 清洗去重   —— 删 HTML 标签、按 标题+甲方+金额 去重、剔除中标/合同等结果类
  · 截止日期   —— 列表阶段一律「未注明」，绝不拿发布时间冒充

用法：
    python3 radar.py check                                   # 环境自检（不调接口）
    python3 radar.py setkey --key <key>                      # 写入 key（不调接口）
    python3 radar.py search --keyword "数字人|虚拟人" --area 北京 --last 1m --probe
    python3 radar.py search --keyword "数字人" --last 1m --limit 20
    python3 radar.py detail --id 304974237                    # 发布时间脚本按 id 自己补
    python3 radar.py areas 杭州                               # 查地名（不调接口）
    python3 radar.py profile set --established 2019-06 --capital-wan 500 --headcount 30 --reg-area 杭州
    python3 radar.py profile cert --name "广播电视节目制作经营许可证" --has no
    python3 radar.py screen --ids "304974237,304974238"         # 只传 id，发布时间脚本自己补
（Windows 上把 python3 换成 python 或 py）

服务器每分钟最多接受 60 次请求。--probe 已并入正式搜索（返回里本来就有 total），不要为试参数反复调用。
本机缓存在 ~/.bidwin/cache：正文永久、结构化 2 天、同样的检索条件 30 分钟内复用，都不重复请求。
"""

import argparse
import hashlib
import html
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

VERSION = "1.2.0"
BASE_URL = "https://gate.gov-bid.com/outer-gateway/bid"   # 写死，不接受环境变量覆盖
TIMEOUT = 45

# 🔴 北京时间硬编码，不依赖服务器系统时区。
# 云端实例实测跑在 UTC：北京 00:00~08:00 这段 UTC 还停在前一天，
# 直接用 datetime.now() 会把「今天」算成昨天，日期窗整体偏一天。
CN_TZ = timezone(timedelta(hours=8))


def now_cn():
    """当前北京时间（naive datetime，便于直接格式化）"""
    return datetime.now(CN_TZ).replace(tzinfo=None)

# ── 锁死参数（红线，不开放 CLI 覆盖）────────────────────────────────────────
SEARCH_TYPE = 2      # 2=精准。1=智能模糊会拆字，实测 1205 条垃圾
SEARCH_MODE = 1      # 1=标题+正文。只搜标题会漏「标题不含AI、正文含」的真标
PAGE_MAX = 50        # 每页硬顶 50（传 100 也只返 50，实测三次一致）

# ── 招中标信息分类（doc 枚举值-码表 2）────────────────────────────────────
CLASS_NAMES = {
    "1": "公开招标", "2": "成交结果", "3": "合同公告", "4": "意向公开",
    "5": "答疑变更", "6": "候选人公示", "7": "开标公示", "8": "重新招标",
    "11": "流标废标", "18": "结果变更", "26": "拍租公告", "28": "竞争性谈判",
    "29": "竞争性磋商", "30": "单一来源采购", "31": "其它",
}
# 阶段含义标注
CLASS_NOTES = {
    "1": "可投标",
    "4": "⚠️早期线索·尚未开始采购，现在投不了，只能提前接触",
    "8": "可投标（重新招标）",
    "28": "可投标（竞争性谈判）",
    "29": "可投标（竞争性磋商）",
    "30": "⚠️单一来源·通常已内定供应商",
    "5": "⚠️答疑变更·须回看原招标公告",
    "26": "拍租公告",
}
# 结果类：已经结束的，日常扫标默认剔除（实测「混入中标公告」）
RESULT_CLASS_IDS = {"2", "3", "6", "7", "11", "18"}
# 日常可投默认分类
DEFAULT_CLASS_IDS = "1,4,8,28,29"
# 采购类型：默认不限（客户行业各异，不预设服务类）
PURCHASE_TYPES = {"服务": "1", "工程": "2", "货物": "3", "其它": "0", "全部": "-100"}
DEFAULT_PURCHASE = "-100"
EXCLUDE_MAXLEN = 150   # 接口硬限制

# ── 地区码表（2026-08-05 快照，1=省 2=市）──────────────────────────────
# 地区一律走 areaCode 6 位码；把地名拼进关键词实测失效。区县级暂不支持。
_PROV_RAW = (
    "北京市:110000|安徽省:340000|福建省:350000|甘肃省:620000|广东省:440000|广西壮族自治区:450000|贵州省:520000|海南省:4"
    "60000|河北省:130000|河南省:410000|黑龙江省:230000|湖北省:420000|湖南省:430000|吉林省:220000|江苏省:320000|江西"
    "省:360000|辽宁省:210000|内蒙古自治区:150000|宁夏回族自治区:640000|青海省:630000|山东省:370000|山西省:140000|陕西省:"
    "610000|上海市:310000|四川省:510000|天津市:120000|西藏自治区:540000|新疆维吾尔自治区:650000|云南省:530000|浙江省:33"
    "0000|重庆市:500000|香港:810000"
)
_CITY_RAW = (
    "安庆市:340800|蚌埠市:340300|池州市:341700|滁州市:341100|阜阳市:341200|淮北市:340600|淮南市:340400|黄山市:34100"
    "0|六安市:341500|马鞍山市:340500|宿州市:341300|铜陵市:340700|芜湖市:340200|宣城市:341800|亳州市:341600|福州市:35"
    "0100|龙岩市:350800|南平市:350700|宁德市:350900|莆田市:350300|泉州市:350500|三明市:350400|厦门市:350200|漳州市:"
    "350600|兰州市:620100|白银市:620400|定西市:621100|甘南藏族自治州:623000|嘉峪关市:620200|金昌市:620300|酒泉市:6209"
    "00|临夏回族自治州:622900|陇南市:621200|平凉市:620800|庆阳市:621000|天水市:620500|武威市:620600|张掖市:620700|广州"
    "市:440100|深圳市:440300|潮州市:445100|东莞市:441900|佛山市:440600|河源市:441600|惠州市:441300|江门市:440700|"
    "揭阳市:445200|茂名市:440900|梅州市:441400|清远市:441800|汕头市:440500|汕尾市:441500|韶关市:440200|阳江市:44170"
    "0|云浮市:445300|湛江市:440800|肇庆市:441200|中山市:442000|珠海市:440400|南宁市:450100|桂林市:450300|百色市:451"
    "000|北海市:450500|崇左市:451400|防城港市:450600|贵港市:450800|河池市:451200|贺州市:451100|来宾市:451300|柳州市:"
    "450200|钦州市:450700|梧州市:450400|玉林市:450900|贵阳市:520100|安顺市:520400|毕节市:520500|六盘水市:520200|黔"
    "东南苗族侗族自治州:522600|黔南布依族苗族自治州:522700|黔西南布依族苗族自治州:522300|铜仁市:520600|遵义市:520300|海口市:460100"
    "|三亚市:460200|儋州市:460400|石家庄市:130100|保定市:130600|沧州市:130900|承德市:130800|邯郸市:130400|衡水市:131"
    "100|廊坊市:131000|秦皇岛市:130300|唐山市:130200|邢台市:130500|张家口市:130700|郑州市:410100|洛阳市:410300|开封市"
    ":410200|安阳市:410500|鹤壁市:410600|焦作市:410800|南阳市:411300|平顶山市:410400|三门峡市:411200|商丘市:411400"
    "|新乡市:410700|信阳市:411500|许昌市:411000|周口市:411600|驻马店市:411700|漯河市:411100|濮阳市:410900|哈尔滨市:23"
    "0100|大庆市:230600|大兴安岭地区:232700|鹤岗市:230400|黑河市:231100|鸡西市:230300|佳木斯市:230800|牡丹江市:231000"
    "|七台河市:230900|齐齐哈尔市:230200|双鸭山市:230500|绥化市:231200|伊春市:230700|武汉市:420100|鄂州市:420700|黄冈市:"
    "421100|黄石市:420200|荆门市:420800|荆州市:421000|十堰市:420300|随州市:421300|咸宁市:421200|襄阳市:420600|孝感"
    "市:420900|宜昌市:420500|恩施土家族苗族自治州:422800|长沙市:430100|张家界市:430800|常德市:430700|郴州市:431000|衡阳市"
    ":430400|怀化市:431200|娄底市:431300|邵阳市:430500|湘潭市:430300|湘西土家族苗族自治州:433100|益阳市:430900|永州市:4"
    "31100|岳阳市:430600|株洲市:430200|长春市:220100|吉林市:220200|白城市:220800|白山市:220600|辽源市:220400|四平市"
    ":220300|松原市:220700|通化市:220500|延边朝鲜族自治州:222400|南京市:320100|苏州市:320500|无锡市:320200|常州市:320"
    "400|淮安市:320800|连云港市:320700|南通市:320600|宿迁市:321300|泰州市:321200|徐州市:320300|盐城市:320900|扬州市:"
    "321000|镇江市:321100|南昌市:360100|抚州市:361000|赣州市:360700|吉安市:360800|景德镇市:360200|九江市:360400|萍"
    "乡市:360300|上饶市:361100|新余市:360500|宜春市:360900|鹰潭市:360600|沈阳市:210100|大连市:210200|鞍山市:210300"
    "|本溪市:210500|朝阳市:211300|丹东市:210600|抚顺市:210400|阜新市:210900|葫芦岛市:211400|锦州市:210700|辽阳市:211"
    "000|盘锦市:211100|铁岭市:211200|营口市:210800|呼和浩特市:150100|阿拉善盟:152900|巴彦淖尔市:150800|包头市:150200|"
    "赤峰市:150400|鄂尔多斯市:150600|呼伦贝尔市:150700|通辽市:150500|乌海市:150300|乌兰察布市:150900|锡林郭勒盟:152500|兴"
    "安盟:152200|银川市:640100|固原市:640400|石嘴山市:640200|吴忠市:640300|中卫市:640500|西宁市:630100|果洛藏族自治州:6"
    "32600|海北藏族自治州:632200|海东市:630200|海南藏族自治州:632500|海西蒙古族藏族自治州:632800|黄南藏族自治州:632300|玉树藏族自治"
    "州:632700|济南市:370100|青岛市:370200|滨州市:371600|德州市:371400|东营市:370500|菏泽市:371700|济宁市:370800|"
    "聊城市:371500|临沂市:371300|日照市:371100|泰安市:370900|威海市:371000|潍坊市:370700|烟台市:370600|枣庄市:37040"
    "0|淄博市:370300|太原市:140100|长治市:140400|大同市:140200|晋城市:140500|晋中市:140700|临汾市:141000|吕梁市:141"
    "100|朔州市:140600|忻州市:140900|阳泉市:140300|运城市:140800|西安市:610100|安康市:610900|宝鸡市:610300|汉中市:6"
    "10700|商洛市:611000|铜川市:610200|渭南市:610500|咸阳市:610400|延安市:610600|榆林市:610800|成都市:510100|绵阳市"
    ":510700|阿坝藏族羌族自治州:513200|巴中市:511900|达州市:511700|德阳市:510600|甘孜藏族自治州:513300|广安市:511600|广元"
    "市:510800|乐山市:511100|凉山彝族自治州:513400|眉山市:511400|南充市:511300|内江市:511000|攀枝花市:510400|遂宁市:51"
    "0900|雅安市:511800|宜宾市:511500|资阳市:512000|自贡市:510300|泸州市:510500|拉萨市:540100|阿里地区:542500|昌都市"
    ":540300|林芝市:540400|那曲市:540600|日喀则市:540200|山南市:540500|乌鲁木齐市:650100|阿克苏地区:652900|巴音郭楞蒙古自"
    "治州:652800|博尔塔拉蒙古自治州:652700|昌吉回族自治州:652300|哈密市:650500|和田地区:653200|喀什地区:653100|克拉玛依市:650"
    "200|克孜勒苏柯尔克孜自治州:653000|吐鲁番市:650400|伊犁哈萨克自治州:654000|昆明市:530100|怒江傈僳族自治州:533300|普洱市:5308"
    "00|丽江市:530700|保山市:530500|楚雄彝族自治州:532300|大理白族自治州:532900|德宏傣族景颇族自治州:533100|迪庆藏族自治州:53340"
    "0|红河哈尼族彝族自治州:532500|临沧市:530900|曲靖市:530300|文山壮族苗族自治州:532600|西双版纳傣族自治州:532800|玉溪市:530400"
    "|昭通市:530600|杭州市:330100|湖州市:330500|嘉兴市:330400|金华市:330700|丽水市:331100|宁波市:330200|绍兴市:3306"
    "00|台州市:331000|温州市:330300|舟山市:330900|衢州市:330800|阿勒泰地区:654300|合肥市:340100|三沙市:460300|县:50"
    "0200|塔城地区:654200|雄安新区:131200"
)


def _parse_area_raw(raw):
    out = {}
    for item in raw.split("|"):
        name, _, code = item.partition(":")
        out[name] = code
    return out


PROVINCES = _parse_area_raw(_PROV_RAW)
CITIES = _parse_area_raw(_CITY_RAW)

_AREA_SUFFIX = (
    "维吾尔自治区", "壮族自治区", "回族自治区", "特别行政区", "自治区",
    "自治州", "地区", "省", "市",
)


def _norm_area(name):
    """「北京市」「北京」「内蒙古自治区」「内蒙古」→ 同一个 key"""
    name = (name or "").strip()
    for suf in _AREA_SUFFIX:
        if name.endswith(suf) and len(name) > len(suf):
            return name[: -len(suf)]
    return name


_PROV_INDEX = {_norm_area(k): v for k, v in PROVINCES.items()}
_CITY_INDEX = {_norm_area(k): v for k, v in CITIES.items()}
_PROV_BY_CODE = {v: k for k, v in PROVINCES.items()}
_CITY_BY_CODE = {v: k for k, v in CITIES.items()}


def resolve_area(name):
    """人话地名 → ('province'|'city', 6位码)；解析不了返回 (None, None)"""
    key = _norm_area(name)
    if key in _PROV_INDEX:
        return "province", _PROV_INDEX[key]
    if key in _CITY_INDEX:
        return "city", _CITY_INDEX[key]
    return None, None


# ── 文本工具 ────────────────────────────────────────────────────────────
_TAG_RE = re.compile(r"<[^>]+>")


def strip_html(s):
    """删除高亮标签。⚠️必须删除、不得替换为空格，否则「AI短视频」→「AI 短 视频」"""
    if not s:
        return ""
    return html.unescape(_TAG_RE.sub("", s)).strip()




# ── 日期（🔴 全部由代码算，模型不得插手）──────────────────────────────────
_LAST_RE = re.compile(r"^(\d+)([dwmy])$", re.IGNORECASE)
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def resolve_window(args):
    """返回 (start_str, end_str, 说明)。今天永远取自北京时间 now_cn()。

    模型自己算日期实测会错一整年，所以这里不接受任何「今天是几号」的外部输入。
    """
    now = now_cn()
    today = now.date()

    if args.start or args.end:
        if not (args.start and args.end):
            die("--start 与 --end 必须成对给出")
        for v in (args.start, args.end):
            if not _DATE_RE.match(v):
                die("日期格式必须是 YYYY-MM-DD，收到：%s" % v)
        start_d = datetime.strptime(args.start, "%Y-%m-%d").date()
        end_d = datetime.strptime(args.end, "%Y-%m-%d").date()
        if start_d > end_d:
            die("--start 晚于 --end")
        if end_d > today:
            end_d = today
        how = "显式区间"
    else:
        days = args.days
        if args.last:
            m = _LAST_RE.match(args.last)
            if not m:
                die("--last 格式应为 7d / 2w / 1m / 1y，收到：%s" % args.last)
            n, unit = int(m.group(1)), m.group(2).lower()
            days = {"d": n, "w": n * 7, "m": n * 30, "y": n * 365}[unit]
        if days is None:
            days = 3  # 日常扫标默认近 3 天（只查当天会漏前两天发布仍可投的标）
        if days < 1:
            die("--days 至少为 1")
        end_d = today
        start_d = today - timedelta(days=days - 1)
        how = "近 %d 天（含今天）" % days

    return (
        start_d.strftime("%Y-%m-%d 00:00:00"),
        end_d.strftime("%Y-%m-%d 23:59:59"),
        "%s；脚本运行时刻 %s（北京时间，不随服务器时区变）"
        % (how, now.strftime("%Y-%m-%d %H:%M:%S")),
    )


# ── HTTP ────────────────────────────────────────────────────────────────
class ApiError(RuntimeError):
    pass


_CALLS = {"n": 0, "cached": 0}


def die(msg, code=2):
    print(json.dumps({"ok": False, "error": msg}, ensure_ascii=False, indent=2))
    sys.exit(code)


KEY_FILE = "~/.bidwin/key"


def key_path():
    return os.path.expanduser(KEY_FILE)


def mask(k):
    return ("*" * 6 + k[-4:]) if len(k) > 4 else "****"


def get_key():
    """key 来源：环境变量 BIDWIN_KEY → ~/.bidwin/key。永远不写进脚本或 skill 文件。"""
    key = os.environ.get("BIDWIN_KEY", "").strip()
    if key:
        return key
    try:
        with open(key_path(), encoding="utf-8-sig") as f:
            lines = [x.strip() for x in f.read().splitlines() if x.strip()]
    except OSError:
        lines = []
    if lines:
        return lines[0]
    die("还没有配置 key。请让用户把 key 发给你，然后执行：radar.py setkey --key <key>", code=4)


def post(path, payload, key):
    url = "%s/%s?key=%s" % (BASE_URL, path, key)
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, method="POST",
        headers={"Content-Type": "application/json", "X-Api-Key": key},
    )
    last = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
                raw = resp.read().decode("utf-8", "replace")
            break
        except urllib.error.HTTPError as e:  # 4xx/5xx 不重试内容类错误
            raw = e.read().decode("utf-8", "replace")
            if e.code in (429, 502, 503, 504) and attempt < 2:
                last = "HTTP %s" % e.code
                time.sleep(2 * (attempt + 1))  # 限速 60 次/分钟
                continue
            raise ApiError("HTTP %s：%s" % (e.code, raw[:300]))
        except urllib.error.URLError as e:
            last = str(e.reason)
            if attempt < 2:
                time.sleep(2 * (attempt + 1))
                continue
            raise ApiError("网络不可达：%s" % last)
    else:
        raise ApiError("重试 3 次仍失败：%s" % last)

    try:
        data = json.loads(raw)
    except ValueError:
        raise ApiError("返回非 JSON：%s" % raw[:300])
    if data.get("code") == 200:
        _CALLS["n"] += 1  # code=200 即计一次调用，与返回内容无关
    else:
        raise ApiError("接口返回 code=%s subCode=%s msg=%s"
                       % (data.get("code"), data.get("subCode"), data.get("msg") or data.get("subMsg")))
    return data


# ── 本地缓存（只存在本机 ~/.bidwin/cache，不上传）──────────────────────────
# 实测一位试用客户一个月里同一条标被反复拉了 613 次。正文不会变，拉过一次就存下来；
# 结构化数据（报名截止等）可能被更正公告改掉，只存几天。缓存读写失败不影响查询。
CACHE_DIR = "~/.bidwin/cache"
STRUCT_TTL_DAYS = 2      # 结构化详情缓存几天
SEARCH_TTL_MIN = 30      # 同样的检索条件多久内直接复用上次结果（「直接看」「更多」不再请求）


def _cache_file(kind, name):
    return os.path.join(os.path.expanduser(CACHE_DIR), kind, "%s.json" % name)


def cache_get(kind, name, max_age=None):
    p = _cache_file(kind, name)
    try:
        if max_age is not None and time.time() - os.path.getmtime(p) > max_age:
            return None
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def cache_put(kind, name, data):
    p = _cache_file(kind, name)
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        os.replace(tmp, p)
    except OSError:
        pass


def _sig(obj):
    return hashlib.sha1(json.dumps(obj, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:20]


# ── 搜索用量（只存在本机 ~/.bidwin/usage.json）─────────────────────────────
# 实测 Agent 会把一个需求拆成 40 多次单词搜索、每天把近一个月重搜一遍。
# 本地计数是软限制（换电脑 / 删文件就没了），只防 Agent 自己越界；硬限额由网关管。
USAGE_FILE = "~/.bidwin/usage.json"
SEARCH_GUARD_N = 8       # 每 SEARCH_GUARD_MIN 分钟内最多真实搜索几次（复用缓存的不算）
SEARCH_GUARD_MIN = 10
USAGE_KEEP_DAYS = 14
REPEAT_MIN_DAYS = 7      # 检索范围 ≥ 这么多天才算「全量搜」


def load_usage():
    try:
        with open(os.path.expanduser(USAGE_FILE), encoding="utf-8") as f:
            d = json.load(f)
    except (OSError, ValueError):
        d = {}
    d.setdefault("searches", [])
    return d


def save_usage(d):
    p = os.path.expanduser(USAGE_FILE)
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
        os.replace(tmp, p)
    except OSError:
        pass


def _cond_sig(payload):
    """检索条件指纹：去掉日期和翻页，只看「查什么」"""
    return _sig({k: v for k, v in payload.items()
                 if k not in ("startDate", "endDate", "pageId", "pageNumber")})


def _window_days(payload):
    s = datetime.strptime(payload["startDate"][:10], "%Y-%m-%d").date()
    e = datetime.strptime(payload["endDate"][:10], "%Y-%m-%d").date()
    return (e - s).days + 1


def guard_search():
    recent = [x for x in load_usage()["searches"] if x["at"] > time.time() - SEARCH_GUARD_MIN * 60]
    if len(recent) >= SEARCH_GUARD_N:
        die("%d 分钟内已经搜了 %d 次。请把同义词用 | 合并成一次搜索，或者先和用户确认要查的范围；"
            "不要换个写法继续拆开搜。用户确认要继续查的，%d 分钟后再试"
            % (SEARCH_GUARD_MIN, len(recent), SEARCH_GUARD_MIN), code=6)


def record_search(payload):
    d = load_usage()
    keep = time.time() - USAGE_KEEP_DAYS * 86400
    d["searches"] = [x for x in d["searches"] if x["at"] > keep]
    d["searches"].append({"at": int(time.time()), "day": now_cn().strftime("%Y-%m-%d"),
                          "cond": _cond_sig(payload), "days": _window_days(payload)})
    save_usage(d)


def repeat_hint(payload):
    """同一组条件前几天已经全量搜过 → 提示改用定时推送（近 3 天 + 推过的跳过）"""
    if _window_days(payload) < REPEAT_MIN_DAYS:
        return None
    today = now_cn().strftime("%Y-%m-%d")
    since = (now_cn() - timedelta(days=REPEAT_MIN_DAYS)).strftime("%Y-%m-%d")
    cond = _cond_sig(payload)
    days = sorted({x["day"] for x in load_usage()["searches"]
                   if x["cond"] == cond and x["days"] >= REPEAT_MIN_DAYS and since <= x["day"] < today})
    if not days:
        return None
    return ("这组条件 %s 已经按近 %d 天以上搜过。用户如果是每天来看新标，"
            "建议改用每日推送（只查近 3 天、推过的自动跳过），比每天重搜一遍省得多"
            % ("、".join(days), REPEAT_MIN_DAYS))


def coverage_of(raw_rows, start, end, has_next):
    """这一页覆盖了检索范围里的几天。接口只按发布时间倒序，第一页就是最新的那一批：
    total 很大时第一页只能看到最近几天甚至几小时，比 total 更能说明用户实际看得到多少。"""
    s = datetime.strptime(start[:10], "%Y-%m-%d").date()
    e = datetime.strptime(end[:10], "%Y-%m-%d").date()
    window = (e - s).days + 1
    times = sorted((r.get("publishTime") or "")[:10] for r in raw_rows if r.get("publishTime"))
    if not has_next or not times:
        return {"window_days": window, "covered_days": window, "complete": True,
                "earliest": times[0] if times else None,
                "text": "检索范围内的结果已全部在这一页"}
    earliest = datetime.strptime(times[0], "%Y-%m-%d").date()
    covered = max(1, min(window, (e - earliest).days + 1))
    if covered == 1:
        text = "这一页 %d 条全是 %s 当天发布的，%d 天的检索范围只看到了最后 1 天" % (
            len(raw_rows), e.strftime("%m-%d"), window)
    else:
        text = "这一页 %d 条最早到 %s，%d 天的检索范围只覆盖了最近 %d 天" % (
            len(raw_rows), earliest.strftime("%m-%d"), window, covered)
    return {"window_days": window, "covered_days": covered, "complete": False,
            "earliest": times[0], "text": text}


# ── 搜过的标：id → 发布时间（只存在本机 ~/.bidwin/cache/ids.json）────────────
# 看详情、screen 只传 id 就行，发布时间由这里补；id 不在里面说明多半抄错了，直接报错不发请求。
# 实测 Agent 手抄 19 组「id|发布时间」时抄错一个，恰好是另一条真实的标，脚本照常筛了，白花请求。
IDS_KEEP_DAYS = 7        # 用户隔天再说「按资质筛一下昨天那批」也要认得


def remember_ids(raw_rows):
    """记下搜索 / 推送拉到的每条标。用接口原始的 publishTime（去重时会把重复条的时间合过来，不用那个）"""
    if not raw_rows:
        return
    d = cache_get("index", "ids") or {}
    now = int(time.time())
    keep = now - IDS_KEEP_DAYS * 86400
    d = {k: v for k, v in d.items() if v.get("at", 0) > keep}
    for r in raw_rows:
        if r.get("id") and r.get("publishTime"):
            d[str(r["id"])] = {"pt": r["publishTime"], "at": now}
    cache_put("index", "ids", d)


def resolve_ids(items):
    """[(id, 发布时间或空)] → [(id, 发布时间)]。发布时间空的从记录里补；
    不在最近 IDS_KEEP_DAYS 天搜索 / 推送结果里的 id 一律报错，一次请求都不发。"""
    known = cache_get("index", "ids") or {}
    keep = time.time() - IDS_KEEP_DAYS * 86400
    out, unknown = [], []
    for bid, pt in items:
        rec = known.get(str(bid))
        if not rec or rec.get("at", 0) <= keep:
            unknown.append(str(bid))
            continue
        out.append((bid, pt or rec["pt"]))
    if unknown:
        die("这些 id 不在最近 %d 天的搜索 / 推送结果里：%s。多半是抄错了，请对照搜索结果核对后重试；"
            "如果是更早查到的标，先用原来的条件再搜一次" % (IDS_KEEP_DAYS, "、".join(unknown)), code=7)
    return out


def get_struct(bid, pt, key):
    """结构化详情：报名截止 / 预算 / 联系人 / 开标。缓存 STRUCT_TTL_DAYS 天"""
    d = cache_get("struct", bid, STRUCT_TTL_DAYS * 86400)
    if d is not None:
        _CALLS["cached"] += 1
        return d
    d = post("getZTBStructreDetail", {"id": bid, "publishTime": pt}, key).get("data") or {}
    if d:
        cache_put("struct", bid, d)
    return d


def get_content(bid, pt, key):
    """正文详情：地区 / 行业 / 阶段 / 全文 / 附件。正文不会变，永久缓存"""
    d = cache_get("content", bid)
    if d is not None:
        _CALLS["cached"] += 1
        return d
    d = post("getZTBProjectDetail", {"id": bid, "publishTime": pt}, key).get("data") or {}
    if d:
        cache_put("content", bid, d)
    return d


# ── 搜索 ────────────────────────────────────────────────────────────────
def norm_money(s):
    return re.sub(r"\s+", "", strip_html(s) or "")


# 采集源会给同一条标加各种前缀标记：[公开]【重要】更新- 等
_TITLE_PREFIX_RE = re.compile(r"^(?:[\[\【（(][^\]\】）)]{0,8}[\]\】）)]|更新|转发|补充)[-—:：]?\s*")
_TITLE_NOISE_RE = re.compile(r"[\s　·・.,，、。;；:：!！?？'\"“”‘’()（）\[\]【】<>《》/\\|_-]+")
# 同一条标常被采成「XXX」与「XXX-竞争性磋商公告」两条 → 去重比对时剥掉尾部公告类后缀。
# 只剥尾部，不动中间：「…项目01、02、04-09包招标公告」与「…公开招标公告」正文不同，仍判为两条。
_TITLE_SUFFIX_RE = re.compile(
    r"(?:的?(?:成交|中标|采购|询价|竞价|磋商|谈判|遴选|资格预审|更正|变更|澄清|答疑|终止|流标|废标)?"
    r"(?:结果|候选人)?(?:公告|公示|公开|通知|announcement))+$"
)
_TITLE_TAIL_RE = re.compile(
    r"(?:竞争性磋商|竞争性谈判|竞争性|公开招标|邀请招标|单一来源|询价采购|重新招标|意向公开)$")


def norm_title(t):
    """标题归一化（仅用于去重比对，输出仍用原标题）"""
    t = strip_html(t)
    prev = None
    while prev != t:                       # 可能叠了多个前缀
        prev = t
        t = _TITLE_PREFIX_RE.sub("", t)
    t = _TITLE_NOISE_RE.sub("", t)
    prev = None
    while prev != t:                       # 「…短视频制作竞争性磋商公告」要剥两层
        prev = t
        t = _TITLE_TAIL_RE.sub("", _TITLE_SUFFIX_RE.sub("", t))
    return t


def dedupe(rows):
    """按（标题 + 甲方 + 金额）归并。同一项目常被采集成多条，发布时间差几十分钟。"""
    seen, out, dropped = {}, [], 0
    for r in rows:
        title = norm_title(r["title"])
        key = (title, (r["party_a"] or ""), norm_money(r["money"]))
        if key in seen:
            hit = seen[key]
            hit["dup_count"] += 1
            if r["publish_time"] < hit["publish_time"]:
                hit["publish_time"] = r["publish_time"]  # 保留最早发布的时间
            dropped += 1
            continue
        r["dup_count"] = 1
        seen[key] = r
        out.append(r)
    return out, dropped


def shape(raw):
    cid = str(raw.get("projectClassID") or "")
    return {
        "id": raw.get("id"),
        "publish_time": raw.get("publishTime") or "",
        "title": strip_html(raw.get("title")),          # 原样输出，禁缩写/截断
        "snippet": strip_html(raw.get("content"))[:220],
        "party_a": (raw.get("partANameList") or [None])[0],
        "money": strip_html(raw.get("projectMoney")) or "未注明",
        "class_id": cid,
        "class_name": CLASS_NAMES.get(cid, "未知(%s)" % cid),
        "class_note": CLASS_NOTES.get(cid, ""),
        "area": area_name_of(raw.get("proviceCode"), raw.get("cityCode")) or "未注明",
        "area_code": {"province": raw.get("proviceCode"), "city": raw.get("cityCode"),
                      "county": raw.get("countyCode")},
        "has_file": raw.get("hasFile"),
        # 🔴 列表接口没有报名截止日期字段。绝不拿 publishTime 冒充（实测踩过）。
        "signup_deadline": "未注明（列表接口无此字段，需拉详情）",
    }


def match_pos(title, keyword):
    """关键词命中位置：「标题」/「仅正文」。
    keyword 语法同接口：竖线=或，空格=同时出现。只要有一组词全部出现在标题里就算标题命中。
    仅正文命中的，多半是招标文件里的套话（如「投标文件的内容制作」），是噪音的主要来源。"""
    if not keyword:
        return "标题"
    t = (title or "").lower()
    for alt in keyword.split("|"):
        terms = [x for x in alt.lower().split() if x]
        if terms and all(x in t for x in terms):
            return "标题"
    return "仅正文"


def run_search(args):
    """执行一次检索，返回 (meta, rows)。同样的条件 SEARCH_TTL_MIN 分钟内复用上次结果。"""
    key = None if args.replay else get_key()   # 回放模式不需要 key
    start, end, how = resolve_window(args)

    # 地区 → 6 位码
    prov_codes, city_codes, area_report, unresolved = [], [], [], []
    for name in args.area or []:
        level, code = resolve_area(name)
        if level == "province":
            prov_codes.append(code)
            area_report.append("%s(省级 %s)" % (name, code))
        elif level == "city":
            city_codes.append(code)
            area_report.append("%s(市级 %s)" % (name, code))
        else:
            unresolved.append(name)
    for code in args.area_code or []:
        if not re.match(r"^\d{6}$", code):
            die("--area-code 必须是 6 位数字，收到：%s" % code)
        (prov_codes if code.endswith("0000") else city_codes).append(code)
        area_report.append("直传 %s" % code)
    if unresolved:
        die("无法识别的地区：%s（用 `radar.py areas <关键字>` 查可用地名，"
            "或用 --area-code 直传 6 位码）" % "、".join(unresolved))

    if args.keyword and len(args.keyword) > 300:
        die("keyword 超过 300 字符上限")
    if args.include and len(args.include) > 150:
        die("--include 超过 150 字符上限")

    # 排除词（去重、保序）
    terms = []
    for t in (args.exclude or "").split("|"):
        t = t.strip()
        if t and t not in terms:
            terms.append(t)
    exclude_kw, dropped_terms = "", []
    for t in terms:
        candidate = (exclude_kw + "|" + t) if exclude_kw else t
        if len(candidate) > EXCLUDE_MAXLEN:
            dropped_terms.append(t)      # 超限只丢弃多出来的，不整条报错
        else:
            exclude_kw = candidate

    purchase = PURCHASE_TYPES.get(args.purchase or "全部")
    if purchase is None:
        die("--purchase 只能是：%s" % " / ".join(PURCHASE_TYPES))

    # 探量已并入正式搜索：正式搜的返回里本来就有 total，单独探量等于多花一次请求
    size = min(args.size, PAGE_MAX)
    payload = {
        "startDate": start,
        "endDate": end,
        "pageId": args.page,
        "pageNumber": size,
        "searchType": SEARCH_TYPE,     # 锁死
        "searchMode": SEARCH_MODE,     # 锁死
        "keyword": args.keyword or "",
        "inCludeKW": args.include or "",
        "excludeKW": exclude_kw,
        "projectClassID": args.classes,
        "purchaseTypeID": purchase,
    }
    if prov_codes or city_codes:
        payload["areaCode"] = {"proviceCodeList": prov_codes,
                               "cityCodeList": city_codes, "countyCodeList": []}
    if args.money_min:
        payload["projectMoneyMin"] = str(args.money_min)
    if args.money_max:
        payload["projectMoneyMax"] = str(args.money_max)

    if args.replay:
        # 离线回放：拿之前 --dump-raw 存下的原始响应重跑筛选逻辑，不调接口
        try:
            with open(args.replay, encoding="utf-8") as f:
                resp = json.load(f)
        except OSError as e:
            die("读不到回放文件 %s：%s" % (args.replay, e))
    else:
        sig = _sig(payload)
        resp = cache_get("search", sig, SEARCH_TTL_MIN * 60)
        if resp is not None:
            _CALLS["cached"] += 1
        else:
            if getattr(args, "guard", False):
                guard_search()
            try:
                resp = post("searchProjectApi", payload, key)
            except ApiError as e:
                die(str(e), code=3)
            cache_put("search", sig, resp)
            if getattr(args, "guard", False):
                record_search(payload)
        if args.dump_raw:
            with open(args.dump_raw, "w", encoding="utf-8") as f:
                json.dump(resp, f, ensure_ascii=False, indent=2)

    d = resp.get("data") or {}
    meta = {
        "ok": True,
        "engine": "bidwin-radar v%s" % VERSION,
        # 用户回复里必须写明本次检索的日期范围，便于当场核对
        "date_range": "%s ~ %s" % (start, end),
        "date_basis": how,
        "server_echo_range": "%s ~ %s" % (d.get("startDate") or d.get("startdate") or "?",
                                          d.get("endDate") or d.get("enddate") or "?"),
        "area_filter": area_report or ["全国（未限定）"],
        "class_filter": [CLASS_NAMES.get(c.strip(), c.strip()) for c in args.classes.split(",")]
        if args.classes != "-100" else ["全部分类"],
        "purchase_filter": {"0": "其它类", "1": "服务类", "2": "工程类", "3": "货物类",
                            "-100": "全部采购类型"}.get(purchase, purchase),
        "exclude_kw_used": exclude_kw or "（无）",
        "total": d.get("total"),
        "page": d.get("pageId"),
        "page_size": d.get("pageNumber"),
        "has_next": d.get("hasNext"),
        "probe": bool(args.probe),
        "coverage": coverage_of(d.get("data") or [], start, end, d.get("hasNext")),
        "api_calls": _CALLS["n"],
        "from_cache": _CALLS["cached"] > 0,
    }
    if args.probe:
        meta["probe_note"] = ("探量已并入正式搜索，结果一并返回；"
                              "随后用同样条件正式搜会直接复用这次结果，不再请求")
    if getattr(args, "guard", False) and not args.replay:
        hint = repeat_hint(payload)
        if hint:
            meta["repeat_hint"] = hint
    if not args.replay:
        remember_ids(d.get("data") or [])

    rows = [shape(r) for r in (d.get("data") or [])]
    raw_n = len(rows)

    # 结果类剔除（实测「混入中标公告」）
    result_dropped = 0
    if not args.allow_result:
        keep = [r for r in rows if r["class_id"] not in RESULT_CLASS_IDS]
        result_dropped = len(rows) - len(keep)
        rows = keep

    rows, dup_dropped = dedupe(rows)

    # 标命中位置；标题命中排前，仅正文命中排后（各自保持发布时间倒序）
    for r in rows:
        r["match_pos"] = match_pos(r["title"], args.keyword)
    rows.sort(key=lambda r: 0 if r["match_pos"] == "标题" else 1)

    if args.limit:
        rows = rows[: args.limit]

    meta["counts"] = {
        "本页返回": raw_n, "剔除结果类": result_dropped, "去重合并": dup_dropped,
        "标题命中": sum(1 for r in rows if r["match_pos"] == "标题"),
        "仅正文命中": sum(1 for r in rows if r["match_pos"] == "仅正文"),
        "最终输出": len(rows),
    }
    if dropped_terms:
        meta["warning"] = ("排除词合计超过 %d 字符上限，以下词未生效：%s"
                           % (EXCLUDE_MAXLEN, "、".join(dropped_terms)))
    if profile_empty(load_profile()):
        meta["profile_hint"] = PROFILE_HINT
    meta["rules_applied"] = [
        "日期由北京时间计算", "精准匹配",
        "HTML 标签已删除", "已按 标题+甲方+金额 去重",
        "截止日期一律「未注明」，禁用发布时间冒充",
    ]
    return meta, rows


def cmd_search(args):
    args.guard = True        # 用户 / Agent 直接发起的搜索才计数；推送内部的不计
    meta, rows = run_search(args)
    emit(meta, args, rows)


# ── 详情 ────────────────────────────────────────────────────────────────
def _flat(v, sep="、"):
    """结构化接口大量字段是数组（多编号/多金额/多联系人）→ 拍平成人话字符串"""
    if v in (None, "", []):
        return None
    if isinstance(v, list):
        parts = [_flat(x) for x in v]
        parts = [p for p in parts if p]
        return sep.join(parts) or None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return strip_html(str(v)) or None


def money_cn(v):
    """元 → 「xx万元」，便于人读；拿不准就原样返回"""
    s = _flat(v)
    if not s:
        return None
    try:
        n = float(str(v[0] if isinstance(v, list) else v))
    except (TypeError, ValueError):
        return s
    return "%s万元（%s元）" % (("%.2f" % (n / 10000)).rstrip("0").rstrip("."), format(int(n), ","))


def countdown_of(deadline):
    """报名截止 → 剩 N 天。⚠️只接受真实的截止字段，绝不接受发布时间"""
    if not deadline:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(deadline.strip()[: 19 if " " in fmt else 10], fmt)
        except ValueError:
            continue
        days = (dt.date() - now_cn().date()).days
        if days < 0:
            return "⚠️已截止（%d 天前）" % -days
        return "今天截止" if days == 0 else ("⚠️仅剩 %d 天" % days if days <= 3 else "剩 %d 天" % days)
    return None


def area_name_of(prov_code, city_code):
    parts = [_PROV_BY_CODE.get(prov_code), _CITY_BY_CODE.get(city_code)]
    return " ".join(x for x in parts if x) or None


def _party(info):
    """partyAInfo[0] → 甲方 / 联系人 / 电话 / 地址"""
    if not info or not isinstance(info, list):
        return {}
    p = info[0] or {}
    return {
        "name": _flat(p.get("name")),
        "contact": _flat(p.get("contactName")),
        "phone": _flat(p.get("contactPhone")),
        "address": _flat(p.get("address")),
        "email": _flat(p.get("email")),
    }


def cmd_detail(args):
    key = get_key()
    out = {"ok": True, "engine": "bidwin-radar v%s" % VERSION, "id": args.id}
    (_, pt), = resolve_ids([(args.id, args.publish_time or "")])
    try:
        # 一条标要取两份才够字段：结构化(截止/预算/联系人) + 正文(地区/行业/阶段/全文)
        st = get_struct(args.id, pt, key)
        ct = {} if args.brief else get_content(args.id, pt, key)
    except ApiError as e:
        die(str(e), code=3)

    deadline = _flat(st.get("siginUpStopDate"))
    a = _party(st.get("partyAInfo"))
    agent = _party(st.get("agencyInfo"))
    cid = str(ct.get("projectClassID") or "")

    out["fields"] = {
        "项目名称": _flat(st.get("projectName")) or _flat(ct.get("title")) or "未注明",
        "项目编号": _flat(st.get("projectNumber")) or "未注明",
        "甲方": a.get("name") or _flat(ct.get("partAName")) or "未注明",
        "甲方联系人": a.get("contact") or "未注明",
        "甲方电话": a.get("phone") or "未注明",
        "代理机构": agent.get("name") or _flat(ct.get("agentName")) or "未注明",
        "代理联系人": " ".join(x for x in [agent.get("contact"), agent.get("phone")] if x) or "未注明",
        "预算": money_cn(st.get("budgetMoney")) or _flat(ct.get("projectMoney")) or "未注明",
        # 🔴 空就是「未注明」。禁止拿发布时间冒充截止日期（2026-07 实测踩过）
        "报名截止": deadline or "未注明",
        "截止倒计时": countdown_of(deadline) or "未注明（无截止日期字段，别猜）",
        "开标时间": _flat(st.get("bidStartDate")) or "未注明",
        "开标地点": _flat(st.get("bidStartAddress")) or "未注明",
        "所在地": area_name_of(ct.get("proviceCode"), ct.get("cityCode")) or "未注明",
        "行业": _flat(ct.get("industryName")) or "未注明",
        "阶段": CLASS_NAMES.get(cid, "未注明"),
        "阶段提示": CLASS_NOTES.get(cid, ""),
        "原文链接": _flat(st.get("collectUrl")) or "未注明",
        "附件数": "未查" if args.brief else len(ct.get("projectFiles") or []),
    }
    if args.with_content:
        out["content"] = strip_html(ct.get("content") or "")[: args.max_chars]
    if ct:
        # 正文已经拉了，资质比对不多花请求；用全文比对，不受 --max-chars 截断影响
        prof = load_profile()
        m = screen_one(strip_html(ct.get("content") or ""), prof, ct.get("proviceCode"), ct.get("cityCode"))
        if profile_empty(prof):
            # 用户跳过了问卷：不展示对照、不追问。提示由查标结果统一放在回复末尾，不跟着单条标走
            out["match"] = {"profile_empty": True}
        else:
            m["profile_empty"] = False
            m["rule"] = MATCH_RULE
            out["match"] = m
    out["reminder"] = "字段为空一律输出「未注明」，不得编造，尤其不得用发布时间当报名截止日期"
    out["api_calls"] = _CALLS["n"]
    out["from_cache"] = _CALLS["cached"]
    print(json.dumps(out, ensure_ascii=False, indent=2))


# ── 地区查询 ────────────────────────────────────────────────────────────
def cmd_areas(args):
    q = (args.query or "").strip()
    res = []
    for level, table in (("省级", PROVINCES), ("市级", CITIES)):
        for name, code in table.items():
            if not q or q in name or _norm_area(q) == _norm_area(name):
                res.append({"level": level, "name": name, "code": code})
    print(json.dumps({"ok": True, "query": q or "(全部)", "count": len(res),
                      "areas": res[:200]}, ensure_ascii=False, indent=2))


def emit(meta, args, rows=None):
    out = dict(meta)
    if rows is not None:
        out["results"] = rows
    text = json.dumps(out, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        print(json.dumps({"ok": True, "written": args.out, "count": len(rows or []),
                          "date_range": meta.get("date_range"), "total": meta.get("total")}, ensure_ascii=False, indent=2))
    else:
        print(text)




# ── 环境 / key ──────────────────────────────────────────────────────────
def cmd_setkey(args):
    k = (args.key or "").strip()
    if not k or any(c.isspace() for c in k):
        die("key 为空或含空格，请让用户重新复制完整的 key")
    p = key_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write(k + "\n")
    try:
        os.chmod(p, 0o600)      # Windows 上无效，忽略
    except OSError:
        pass
    print(json.dumps({"ok": True, "saved_to": p, "key": mask(k)}, ensure_ascii=False, indent=2))


def cmd_check(args):
    import platform
    k = os.environ.get("BIDWIN_KEY", "").strip()
    src = "环境变量 BIDWIN_KEY" if k else None
    if not k:
        try:
            with open(key_path(), encoding="utf-8-sig") as f:
                k = f.read().strip().splitlines()[0].strip()
            src = key_path()
        except (OSError, IndexError):
            k = ""
    print(json.dumps({
        "ok": sys.version_info >= (3, 8),
        "engine": "bidwin-radar v%s" % VERSION,
        "os": platform.system(),
        "python": platform.python_version(),
        "python_ok": sys.version_info >= (3, 8),
        "key_configured": bool(k),
        "key": mask(k) if k else None,
        "key_source": src,
        "profile_exists": os.path.exists(profile_path()),
        "beijing_now": now_cn().strftime("%Y-%m-%d %H:%M:%S"),
        "note": "本命令不调接口",
    }, ensure_ascii=False, indent=2))


# ── 每日推送：订阅条件（只存在本机 ~/.bidwin/watch.json，不上传）──────────────
WATCH_FILE = "~/.bidwin/watch.json"
WATCH_MAX = 3            # 最多保存几组条件
WATCH_DAYS = 3           # 每次推送查近几天（只查当天会漏前两天发布、仍可投的标）
SEEN_KEEP_DAYS = 60      # 推送记录保留多久


def watch_path():
    return os.path.expanduser(WATCH_FILE)


def load_watch():
    try:
        with open(watch_path(), encoding="utf-8") as f:
            d = json.load(f)
    except (OSError, ValueError):
        d = {}
    d.setdefault("watches", [])
    d.setdefault("seen", {})
    return d


def save_watch(d):
    p = watch_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)


def _watch_view(w):
    parts = ["主题：%s" % w["keyword"]]
    if w.get("include"):
        parts.append("必含：%s" % w["include"])
    if w.get("exclude"):
        parts.append("排除：%s" % w["exclude"])
    parts.append("地区：%s" % ("、".join(w["area"]) if w.get("area") else "全国"))
    if w.get("money_min") or w.get("money_max"):
        parts.append("金额：%s ~ %s 元" % (w.get("money_min") or "不限", w.get("money_max") or "不限"))
    if w.get("purchase") and w["purchase"] != "全部":
        parts.append("采购类型：%s" % w["purchase"])
    return {"name": w["name"], "conditions": "；".join(parts), "created": w.get("created")}


def cmd_watch_set(args):
    if not args.keyword:
        die("--keyword 必填")
    if PURCHASE_TYPES.get(args.purchase or "全部") is None:
        die("--purchase 只能是：%s" % " / ".join(PURCHASE_TYPES))
    for name in args.area or []:
        if resolve_area(name)[0] is None:
            die("无法识别的地区：%s（先用 areas 命令查正确地名）" % name)
    d = load_watch()
    w = {
        "name": args.name.strip(), "keyword": args.keyword, "include": args.include or "",
        "exclude": args.exclude or "", "area": args.area or [], "money_min": args.money_min,
        "money_max": args.money_max, "purchase": args.purchase or "全部",
        "created": now_cn().strftime("%Y-%m-%d %H:%M:%S"),
    }
    others = [x for x in d["watches"] if x["name"] != w["name"]]
    if len(others) >= WATCH_MAX:
        die("最多保存 %d 组推送条件，已满：%s。请先删除一组（watch delete --name <名称>）"
            % (WATCH_MAX, "、".join(x["name"] for x in others)))
    replaced = len(others) != len(d["watches"])
    d["watches"] = others + [w]
    save_watch(d)
    print(json.dumps({"ok": True, "action": "已更新" if replaced else "已保存",
                      "watch": _watch_view(w), "total_watches": len(d["watches"]),
                      "saved_to": watch_path(), "note": "只保存在本机，不上传"},
                     ensure_ascii=False, indent=2))


def cmd_watch_list(args):
    d = load_watch()
    print(json.dumps({"ok": True, "count": len(d["watches"]),
                      "watches": [_watch_view(w) for w in d["watches"]],
                      "pushed_records": len(d["seen"])}, ensure_ascii=False, indent=2))


def cmd_watch_delete(args):
    d = load_watch()
    left = [x for x in d["watches"] if x["name"] != args.name]
    if len(left) == len(d["watches"]):
        die("没有名为「%s」的推送条件" % args.name)
    d["watches"] = left
    save_watch(d)
    print(json.dumps({"ok": True, "deleted": args.name, "remaining": len(left)},
                     ensure_ascii=False, indent=2))


def cmd_watch_run(args):
    """按已保存的条件查近 N 天，跳过以前推送过的标，只返回新标。"""
    d = load_watch()
    if not d["watches"]:
        die("还没有保存任何推送条件。请先让用户设置每日推送", code=5)
    today = now_cn().strftime("%Y-%m-%d")
    cutoff = (now_cn() - timedelta(days=SEEN_KEEP_DAYS)).strftime("%Y-%m-%d")
    seen = {k: v for k, v in d["seen"].items() if v >= cutoff}
    out, shown, date_range = [], set(), None
    for w in d["watches"]:
        ns = argparse.Namespace(
            keyword=w["keyword"], include=w.get("include") or None, exclude=w.get("exclude") or None,
            days=args.days, last=None, start=None, end=None, area=w.get("area") or None,
            area_code=None, classes=DEFAULT_CLASS_IDS, purchase=w.get("purchase") or "全部",
            money_min=w.get("money_min"), money_max=w.get("money_max"), page=1, size=PAGE_MAX,
            probe=False, allow_result=False, limit=None, out=None, dump_raw=None,
            replay=args.replay,
        )
        meta, rows = run_search(ns)
        date_range = meta["date_range"]
        new = []
        for r in rows:
            k = str(r["id"])
            if k in seen or k in shown:
                continue
            shown.add(k)
            new.append(r)
        cov = meta["coverage"]
        out.append({
            "name": w["name"], "conditions": _watch_view(w)["conditions"], "total": meta["total"],
            "has_more_than_one_page": bool(meta.get("has_next")),
            "coverage": cov,
            "coverage_warning": None if cov["complete"] else (
                "这组条件近 %d 天的新标超过一页，%s，会漏掉一部分；建议在推送结果后提醒用户收窄条件"
                % (args.days, cov["text"])),
            "already_pushed_skipped": len(rows) - len(new), "new_count": len(new), "results": new,
        })
    if not args.dry_run:
        for k in shown:
            seen[k] = today
        d["seen"] = seen
        save_watch(d)
    prof = load_profile()
    empty = profile_empty(prof)
    hint = None
    if empty and not prof.get("push_hint_shown"):
        hint = "填写公司基本情况后，推送可以只给你门槛够得上的标。随时发送「填写我的资质」就行。"
        if not args.dry_run:
            prof["push_hint_shown"] = True
            save_profile(prof)
    print(json.dumps({"ok": True, "engine": "bidwin-radar v%s" % VERSION, "date_range": date_range,
                      "date_basis": "近 %d 天（含今天），已推送过的标不再重复" % args.days,
                      "dry_run": bool(args.dry_run),
                      "api_calls": _CALLS["n"],
                      "new_total": sum(x["new_count"] for x in out),
                      "screen": {
                          # 列表没有截止日期：不管有没有资质档案，都要走 screen 去掉已过报名截止的标
                          "required": True,
                          "match_enabled": not empty,
                          "push_mode": prof.get("push_mode", "fold"),
                          "next_step": ("按相关性筛完后，把保留的标用 screen --ids 查报名截止"
                                        + ("并做资质精筛" if not empty else "（档案为空，只查截止，每条 1 次请求）")
                                        + "；expired 的不推送"),
                          "profile_hint": hint,
                      },
                      "watches": out},
                     ensure_ascii=False, indent=2))


# ── 资质档案（只存在本机 ~/.bidwin/profile.json，不上传）────────────────────
# 安装时只问 4 项基本门槛（成立时间 / 注册资本 / 参保人数 / 注册地），可跳过。
# 资质证书、业绩不设问卷：遇到哪条标要求了再问一句「有 / 没有」，答案记下来，下次不再问。
# 「没有」也要记——记住没有，才能在下次准确判「不符合」；只记「有」的话名字对不上就只能判「需确认」。
PROFILE_FILE = "~/.bidwin/profile.json"
PUSH_MODES = {"fold": "不符合的折叠成一行（默认）", "all": "全部展示，不符合的标注出来"}
SCREEN_MAX = 30          # screen 一次最多筛几条（每条 2 次请求，限速 60 次/分钟）


def profile_path():
    return os.path.expanduser(PROFILE_FILE)


def load_profile():
    try:
        with open(profile_path(), encoding="utf-8") as f:
            d = json.load(f)
    except (OSError, ValueError):
        d = {}
    d.setdefault("basics", {})
    d.setdefault("certs", {})
    d.setdefault("cases", {})
    d.setdefault("ask", True)
    d.setdefault("push_mode", "fold")
    return d


def save_profile(d):
    p = profile_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    d["updated"] = now_cn().strftime("%Y-%m-%d %H:%M:%S")
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)


def profile_empty(d):
    return not (d.get("basics") or d.get("certs") or d.get("cases"))


def _yn(v):
    return "有" if v else "没有"


def _profile_view(d):
    b = d.get("basics") or {}
    return {
        "成立时间": b.get("established") or "未填",
        "注册资本": ("%g 万元" % b["capital_wan"]) if b.get("capital_wan") is not None else "未填",
        "参保人数": ("%d 人" % b["headcount"]) if b.get("headcount") is not None else "未填",
        "注册地": " ".join(x for x in [b.get("reg_prov"), b.get("reg_city")] if x) or "未填",
        "资质证书": {k: _yn(v) for k, v in (d.get("certs") or {}).items()} or "暂无记录（遇到要求时再问）",
        "业绩案例": {k: _yn(v) for k, v in (d.get("cases") or {}).items()} or "暂无记录（遇到要求时再问）",
        "追问": "开" if d.get("ask", True) else "关（只标注，不再追问）",
        "推送精筛": PUSH_MODES.get(d.get("push_mode", "fold")),
    }


_EST_RE = re.compile(r"^(\d{4})(?:-(\d{1,2}))?(?:-(\d{1,2}))?$")


def _est_bounds(s):
    """成立时间 → (最早可能日期, 最晚可能日期)。用户常只记得年份或年月，按区间判定，不硬猜。"""
    m = _EST_RE.match((s or "").strip())
    if not m:
        return None, None
    y, mo, dd = int(m.group(1)), m.group(2), m.group(3)
    try:
        if dd:
            d = datetime(y, int(mo), int(dd)).date()
            return d, d
        if mo:
            first = datetime(y, int(mo), 1).date()
            nxt = datetime(y + (int(mo) == 12), int(mo) % 12 + 1, 1).date()
            return first, nxt - timedelta(days=1)
        return datetime(y, 1, 1).date(), datetime(y, 12, 31).date()
    except ValueError:
        return None, None


def cmd_profile_set(args):
    d = load_profile()
    b = d["basics"]
    changed = []
    if args.established:
        lo, _ = _est_bounds(args.established)
        if lo is None:
            die("--established 格式应为 YYYY / YYYY-MM / YYYY-MM-DD，收到：%s" % args.established)
        if lo > now_cn().date():
            die("成立时间晚于今天：%s" % args.established)
        b["established"] = args.established.strip()
        changed.append("成立时间")
    if args.capital_wan is not None:
        if args.capital_wan <= 0:
            die("--capital-wan 必须大于 0（单位：万元）")
        b["capital_wan"] = args.capital_wan
        changed.append("注册资本")
    if args.headcount is not None:
        if args.headcount < 0:
            die("--headcount 不能为负数")
        b["headcount"] = args.headcount
        changed.append("参保人数")
    if args.reg_area:
        level, code = resolve_area(args.reg_area)
        if level is None:
            die("无法识别的注册地：%s（填到省或地级市即可，如「上海」「杭州」；先用 areas 命令查）"
                % args.reg_area)
        if level == "province":
            b["reg_prov"], b["reg_city"] = _PROV_BY_CODE[code], None
        else:
            b["reg_prov"], b["reg_city"] = _PROV_BY_CODE.get(code[:2] + "0000"), _CITY_BY_CODE[code]
        changed.append("注册地")
    if args.ask:
        d["ask"] = args.ask == "on"
        changed.append("追问")
    if args.push_mode:
        d["push_mode"] = args.push_mode
        changed.append("推送精筛")
    if not changed:
        die("没有要修改的项")
    save_profile(d)
    print(json.dumps({"ok": True, "updated": changed, "profile": _profile_view(d),
                      "saved_to": profile_path(), "note": "只保存在本机，不上传"},
                     ensure_ascii=False, indent=2))


def _cmd_profile_item(args, bucket, label):
    d = load_profile()
    name = (args.name or "").strip()
    if not name:
        die("--name 不能为空")
    if args.remove:
        if name not in d[bucket]:
            die("档案里没有「%s」" % name)
        d[bucket].pop(name)
        action = "已删除"
    else:
        if args.has not in ("yes", "no"):
            die("--has 只能是 yes 或 no")
        d[bucket][name] = args.has == "yes"
        action = "已记录：%s「%s」" % (_yn(d[bucket][name]), name)
    save_profile(d)
    print(json.dumps({"ok": True, "kind": label, "action": action,
                      label: {k: _yn(v) for k, v in d[bucket].items()}},
                     ensure_ascii=False, indent=2))


def cmd_profile_cert(args):
    _cmd_profile_item(args, "certs", "资质证书")


def cmd_profile_case(args):
    _cmd_profile_item(args, "cases", "业绩案例")


def cmd_profile_show(args):
    d = load_profile()
    print(json.dumps({"ok": True, "empty": profile_empty(d), "profile": _profile_view(d),
                      "saved_to": profile_path()}, ensure_ascii=False, indent=2))


# ── 硬门槛比对（移植自 BidWin screen）──────────────────────────────────────
# 🔴 能算的全在代码里算完，模型只念 match_line 成品句，不得自行改写 checks。
#    BidWin 实测：模型把「未提及」译成了「未核对」，意思完全相反。
# 每项五态：满足 / 不满足 / 未提及（查过了公告没写）/ 待补充（公告有要求、档案没这项）/ 需确认。
# 只有「不满足」会导致折叠，而且只由确定性的项触发；证书名对不上、业绩这类语义判断永远不折叠。
_NUM = r"([0-9][0-9,，.]*)"
_UNIT = r"(亿元|亿|万元|万|元)"
# 「要求」的引导词——没有这些的多半是在描述别人，不是对投标人的硬性要求
_REQ = r"(?:须|必须|应|应当|需|要求|满|不得低于|不得少于|不低于|不少于|不小于|达到|具备|拥有|持有|≥|>=)"
_MUNICIPALITIES = {"北京市", "上海市", "天津市", "重庆市"}


def _to_wan(num, unit):
    """金额统一换算成「万元」"""
    try:
        v = float(str(num).replace(",", "").replace("，", ""))
    except ValueError:
        return None
    if unit in ("亿元", "亿"):
        return v * 10000
    if unit in ("万元", "万"):
        return v
    return v / 10000          # 元


def _ev(text, m, span=40):
    """截取命中处上下文当证据"""
    a = max(0, m.start() - span)
    b = min(len(text), m.end() + span)
    return re.sub(r"\s+", " ", text[a:b]).strip()


def _req_capital(text):
    """注册资本要求 → (要求万元, evidence)"""
    for pat in (r"(?:注册|实缴)资[本金][^。；;\n]{0,30}?" + _REQ + r"[^。；;\n]{0,15}?" + _NUM + r"\s*" + _UNIT,
                _REQ + r"[^。；;\n]{0,10}?(?:注册|实缴)资[本金][^。；;\n]{0,15}?" + _NUM + r"\s*" + _UNIT):
        m = re.search(pat, text)
        if m:
            wan = _to_wan(m.group(1), m.group(2))
            if wan:
                return wan, _ev(text, m)
    return None, None


def _req_years(text):
    """成立年限要求 → (年, evidence)"""
    for pat in (r"(?:成立|注册|经营)[^。；;\n]{0,20}?" + _REQ + r"[^。；;\n]{0,10}?([0-9]+)\s*年",
                _REQ + r"[^。；;\n]{0,12}?(?:成立|注册|经营)[^。；;\n]{0,10}?([0-9]+)\s*年",
                r"([0-9]+)\s*年以上[^。；;\n]{0,10}?(?:经营|成立|注册)"):
        m = re.search(pat, text)
        if m:
            try:
                return int(m.group(1)), _ev(text, m)
            except ValueError:
                pass
    return None, None


def _req_headcount(text):
    """参保人数要求 → (人数, evidence)"""
    m = re.search(r"(?:参保|社保|缴纳社会保险)[^。；;\n]{0,25}?" + _REQ + r"[^。；;\n]{0,10}?([0-9]+)\s*(?:人|名)", text)
    if m:
        try:
            return int(m.group(1)), _ev(text, m)
        except ValueError:
            pass
    return None, None


def _req_prequalified(text):
    """是否要求「事先入库/入围」才能投 → evidence
    实测坑：①「应」会匹配到「供应商」里的应；②「被列入不良行为供应商库」是排除失信企业，不是入库门槛"""
    for m in re.finditer(r"(?:须|必须|(?<!供)应|仅限|限于|应当)[^。；;\n]{0,25}?"
                         r"(?:入[围库]|供应商库|名录库|合格供应商名单|资格库)[^。；;\n]{0,15}", text):
        # ③「使用平台前完成供应商注册，加入平台供应商库」是谁都能办的平台注册，不是门槛
        if not re.search(r"列入|不良|黑名单|失信|禁止|违法|平台|注册", m.group(0)):
            return _ev(text, m)
    return None


_LOCAL_ALT = re.compile(r"或|分公司|分支机构|服务网点|网点|办事处|驻点|服务机构|常驻")


def _req_local(text):
    """是否要求本地注册 → (层级 province/city/district, evidence, 是否允许分支机构等替代)
    实测：「本地注册企业或设有固定服务网点」「非本市注册的须在本市注册分公司」——外地企业也能投，不能判不满足"""
    for pat in (
        # 「注册地…在本市」——本市后面不一定还有词
        r"(?:注册地|注册所在地|工商注册地)[^。；;\n]{0,15}?(?:在|为|位于|属于)?[^。；;\n]{0,6}?"
        r"(本市|本地|当地|本区|本县|本省)",
        # 「须在本市注册/工商登记」
        r"(本市|本地|当地|本区|本县|本省)[^。；;\n]{0,8}?(?:注册|工商登记)",
    ):
        m = re.search(pat, text)
        if m:
            w = m.group(1)
            level = "province" if w == "本省" else ("district" if w in ("本区", "本县") else "city")
            ev = _ev(text, m)
            return level, ev, bool(_LOCAL_ALT.search(ev))
    return None, None, False


def _find_cert(text, cert_list):
    """在正文里找这批证书里被要求的 → [(证书名, evidence)]"""
    hits = []
    for cert in cert_list or []:
        key = cert.replace("证", "").replace("经营", "")[:8]
        m = re.search(re.escape(cert), text) or (re.search(re.escape(key), text) if len(key) >= 5 else None)
        if m:
            hits.append((cert, _ev(text, m)))
    return hits


_SENT_RE = re.compile(r"[^。；;\n]+")
_CERT_WORDS = re.compile(r"资质|许可证|证书|认证|备案")
# 几乎每份公告都有的套话，不是真正的资质门槛
_CERT_NOISE = re.compile(r"营业执照|第二十二条|独立承担|信用|CA|数字证书|身份证|授权|社会保险|社保|纳税|税收|财务"
                         r"|违法|联合体|中小企业|残疾人|监狱|声明函|承诺函|签章|保证金|质疑|复印件|原件")
# 真正的资质门槛 = 主语是投标方 + 「具有 / 持有」某个具体的证。
# 实测噪音：「报价需含资质材料」（递交材料）、「供应商资质要求：」（空表单字段）、「算法通过网信办备案」（对产品的要求）
_CERT_SUBJECT = re.compile(r"供应商|投标人|响应人|报价人|申请人|比选人|磋商人|应答人|竞标人|投标单位")
_CERT_NAME = re.compile(r"(?:具有|具备|持有|取得)(?:有效的?|合法的?|国家)?"
                        r"(《?[^，,。；;、：:（(]{2,30}?(?:许可证|资质证书|认证证书|资质|认证|证书|备案)》?)")
_CERT_GENERIC = re.compile(r"^《?(?:相关|相应|有效|企业|投标|资格|上述|以下|所需|合法|必要|本项目|供应商|独立|法人)"
                           r"|相适应|相应的|相关的")


def _cert_required(s):
    if not _CERT_SUBJECT.search(s):
        return False
    # 「不具备…资质的」是排除条款，不是资质要求
    return any(not _CERT_GENERIC.search(m.group(1)) and s[max(0, m.start() - 1):m.start()] != "不"
               for m in _CERT_NAME.finditer(s))


_CASE_WORDS = re.compile(r"业绩|案例|类似项目|同类项目")
_SCORE_WORDS = re.compile(r"得分|加分|分值|评分|计分|每提供|最高得|最多得|得[0-9一二三四五]")
# 实测：「业绩资格审核条款：供应商需提供近 18 个月内…业绩」——门槛常写成「需提供 / 资格审核」
_GATE_WORDS = re.compile(r"须|必须|应当?具[有备]|[需须应]提供|否则|无效|不予|资格要求|资格条件|资格审[核查]|业绩要求")


def _clauses(text, word_re, keep, limit=3, span=70):
    """按句切，挑出含关键词且符合 keep 条件的句子（截取关键词附近），去重后最多 limit 条"""
    out = []
    for m in _SENT_RE.finditer(text):
        s = m.group(0)
        k = word_re.search(s)
        if not k or not keep(s):
            continue
        a = max(0, k.start() - span)
        frag = re.sub(r"\s+", " ", s[a:k.end() + span]).strip()
        if frag and all(frag not in x and x not in frag for x in out):
            out.append(frag)
        if len(out) >= limit:
            break
    return out


def _local_check(level, basics, bid_prov, bid_city):
    """本地注册比对 → (结论, 说明)"""
    rp, rc = basics.get("reg_prov"), basics.get("reg_city")
    where = " ".join(x for x in [bid_prov, bid_city] if x) or "未知"
    if not bid_prov:
        return "需确认", "公告要求本地注册，但标的所在地未知"
    if level == "province" or bid_prov in _MUNICIPALITIES:
        ok = rp == bid_prov
        return ("满足" if ok else "不满足"), "公告要求本地注册，标的在 %s，你们注册在 %s" % (where, rp)
    if rp != bid_prov:
        return "不满足", "公告要求本地注册，标的在 %s，你们注册在 %s" % (where, rp)
    if not bid_city or not rc:
        return "需确认", "公告要求本地注册，标的在 %s，请确认是否同一城市" % where
    if rc != bid_city:
        return "不满足", "公告要求本地注册，标的在 %s，你们注册在 %s" % (where, rc)
    if level == "district":
        return "需确认", "公告要求在本区县注册，标的在 %s，请确认是否同一区县" % where
    return "满足", "公告要求本地注册，标的在 %s，你们注册地相符" % where


def screen_one(text, prof, prov_code=None, city_code=None):
    """对单条标的正文做门槛比对。不调接口。"""
    b = prof.get("basics") or {}
    ask_on = prof.get("ask", True)
    checks, asks = {}, []

    def mark(name, state, detail=None, ev=None, ask=None):
        checks[name] = {"结论": state}
        if detail:
            checks[name]["说明"] = detail
        if ev:
            checks[name]["原文"] = ev
        if ask and ask_on:
            asks.append(ask)

    # 1) 注册资本
    need, ev = _req_capital(text)
    have = b.get("capital_wan")
    if need is None:
        mark("注册资本", "未提及")
    elif have is None:
        mark("注册资本", "待补充", "要求 %g 万" % need, ev,
             {"field": "capital_wan", "question": "这条要求注册资本不低于 %g 万，你们公司注册资本是多少万？" % need})
    elif have < need:
        mark("注册资本", "不满足", "要求 %g 万，你们 %g 万" % (need, have), ev)
    elif "实缴" in ev:
        # 档案里是注册资本，实缴未知 → 不判满足
        mark("注册资本", "需确认", "要求实缴 %g 万，你们注册资本 %g 万，请确认实缴金额" % (need, have), ev)
    else:
        mark("注册资本", "满足", "要求 %g 万，你们 %g 万" % (need, have), ev)

    # 2) 成立年限 —— 成立时间可能只精确到年/月，用最早、最晚两个端点判定
    need, ev = _req_years(text)
    lo, hi = _est_bounds(b.get("established"))
    if need is None:
        mark("成立年限", "未提及")
    elif lo is None:
        mark("成立年限", "待补充", "要求成立满 %d 年" % need, ev,
             {"field": "established", "question": "这条要求公司成立满 %d 年，你们是哪年成立的？" % need})
    else:
        today = now_cn().date()
        min_yrs = (today - hi).days / 365.25
        max_yrs = (today - lo).days / 365.25
        if min_yrs >= need:
            mark("成立年限", "满足", "要求 %d 年，你们约 %.1f 年" % (need, min_yrs), ev)
        elif max_yrs < need:
            mark("成立年限", "不满足", "要求 %d 年，你们约 %.1f 年" % (need, max_yrs), ev)
        else:
            mark("成立年限", "需确认", "要求满 %d 年，你们成立于 %s，差不多刚好卡线" % (need, b["established"]), ev,
                 {"field": "established", "question": "这条要求成立满 %d 年，你们刚好卡线，能告诉我具体成立日期吗？" % need})

    # 3) 参保人数
    need, ev = _req_headcount(text)
    have = b.get("headcount")
    if need is None:
        mark("参保人数", "未提及")
    elif have is None:
        mark("参保人数", "待补充", "要求 %d 人" % need, ev,
             {"field": "headcount", "question": "这条要求参保人数不少于 %d 人，你们公司现在大概多少人交社保？" % need})
    else:
        mark("参保人数", "满足" if have >= need else "不满足", "要求 %d 人，你们 %d 人" % (need, have), ev)

    # 4) 本地注册
    level, ev, alt = _req_local(text)
    bid_prov, bid_city = _PROV_BY_CODE.get(prov_code), _CITY_BY_CODE.get(city_code)
    if not level:
        mark("本地注册", "未提及")
    elif alt:
        where = " ".join(x for x in [bid_prov, bid_city] if x) or "当地"
        st = "需确认"
        if b.get("reg_prov"):
            st, _ = _local_check(level, b, bid_prov, bid_city)
            st = "满足" if st == "满足" else "需确认"
        mark("本地注册", st, "公告要求在 %s 本地注册，或在当地设分支机构 / 服务网点" % where, ev)
    elif not b.get("reg_prov"):
        where = " ".join(x for x in [bid_prov, bid_city] if x) or "当地"
        mark("本地注册", "待补充", "公告要求本地注册（标的在 %s）" % where, ev,
             {"field": "reg_area", "question": "这条要求本地注册（标的在 %s），你们公司注册在哪个城市？" % where})
    else:
        st, detail = _local_check(level, b, bid_prov, bid_city)
        mark("本地注册", st, detail, ev)

    # 5) 供应商库 —— 库各不相同，没法记在档案里，一律请用户自己确认
    ev = _req_prequalified(text)
    if ev:
        mark("供应商库", "需确认", "公告要求事先入库 / 入围的供应商才能投，请确认你们是否在该库中", ev)
    else:
        mark("供应商库", "未提及")

    # 6) 资质证书 —— 档案里记过的（有 / 没有）由代码判；没记过的交给模型读原文后问用户
    certs = prof.get("certs") or {}
    known = _find_cert(text, list(certs))
    lacking = [(c, e) for c, e in known if not certs[c]]
    held = [c for c, _ in known if certs[c]]
    if lacking:
        mark("资质证书", "不满足", "本项目要求《%s》，你们没有" % "》《".join(c for c, _ in lacking), lacking[0][1])
    elif held:
        mark("资质证书", "满足", "本项目要求的《%s》你们有" % "》《".join(held))
    else:
        mark("资质证书", "未提及")
    known_names = [c for c, _ in known]
    cert_clauses = _clauses(
        text, _CERT_WORDS,
        lambda s: _cert_required(s) and not _CERT_NOISE.search(s) and not any(c in s for c in known_names))

    # 7) 业绩 —— 先分「资格门槛」还是「评分加分项」；语义比对交给模型，永远不折叠
    case_gate = _clauses(text, _CASE_WORDS, lambda s: not _SCORE_WORDS.search(s) and _GATE_WORDS.search(s))
    case_score = _clauses(text, _CASE_WORDS, lambda s: bool(_SCORE_WORDS.search(s)), limit=2)

    extra = []
    if cert_clauses:
        extra.append("公告另有资质证书要求，待确认")
    if case_gate:
        extra.append("公告设有业绩门槛，待确认")
    fold = any(checks[k]["结论"] == "不满足" for k in checks)
    return {
        "fold": fold,
        "match_line": render_match(checks, extra),
        "checks": checks,
        "cert_clauses": cert_clauses,
        "case_gate_clauses": case_gate,
        "case_score_clauses": case_score,
        "known_certs": {k: _yn(v) for k, v in certs.items()},
        "known_cases": {k: _yn(v) for k, v in (prof.get("cases") or {}).items()},
        "ask_enabled": ask_on,
        "asks": asks,
    }


def render_match(checks, extra=()):
    """把 checks 渲染成可以直接抄进回复的成品句。只报异常，符合就一句话。"""
    def say(k, c):
        d = c.get("说明") or k
        # 数字类门槛的说明里没有项目名（「要求 100 万」），补上，否则用户不知道指什么
        return "%s%s" % (k, d) if k in ("注册资本", "成立年限", "参保人数") and not d.startswith(k) else d

    bad = [say(k, c) for k, c in checks.items() if c["结论"] == "不满足"]
    warn = [say(k, c) for k, c in checks.items() if c["结论"] in ("需确认", "待补充")]
    warn += list(extra)
    if bad:
        return "❌ 门槛不符：" + "；".join(bad)
    if warn:
        return "需确认：" + "；".join(warn)
    return "✅ 未发现与你们资质不符的门槛"


PROFILE_HINT = "填写公司基本情况后，我可以帮你判断每条标的门槛够不够得上。随时发送「填写我的资质」就行。"

MATCH_RULE = (
    "match_line 是成品句，原样抄进回复，不得改写或自行翻译 checks——「未提及」意思是查过了公告没写，不是「没核对」。"
    "cert_clauses / case_gate_clauses / case_score_clauses 是原文摘录，按 SKILL.md「资质匹配」处理；"
    "asks 为空或 ask_enabled 为 false 时不追问。"
)


def _expired(deadline):
    """报名截止已过 → True。没有截止日期的一律 False（不能当成已截止）。
    带时分秒的按时刻比：「今天 10:00 截止」中午推送时已经过了，countdown_of 只按天算会显示「今天截止」"""
    if not deadline:
        return False
    try:
        return datetime.strptime(deadline.strip()[:19], "%Y-%m-%d %H:%M:%S") < now_cn()
    except ValueError:
        return (countdown_of(deadline) or "").startswith("⚠️已截止")


def cmd_screen(args):
    """批量查报名截止 + 门槛比对（推送 / 用户要求「按资质筛一下」时用）。
    列表接口没有截止日期，只能逐条拉结构化详情；过期的在这里标出来，推送时不推。
    先查截止：已过期的不再拉正文。有资质档案：未过期的每条 2 次请求（截止 + 正文比对）；
    档案为空：每条 1 次（只查截止）。拉过的都进缓存，再筛或看详情不重复请求。"""
    key = get_key()
    prof = load_profile()
    empty = profile_empty(prof)
    items = []
    for item in re.split(r";;|,", args.ids or ""):
        item = item.strip()
        if not item:
            continue
        i, _, pt = item.partition("|")     # 兼容旧写法 id|publish_time
        try:
            items.append((int(i.strip()), pt.strip()))
        except ValueError:
            die("--ids 格式错误：%s（应为 id，多条用逗号分隔）" % item)
    if not items:
        die('没有要筛的标：--ids "id,id,id"')
    over = max(0, len(items) - SCREEN_MAX)
    pairs = resolve_ids(items[:SCREEN_MAX])

    out, ask_count = [], {}
    for idx, (bid, pt) in enumerate(pairs):
        before = _CALLS["n"]
        try:
            st = get_struct(bid, pt, key)
            deadline = _flat(st.get("siginUpStopDate"))
            expired = _expired(deadline)
            ct = {} if (empty or expired) else get_content(bid, pt, key)
        except ApiError as e:
            out.append({"id": bid, "publish_time": pt, "error": str(e), "expired": False, "fold": False,
                        "match_line": None if empty else "未能比对（拉取失败），不折叠"})
            continue
        row = {
            "id": bid,
            "publish_time": pt,
            "title": strip_html(st.get("projectName") or ct.get("title") or ""),
            "signup_deadline": deadline or "未注明",
            "countdown": countdown_of(deadline) or "未注明",
            "expired": expired,
        }
        if empty:
            row.update({"fold": False, "match_line": None})
        elif expired:
            # 已过报名截止的不拉正文、不比对，推送时直接不推
            row.update({"fold": False, "match_line": "已过报名截止，未做资质比对"})
        else:
            m = screen_one(strip_html(ct.get("content") or ""), prof, ct.get("proviceCode"), ct.get("cityCode"))
            for a in m["asks"]:
                ask_count.setdefault(a["field"], [0, a])[0] += 1
            row.update(m)
        out.append(row)
        if _CALLS["n"] > before and idx < len(pairs) - 1:
            time.sleep(1.1)          # 限速 60 次/分钟；全从缓存取的不用等

    # 同一个缺项在多条标里出现，只问一次；按出现次数挑最多 2 个
    ask_next = [dict(v[1], hits=v[0]) for v in sorted(ask_count.values(), key=lambda v: -v[0])][:2]
    print(json.dumps({
        "ok": True, "engine": "bidwin-radar v%s" % VERSION,
        "push_mode": prof.get("push_mode", "fold"),
        "profile_empty": empty,
        "summary": {"送检": len(pairs), "已截止": sum(1 for r in out if r.get("expired")),
                    "门槛不符": sum(1 for r in out if r.get("fold") and not r.get("expired")),
                    "拉取失败": sum(1 for r in out if r.get("error")),
                    "超出上限未筛": over},
        "ask_next": ask_next if prof.get("ask", True) and not empty else [],
        "api_calls": _CALLS["n"],
        "from_cache": _CALLS["cached"],
        "rule": ("expired 为 true 的已过报名截止：推送时直接不推；列表精筛时折叠成一行。"
                 + ("资质档案为空，只查了截止日期，不做资质比对、不追问。" if empty else MATCH_RULE)
                 + " 结果顺序与传入顺序一致，不重排。"),
        "results": out,
    }, ensure_ascii=False, indent=2))


# ── CLI ─────────────────────────────────────────────────────────────────
def build_parser():
    p = argparse.ArgumentParser(
        prog="radar.py",
        description="必赢 · 标讯雷达 查询引擎",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--version", action="version", version="bidwin-radar v" + VERSION)
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("search", help="搜标讯（主力命令）")
    s.add_argument("--keyword", help='主题词。空格=同时出现，竖线=或。例："AIGC|数字人"')
    s.add_argument("--include", help="结果必含词，竖线分隔（≤150 字符）")
    s.add_argument("--exclude", help="排除词，竖线分隔（≤150 字符）")
    s.add_argument("--days", type=int, help="近 N 天（含今天）。不给则默认 3 天")
    s.add_argument("--last", help="近 N 天/周/月/年，如 7d / 2w / 1m")
    s.add_argument("--start", help="起始日 YYYY-MM-DD（须与 --end 同时给）")
    s.add_argument("--end", help="结束日 YYYY-MM-DD")
    s.add_argument("--area", action="append", help="地名人话，可重复。例：--area 北京 --area 杭州")
    s.add_argument("--area-code", action="append", help="直传 6 位地区码（码表查不到时用）")
    s.add_argument("--classes", default=DEFAULT_CLASS_IDS,
                   help="招标阶段码，逗号分隔。默认 1,4,8,28,29（可投类）；-100=全部")
    s.add_argument("--purchase", default="全部", help="采购类型：服务 / 工程 / 货物 / 其它 / 全部（默认）")
    s.add_argument("--money-min", type=int, help="项目金额下限（元）")
    s.add_argument("--money-max", type=int, help="项目金额上限（元）")
    s.add_argument("--page", type=int, default=1, help="页码，默认 1")
    s.add_argument("--size", type=int, default=PAGE_MAX, help="每页条数，硬顶 50")
    s.add_argument("--probe", action="store_true", help="兼容旧用法：已并入正式搜索，照常返回结果和 total")
    s.add_argument("--allow-result", action="store_true", help="保留中标/合同等结果类（默认剔除）")
    s.add_argument("--limit", type=int, help="最多输出几条")
    s.add_argument("--out", help="写入 JSON 文件而不是打印全文")
    s.add_argument("--dump-raw", help=argparse.SUPPRESS)
    s.add_argument("--replay", help=argparse.SUPPRESS)   # 离线回放，开发自检用，不调接口
    s.set_defaults(func=cmd_search)

    d = sub.add_parser("detail", help="拉单条标详情")
    d.add_argument("--id", required=True, type=int)
    d.add_argument("--publish-time", help="可省略：脚本按 id 从最近的搜索结果里补")
    d.add_argument("--with-content", action="store_true", help="附带正文")
    d.add_argument("--max-chars", type=int, default=3000, help="正文最多输出多少字，默认 3000")
    d.add_argument("--brief", action="store_true", help="只取结构化字段，少一次请求（拿不到地区/行业/阶段）")
    d.set_defaults(func=cmd_detail)

    a = sub.add_parser("areas", help="查地区码（不调接口）")
    a.add_argument("query", nargs="?", help="地名关键字，留空列全部")
    a.set_defaults(func=cmd_areas)

    k = sub.add_parser("setkey", help="写入 key 到 ~/.bidwin/key（不调接口）")
    k.add_argument("--key", required=True)
    k.set_defaults(func=cmd_setkey)

    w = sub.add_parser("watch", help="每日推送的订阅条件（只存本机）")
    wsub = w.add_subparsers(dest="watch_cmd", required=True)
    ws = wsub.add_parser("set", help="保存 / 更新一组推送条件（不调接口）")
    ws.add_argument("--name", required=True, help="这组条件的名字，如「北京数字人」")
    ws.add_argument("--keyword", required=True)
    ws.add_argument("--include")
    ws.add_argument("--exclude")
    ws.add_argument("--area", action="append")
    ws.add_argument("--money-min", type=int)
    ws.add_argument("--money-max", type=int)
    ws.add_argument("--purchase", default="全部")
    ws.set_defaults(func=cmd_watch_set)
    wl = wsub.add_parser("list", help="查看已保存的推送条件（不调接口）")
    wl.set_defaults(func=cmd_watch_list)
    wd = wsub.add_parser("delete", help="删除一组推送条件（不调接口）")
    wd.add_argument("--name", required=True)
    wd.set_defaults(func=cmd_watch_delete)
    wr = wsub.add_parser("run", help="按已保存条件推送新标（定时任务调用）")
    wr.add_argument("--days", type=int, default=WATCH_DAYS)
    wr.add_argument("--dry-run", action="store_true", help="试跑：不记录为已推送")
    wr.add_argument("--replay", help=argparse.SUPPRESS)
    wr.set_defaults(func=cmd_watch_run)

    sc = sub.add_parser("screen", help="批量查报名截止 + 资质门槛比对（已过期的不拉正文，拉过的走缓存）")
    sc.add_argument("--ids", required=True, help='"id,id,id"（发布时间脚本自己补），一次最多 %d 条' % SCREEN_MAX)
    sc.set_defaults(func=cmd_screen)

    pf = sub.add_parser("profile", help="资质档案（只存本机，不调接口）")
    pfsub = pf.add_subparsers(dest="profile_cmd", required=True)
    pfs = pfsub.add_parser("set", help="写入 / 修改基本门槛与偏好")
    pfs.add_argument("--established", help="成立时间 YYYY / YYYY-MM / YYYY-MM-DD")
    pfs.add_argument("--capital-wan", type=float, help="注册资本（万元）")
    pfs.add_argument("--headcount", type=int, help="参保人数")
    pfs.add_argument("--reg-area", help="注册地，省或地级市，如「上海」「杭州」")
    pfs.add_argument("--ask", choices=["on", "off"], help="遇到缺项时是否追问")
    pfs.add_argument("--push-mode", choices=list(PUSH_MODES), help="fold=不符合的折叠（默认）/ all=全部展示")
    pfs.set_defaults(func=cmd_profile_set)
    for name, fn, hint in (("cert", cmd_profile_cert, "资质证书名，如「广播电视节目制作经营许可证」"),
                           ("case", cmd_profile_case, "业绩类型，如「汽车品牌TVC」")):
        x = pfsub.add_parser(name, help="记录一项%s有 / 没有" % ("资质证书" if name == "cert" else "业绩"))
        x.add_argument("--name", required=True, help=hint)
        x.add_argument("--has", choices=["yes", "no"])
        x.add_argument("--remove", action="store_true", help="从档案里删掉这一项")
        x.set_defaults(func=fn)
    pfv = pfsub.add_parser("show", help="查看资质档案")
    pfv.set_defaults(func=cmd_profile_show)

    c = sub.add_parser("check", help="环境自检（不调接口）")
    c.set_defaults(func=cmd_check)
    return p


def main():
    # Windows 控制台默认 GBK，中文和 ⚠️ 会报 UnicodeEncodeError → 强制 UTF-8
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    args = build_parser().parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
