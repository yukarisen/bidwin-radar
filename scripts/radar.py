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
    python3 radar.py detail --id 304974237 --publish-time "2026-08-01 09:00:00"
    python3 radar.py areas 杭州                               # 查地名（不调接口）
（Windows 上把 python3 换成 python 或 py）

每次接口调用都有成本，探量（--probe）也算一次，不要为试参数反复调用。
"""

import argparse
import html
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

VERSION = "1.0.0"
BASE_URL = os.environ.get("BIDWIN_SERVER_URL", "https://gate.gov-bid.com") + "/outer-gateway/bid"
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


_CALLS = {"n": 0}


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
        _CALLS["n"] += 1  # code=200 即计费，与返回内容无关
    else:
        raise ApiError("接口返回 code=%s subCode=%s msg=%s"
                       % (data.get("code"), data.get("subCode"), data.get("msg") or data.get("subMsg")))
    return data


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


def cmd_search(args):
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

    size = 0 if args.probe else min(args.size, PAGE_MAX)
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
        # 离线回放：拿之前 --dump-raw 存下的原始响应重跑筛选逻辑，不调接口、不扣费
        try:
            with open(args.replay, encoding="utf-8") as f:
                resp = json.load(f)
        except OSError as e:
            die("读不到回放文件 %s：%s" % (args.replay, e))
    else:
        try:
            resp = post("searchProjectApi", payload, key)
        except ApiError as e:
            die(str(e), code=3)
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
    }

    if args.probe:
        meta["hint"] = ("total ≤ 100 可直接正式搜；> 100 建议先帮用户缩窄"
                        "（限地区 / 缩时间 / 加排除词），不要直接甩一堆结果")
        emit(meta, args)
        return

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
    meta["rules_applied"] = [
        "日期由北京时间计算", "精准匹配",
        "HTML 标签已删除", "已按 标题+甲方+金额 去重",
        "截止日期一律「未注明」，禁用发布时间冒充",
    ]
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
    try:
        st = post("getZTBStructreDetail",
                  {"id": args.id, "publishTime": args.publish_time}, key).get("data") or {}
        # 一条标要取两份才够字段：结构化(截止/预算/联系人) + 正文(地区/行业/阶段/全文)
        ct = {} if args.brief else (post("getZTBProjectDetail",
              {"id": args.id, "publishTime": args.publish_time}, key).get("data") or {})
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
    out["reminder"] = "字段为空一律输出「未注明」，不得编造，尤其不得用发布时间当报名截止日期"
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
        "beijing_now": now_cn().strftime("%Y-%m-%d %H:%M:%S"),
        "note": "本命令不调接口",
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
    s.add_argument("--probe", action="store_true", help="只探量取 total，不返回数据")
    s.add_argument("--allow-result", action="store_true", help="保留中标/合同等结果类（默认剔除）")
    s.add_argument("--limit", type=int, help="最多输出几条")
    s.add_argument("--out", help="写入 JSON 文件而不是打印全文")
    s.add_argument("--dump-raw", help=argparse.SUPPRESS)
    s.add_argument("--replay", help=argparse.SUPPRESS)   # 离线回放，开发自检用，不调接口
    s.set_defaults(func=cmd_search)

    d = sub.add_parser("detail", help="拉单条标详情")
    d.add_argument("--id", required=True, type=int)
    d.add_argument("--publish-time", required=True, help="列表里的 publish_time 原值")
    d.add_argument("--with-content", action="store_true", help="附带正文")
    d.add_argument("--max-chars", type=int, default=3000, help="正文最多输出多少字，默认 3000")
    d.add_argument("--brief", action="store_true", help="只取结构化字段，省 1 次（拿不到地区/行业/阶段）")
    d.set_defaults(func=cmd_detail)

    a = sub.add_parser("areas", help="查地区码（不调接口）")
    a.add_argument("query", nargs="?", help="地名关键字，留空列全部")
    a.set_defaults(func=cmd_areas)

    k = sub.add_parser("setkey", help="写入 key 到 ~/.bidwin/key（不调接口）")
    k.add_argument("--key", required=True)
    k.set_defaults(func=cmd_setkey)

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
