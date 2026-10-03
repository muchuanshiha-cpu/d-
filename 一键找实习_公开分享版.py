# -*- coding: utf-8 -*-
"""
一键找实习 —— 实习岗位雷达 V2 单文件版（最终版）
读取    脚本所在文件夹/简历.docx
生成    脚本所在文件夹/适合你的岗位.docx
双击本文件，输入实习搜索起点（例如：北京市某区某地铁站），按 Enter 自动运行，
完成后自动打开 Word。本版已取消通勤时间计算，只计算岗位"距离搜索中心"的距离：
有高德 Key 用高德定位算真实直线距离；无 Key 按北京各区中心估算（明确标注）；无法定位写"未计算"。
"""

# ===================== 模块：config =====================

# -*- coding: utf-8 -*-
"""
实习岗位雷达 V2 - 配置文件
双击 run.bat 运行。所有选项都在本文件里改。
"""

# 默认搜索起点：仅在用户启动时未输入、且未传命令行参数时使用。
# 每次运行由使用者输入本次的实习搜索起点；留空表示不设置个人默认地址。
START_POINT = ""

# 最终输出岗位数上限（最多 66，找到多少输出多少，不凑数）。
TARGET_JOBS = 66

# 并发请求数（必须小，普通笔记本推荐 2）。
MAX_WORKERS = 2

# 单个网页请求超时（秒）。
TIMEOUT = 12

# 全局最小请求间隔（秒）。越小越快，越容易被限流；建议 1.0 以上。
REQUEST_INTERVAL = 1.0

# 单个网页体积上限（KB）。超过视为异常页，丢弃。
MAX_HTML_KB = 2048

# 搜索矩阵总查询数上限。越大召回越多，运行越久。默认 240。
MAX_QUERIES = 240

# 可选：高德开放平台 Web 服务 Key（免费申请）。
# 填写后程序用高德地理编码定位搜索中心与岗位地址，计算真实经纬度直线距离；
# 留空则按北京各区中心估算直线距离（Word 中明确标注“按区中心估算”），无法定位写“距离未计算”。
AMAP_KEY = ""

# ===================== 模块：crawler =====================

# -*- coding: utf-8 -*-
"""
py —— 公开页面抓取
只读取必要文本：单页上限 2MB、超时、限速、并发由 main 控制。
遇到验证码/登录/403/429/访问限制一律返回 None（跳过），不绕过。
"""
import time
import time
import threading
import requests
from bs4 import BeautifulSoup


UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")

_HEADERS = {
    "User-Agent": UA,
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

_lock = threading.Lock()
_last_ts = 0.0


def _throttle():
    """全局最小请求间隔（并发安全）"""
    global _last_ts
    with _lock:
        now = time.time()
        wait = REQUEST_INTERVAL - (now - _last_ts)
        if wait > 0:
            time.sleep(wait)
        _last_ts = time.time()


def fetch(url, timeout=None):
    """抓取一个公开页面。成功返回 HTML 文本；失败/被拦/超限返回 None。"""
    if not url or not url.startswith("http"):
        return None
    _throttle()
    timeout = timeout or TIMEOUT
    try:
        r = requests.get(url, headers=_HEADERS, timeout=timeout, allow_redirects=True)
    except Exception:
        return None
    if r.status_code != 200:
        return None
    if len(r.content) > MAX_HTML_KB * 1024:
        return None
    html = _decode(r)
    if _is_blocked(html):
        return None
    return html


def fetch_short(url, timeout=None):
    """抓取并转成纯文本（用于解析岗位内容），限制长度防止内存膨胀。"""
    html = fetch(url, timeout)
    if not html:
        return None
    return html_to_text(html)


def fetch_raw(url, timeout=None):
    """限速但不做反爬过滤的抓取，返回 (html, final_url)。
    供搜索引擎模块自行判断“验证码/无结果”类拦截页。失败返回 (None, None)。"""
    if not url or not url.startswith("http"):
        return (None, None)
    _throttle()
    timeout = timeout or TIMEOUT
    try:
        r = requests.get(url, headers=_HEADERS, timeout=timeout, allow_redirects=True)
    except Exception:
        return (None, None)
    if r.status_code != 200:
        return (None, None)
    if len(r.content) > MAX_HTML_KB * 1024:
        return (None, None)
    return (_decode(r), r.url)


def _decode(r):
    raw = r.content
    for enc in ("utf-8", "gbk", "gb18030"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return r.text


def _is_blocked(html):
    s = html[:2500].lower()
    for kw in ("antispider", "captcha", "wappass", "访问过于频繁",
               "验证码", "请稍候", "访问受限", "安全验证"):
        if kw in s:
            return True
    return False


def html_to_text(html, max_len=60000):
    """HTML -> 纯文本（去掉脚本/样式/导航噪音），最多保留 max_len 字符。"""
    try:
        soup = BeautifulSoup(html, "html.parser")
    except Exception:
        return ""
    for tag in soup(["script", "style", "noscript", "iframe", "svg", "canvas", "form", "nav", "footer", "header"]):
        tag.decompose()
    text = soup.get_text("\n", strip=True)
    return text[:max_len]

# ===================== 模块：location =====================

# -*- coding: utf-8 -*-
"""
location.py —— 地点标准化与区域关系（最终版）
只做文本层面的客观判断（同区/邻居区/北京市内/周边），不伪造距离。
"""
import re

BJ_DISTRICTS = ["东城", "西城", "朝阳", "海淀", "丰台", "石景山", "通州", "大兴",
                "顺义", "房山", "昌平", "门头沟", "平谷", "怀柔", "密云", "延庆",
                "亦庄", "经开区"]

SURROUND_KEYWORDS = ["燕郊", "廊坊", "固安", "香河", "三河", "大厂", "天津",
                     "保定", "张家口", "承德", "唐山"]

OTHER_CITY_RE = re.compile(r"(上海|广州|深圳|杭州|成都|南京|武汉|西安|苏州|郑州|"
                           r"青岛|济南|重庆|长沙|合肥|福州|厦门|大连|沈阳|长春|哈尔滨|"
                           r"昆明|贵阳|南宁|南昌|太原|石家庄|呼和浩特|兰州|西宁|银川|"
                           r"乌鲁木齐|拉萨|海口|三亚|无锡|宁波|温州|佛山|东莞|珠海)")


def extract_district(text):
    """从文本提取北京区名（含 亦庄/经开区）。"""
    if not text:
        return ""
    for d in BJ_DISTRICTS:
        if d in text:
            return d
    m = re.search(r"北京市?([\u4e00-\u9fa5]{2,6}(?:区|县))", text)
    if m:
        d = m.group(1)
        for name in BJ_DISTRICTS:
            if name in d:
                return name
    return ""


def classify(job_district, job_text, profile_prefs, start_point):
    """返回区域层级：same / preferred / bj / suburb / far / unknown
    只基于公开文本关系判断，与距离计算相互独立。"""
    text = (job_district or "") + " " + (job_text or "")[:2000]
    start_district = extract_district(start_point)

    if OTHER_CITY_RE.search(text):
        return "far"
    if any(k in text for k in SURROUND_KEYWORDS):
        return "suburb"

    district = extract_district(text)
    if not district:
        # 无区信息但明确是北京
        if "北京" in text:
            return "bj"
        return "unknown"

    if start_district and district == start_district:
        return "same"
    # 大兴区与亦庄/经开区视为相近
    if start_district in ("大兴", "亦庄", "经开区") and district in ("大兴", "亦庄", "经开区"):
        return "same"
    if district in (profile_prefs or []):
        return "preferred"
    if district in BJ_DISTRICTS:
        return "bj"
    return "unknown"


def region_rank(tier):
    """区域层级用于排序（越小越优先）。仅用于程序内排序。"""
    return {"same": 0, "preferred": 1, "bj": 2, "unknown": 3, "suburb": 4, "far": 5}.get(tier, 5)


def is_other_city(job_district, job_text):
    """是否明确为北京以外其他城市（含其城区名）。"""
    text = (job_district or "") + " " + (job_text or "")[:2000]
    if OTHER_CITY_RE.search(text):
        return True
    return False


def is_beijing_or_surround(job_district, job_text):
    text = (job_district or "") + " " + (job_text or "")[:2000]
    if OTHER_CITY_RE.search(text):
        return False
    if any(k in text for k in SURROUND_KEYWORDS):
        return True
    if "北京" in text:
        return True
    return False


# 北京各区相邻关系（用于搜索区域由近到远扩展）
BJ_NEIGHBORS = {
    "东城": ["西城", "朝阳", "丰台"], "西城": ["东城", "海淀", "丰台"],
    "朝阳": ["东城", "西城", "海淀", "通州", "大兴", "丰台"],
    "海淀": ["西城", "朝阳", "昌平", "石景山", "门头沟", "丰台"],
    "丰台": ["东城", "西城", "海淀", "朝阳", "大兴", "房山", "石景山"],
    "石景山": ["海淀", "丰台", "门头沟"],
    "通州": ["朝阳", "大兴", "顺义", "亦庄"],
    "大兴": ["丰台", "朝阳", "通州", "房山", "亦庄"],
    "顺义": ["朝阳", "通州", "昌平", "怀柔", "平谷", "密云"],
    "房山": ["丰台", "大兴", "门头沟"],
    "昌平": ["海淀", "顺义", "怀柔", "门头沟"],
    "门头沟": ["海淀", "石景山", "房山", "昌平"],
    "平谷": ["顺义", "密云"], "怀柔": ["顺义", "昌平", "密云"],
    "密云": ["顺义", "怀柔", "平谷"], "延庆": ["昌平", "怀柔"],
    "亦庄": ["大兴", "通州", "朝阳"], "经开区": ["大兴", "通州", "朝阳"],
}


def extract_street_word(text):
    """从用户输入的地址提取可作搜索词的街道/路段词。
    例：某区某路88号 → 提取道路或区域关键词；某城市某地铁站 → 提取地名关键词。"""
    if not text:
        return ""
    t = (text or "").strip()
    d = extract_district(t)
    if d:
        t = t.replace(d, "", 1).replace("区", "", 1)
    t = re.sub(r"北京市?|省|市", "", t)
    t = re.sub(r"[0-9０-９]+(?:号|栋|楼|室|单元)?", "", t)
    t = re.sub(r"(大街|大路|路|街|道|胡同|巷|大道|附近|号楼|大厦)", "", t)
    t = re.sub(r"\s+", "", t)
    if not t:
        return ""
    return t[:4]

# ===================== 模块：distance（最终版：取消通勤时间，只算距离） =====================

# -*- coding: utf-8 -*-
"""
distance —— 距离搜索中心计算（施工令第三版：取消通勤/公共交通时间）
- 有 AMAP_KEY：高德地理编码定位搜索中心与岗位地址 → 真实经纬度直线距离
- 无 Key：用北京各区中心经纬度估算直线距离，Word 明确标注“按区中心估算”
- 无法定位：距离未计算
- 禁止把直线距离写成道路距离，禁止凭感觉编公里数
"""
import time


# 北京各区中心经纬度（公开常识数据，仅用于无 Key 时粗粒度估算，非精确地址）
BJ_DISTRICT_CENTER = {
    "东城": (116.4165, 39.9283), "西城": (116.3659, 39.9124),
    "朝阳": (116.4432, 39.9215), "海淀": (116.2981, 39.9599),
    "丰台": (116.2870, 39.8584), "石景山": (116.2229, 39.9066),
    "通州": (116.6564, 39.9098), "大兴": (116.3415, 39.7268),
    "顺义": (116.6546, 40.1302), "房山": (116.1433, 39.7486),
    "昌平": (116.2312, 40.2207), "门头沟": (116.1017, 39.9406),
    "平谷": (117.1214, 40.1406), "怀柔": (116.6318, 40.3162),
    "密云": (116.8430, 40.3766), "延庆": (115.9750, 40.4566),
    "亦庄": (116.5085, 39.7909), "经开区": (116.5085, 39.7909),
}

_GEOCODE_CACHE = {}


def _haversine_km(a, b):
    """两经纬度直线距离（公里，保留1位小数）。"""
    from math import radians, sin, cos, asin, sqrt
    la1, lo1 = radians(a[1]), radians(a[0])
    la2, lo2 = radians(b[1]), radians(b[0])
    dlat, dlon = la2 - la1, lo2 - lo1
    h = sin(dlat / 2) ** 2 + cos(la1) * cos(la2) * sin(dlon / 2) ** 2
    return round(12742.0 * asin(sqrt(h)), 1)


class DistanceProvider:
    """计算岗位工作地点距搜索中心的直线距离。"""

    def __init__(self, amap_key=None):
        self.key = (amap_key or AMAP_KEY or "").strip()

    @property
    def available(self):
        return bool(self.key)

    def _geocode_amap(self, address):
        try:
            import requests
            from urllib.parse import quote
            url = ("https://restapi.amap.com/v3/geocode/geo?address=%s&city=%s&key=%s"
                   % (quote(address), quote("北京"), quote(self.key)))
            r = requests.get(url, timeout=TIMEOUT)
            data = r.json()
            if data.get("status") == "1" and data.get("geocodes"):
                loc = data["geocodes"][0].get("location", "")
                if loc and "," in loc:
                    lng, lat = loc.split(",")
                    return (float(lng), float(lat))
            return None
        except Exception:
            return None

    def locate(self, address, district=""):
        """地址→(经纬度, 是否高德精确)。高德优先，失败/无Key用区中心；都不行 (None, False)。"""
        if address:
            if address in _GEOCODE_CACHE:
                return _GEOCODE_CACHE[address]
            coord, exact = None, False
            if self.available:
                coord = self._geocode_amap(address)
                exact = coord is not None
            if coord is None:
                d = extract_district(address)
                coord = BJ_DISTRICT_CENTER.get(d)
            _GEOCODE_CACHE[address] = (coord, exact)
            return _GEOCODE_CACHE[address]
        if district:
            return (BJ_DISTRICT_CENTER.get(district), False)
        return (None, False)

    def estimate(self, center_text, job_address, job_district):
        """返回 (公里, 距离类型说明)；无法定位返回 None。"""
        center, c_exact = self.locate(center_text)
        job, j_exact = self.locate(job_address, job_district)
        if center is None or job is None:
            return None
        km = _haversine_km(center, job)
        if c_exact and j_exact:
            return km, "直线距离（高德定位）"
        return km, "直线距离（按区中心估算）"


ROUTE = DistanceProvider()

# ===================== 模块：profile =====================

# -*- coding: utf-8 -*-
"""
profile.py —— 简历画像：学校 / 学历 / 专业 / 技能 / 学生身份 / 求职方向 / 地点偏好
只根据简历原文与公开学校资料判断，不猜测。
"""

import re

# ---------------------------------------------------------------------------
# 学校知识库：学校名 -> (学校类型, 主要办学层次, 备注)
# 办学层次为公开资料中的主要层次；最终学历一律以本人真实学籍、毕业证为准。
# 后续可继续往这里增加学校。
# ---------------------------------------------------------------------------
SCHOOL_KNOWLEDGE = {
    # ---- 高职（专科）院校 ----
    "北京电子科技职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京信息职业技术学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京工业职业技术学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京农业职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京青年政治学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京政法职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京劳动保障职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京交通运输职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京卫生职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京经济管理职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京社会管理职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京体育职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京戏曲艺术职业学院": ("高等职业院校", "高职/专科", "公办高职"),
    "北京科技职业学院": ("高等职业院校", "高职/专科", "民办高职"),
    "北京培黎职业学院": ("高等职业院校", "高职/专科", "民办高职"),
    "北京汇佳职业学院": ("高等职业院校", "高职/专科", "民办高职"),
    "北京经贸职业学院": ("高等职业院校", "高职/专科", "民办高职"),
    "北京艺术传媒职业学院": ("高等职业院校", "高职/专科", "民办高职"),
    "北京网络职业学院": ("高等职业院校", "高职/专科", "民办高职"),
    # ---- 本科院校（含部分双一流/市属/民办本科）----
    "清华大学": ("普通本科", "本科", "双一流"),
    "北京大学": ("普通本科", "本科", "双一流"),
    "中国人民大学": ("普通本科", "本科", "双一流"),
    "北京师范大学": ("普通本科", "本科", "双一流"),
    "北京航空航天大学": ("普通本科", "本科", "双一流"),
    "北京理工大学": ("普通本科", "本科", "双一流"),
    "北京邮电大学": ("普通本科", "本科", "双一流"),
    "中央财经大学": ("普通本科", "本科", "双一流"),
    "对外经济贸易大学": ("普通本科", "本科", "双一流"),
    "中国政法大学": ("普通本科", "本科", "双一流"),
    "首都经济贸易大学": ("普通本科", "本科", "市属本科"),
    "北京工商大学": ("普通本科", "本科", "市属本科"),
    "北京联合大学": ("普通本科", "本科", "市属本科"),
    "北京信息科技大学": ("普通本科", "本科", "市属本科"),
    "北方工业大学": ("普通本科", "本科", "市属本科"),
    "北京物资学院": ("普通本科", "本科", "市属本科"),
    "北京农学院": ("普通本科", "本科", "市属本科"),
    "北京石油化工学院": ("普通本科", "本科", "市属本科"),
    "北京印刷学院": ("普通本科", "本科", "市属本科"),
    "北京服装学院": ("普通本科", "本科", "市属本科"),
    "北京建筑大学": ("普通本科", "本科", "市属本科"),
    "北京工业大学": ("普通本科", "本科", "双一流/市属"),
    "首都师范大学": ("普通本科", "本科", "双一流/市属"),
    "首都医科大学": ("普通本科", "本科", "市属本科"),
    "北京城市学院": ("普通本科", "本科", "民办本科"),
    "北京第二外国语学院": ("普通本科", "本科", "市属本科"),
    "北京语言大学": ("普通本科", "本科", "教育部直属"),
    "北京科技大学": ("普通本科", "本科", "双一流"),
    "北京化工大学": ("普通本科", "本科", "双一流"),
    "北京交通大学": ("普通本科", "本科", "双一流"),
    "北京林业大学": ("普通本科", "本科", "双一流"),
    "中国农业大学": ("普通本科", "本科", "双一流"),
    "华北电力大学": ("普通本科", "本科", "双一流"),
    "中国石油大学（北京）": ("普通本科", "本科", "双一流"),
    "中国地质大学（北京）": ("普通本科", "本科", "双一流"),
    "中国矿业大学（北京）": ("普通本科", "本科", "双一流"),
    "中央民族大学": ("普通本科", "本科", "双一流"),
}

# 简历中常见的学历/身份线索（用于兜底判断）
_DEGREE_WORDS = {
    "博士": "博士",
    "硕士": "硕士",
    "研究生": "硕士",
    "本科": "本科",
    "学士": "本科",
    "大专": "高职/专科",
    "专科": "高职/专科",
    "高职": "高职/专科",
    "中专": "中专/职高",
    "中职": "中专/职高",
}


# ---------------------------------------------------------------------------
# 专业扩展词（仅用于“发现岗位”，最终是否适合由匹配模块决定）
# ---------------------------------------------------------------------------
MAJOR_EXPANSION = {
    "大数据与会计": [
        "会计实习", "财务实习", "会计助理", "财务助理", "出纳实习", "审计实习",
        "税务实习", "财务数据", "财务分析", "数据分析", "经营分析", "商业分析",
        "财务共享", "ERP", "预算", "成本", "资金", "应收", "应付", "总账",
        "核算", "数据运营", "BI", "Excel", "SQL", "Python", "数据统计",
    ],
}

# 岗位关键词基础表（施工令第12条）
JOB_KEYWORDS = [
    "会计实习", "财务实习", "会计助理", "财务助理", "出纳实习", "审计实习",
    "税务实习", "财务分析实习", "财务数据实习", "数据分析实习", "数据运营实习",
    "商业分析实习", "BI实习", "ERP实习", "财务共享实习", "成本会计实习",
    "预算实习", "资金实习", "应收实习", "应付实习", "核算实习", "统计实习",
    "经营分析实习",
]

# 学生关键词（施工令第13条）
STUDENT_KEYWORDS = [
    "实习", "实习生", "在校生", "学生", "大学生", "高职", "大专", "专科",
    "应届", "可实习", "接受实习", "长期实习", "日常实习", "寒假实习",
    "暑期实习", "兼职实习", "实习岗位", "实习见习",
]

# 地区关键词（施工令第14条，按起点扩散）
REGION_KEYWORDS_BJ = [
    "大兴", "亦庄", "经开区", "通州", "丰台", "朝阳", "东城", "西城",
    "海淀", "顺义", "房山", "昌平", "石景山", "门头沟", "平谷", "怀柔",
    "密云", "延庆",
]

# 北京周边（R4 扩展候选）
REGION_KEYWORDS_SURROUND = ["燕郊", "廊坊", "固安", "香河", "三河", "天津"]


def analyze(resume_text):
    """根据简历原文生成画像。"""
    text = resume_text or ""
    # ---------- 学校 ----------
    school = ""
    for name in SCHOOL_KNOWLEDGE:
        if name in text:
            school = name
            break
    if not school:
        # 尝试识别"XX职业学院 / XX大学 / XX学院"形态
        m = re.search(r"([\u4e00-\u9fa5]{2,20}(?:职业学院|职业技术学院|大学|学院))", text)
        if m:
            school = m.group(1)

    school_type, school_level, school_note = _school_info(school)

    # ---------- 学历 ----------
    edu_level = school_level  # 默认以学校层次判断
    for kw, lv in _DEGREE_WORDS.items():
        if kw in text:
            edu_level = lv
            break
    if not edu_level:
        edu_level = "未能从公开来源确认"

    # ---------- 学籍/身份 ----------
    student_status = "未能确认"
    if re.search(r"在读|在校|预计\d{4}年毕业|大[一二三四五]", text):
        student_status = "在校生"
    elif re.search(r"应届|20\d{2}届", text):
        student_status = "应届毕业生"
    elif re.search(r"毕业生|已毕业", text):
        student_status = "毕业生"

    grad_year = ""
    m = re.search(r"(20\d{2})年?毕业", text)
    if m:
        grad_year = m.group(1)

    # ---------- 专业 ----------
    major = ""
    # 形式1：XXX专业（最常见，如 大数据与会计专业）
    m = re.search(r"([\u4e00-\u9fa5（）()·]{2,20})专业", text)
    if m:
        major = m.group(1)
    # 形式2：专业：XXX
    if not major:
        m = re.search(r"专业[:：]\s*([\u4e00-\u9fa5（）()·]{2,20})", text)
        if m:
            major = m.group(1)
    # 形式3：简历提到的核心专业词兜底
    if not major:
        m = re.search(r"(大数据与会计|会计学|会计|财务管理|审计|税务)", text)
        if m:
            major = m.group(1)

    # ---------- 技能 ----------
    skills = _extract_skills(text)

    # ---------- 证书 ----------
    certs = _extract_certs(text)

    # ---------- 求职方向 ----------
    direction = _extract_direction(text)

    # ---------- 地点偏好 ----------
    location_pref = _extract_location_pref(text)

    # ---------- 关键词扩展 ----------
    job_kws, student_kws = _build_keywords(major, direction, location_pref)

    return {
        "name": _extract_name(text),
        "school": school,
        "school_type": school_type,
        "school_level": school_level,
        "school_note": school_note,
        "edu_level": edu_level,
        "student_status": student_status,
        "graduation_year": grad_year,
        "major": major,
        "skills": skills,
        "certs": certs,
        "direction": direction,
        "location_pref": location_pref,
        "job_keywords": job_kws,
        "student_keywords": student_kws,
    }


def _school_info(school):
    if not school:
        return ("", "", "未能从公开来源确认")
    if school in SCHOOL_KNOWLEDGE:
        t, lv, note = SCHOOL_KNOWLEDGE[school]
        return (t, lv, note)
    if "职业" in school or "技术学院" in school:
        return ("高等职业院校", "高职/专科", "按校名判断，具体层次请以学籍为准")
    if "大学" in school or "学院" in school:
        return ("普通本科", "本科", "按校名判断，具体层次请以学籍为准")
    return ("", "", "未能从公开来源确认")


def _extract_skills(text):
    skills = []
    rules = [
        ("Excel", ["excel", "电子表格", "数据透视表", "函数"]),
        ("数据透视表", ["数据透视表"]),
        ("用友U8", ["用友", "u8", "ufida"]),
        ("财联", ["财联"]),
        ("财务软件", ["财务软件", "财务信息化", "会计信息化"]),
        ("SQL", ["sql"]),
        ("Python", ["python"]),
        ("脚本", ["脚本"]),
        ("AI工具", ["ai工具", "ai 工具"]),
        ("Windows系统维护", ["windows系统", "刷机", "系统安装"]),
    ]
    low = text.lower()
    for name, kws in rules:
        if any(k in low for k in kws):
            skills.append(name)
    return skills


def _extract_certs(text):
    certs = []
    for kw in ["初级会计", "会计从业", "初级会计师", "英语四级", "英语六级", "计算机二级"]:
        if kw.lower() in text.lower():
            certs.append(kw)
    if any(k in text for k in ["驾驶证", "C1", "C1D"]):
        certs.append("C1D驾驶证")
    return list(dict.fromkeys(certs))


def _extract_direction(text):
    # 优先取简历明确写的求职方向
    m = re.search(r"求职方向[:：\s]*([^\n]+)", text)
    if m:
        dirs = re.split(r"[、/，,]", m.group(1).strip())
        return [d.strip() for d in dirs if d.strip()]
    dirs = []
    for kw in ["财务助理", "会计助理", "出纳", "财务数据", "数据统计", "审计", "税务", "数据分析"]:
        if kw in text:
            dirs.append(kw)
    return dirs


def _extract_location_pref(text):
    prefs = []
    for kw in ["大兴", "亦庄", "经开区", "通州", "丰台", "朝阳", "海淀", "顺义", "房山", "昌平"]:
        if kw in text:
            prefs.append(kw)
    if "周边" in text:
        prefs.append("周边")
    return list(dict.fromkeys(prefs))


def _build_keywords(major, direction, location_pref):
    job_kws = list(JOB_KEYWORDS)
    # 专业扩展
    for base, extra in MAJOR_EXPANSION.items():
        if base in major:
            for kw in extra:
                if kw not in job_kws:
                    job_kws.append(kw)
    # 求职方向动态扩展
    for d in direction:
        for suffix in ("实习", "助理"):
            k = d + suffix
            if k not in job_kws:
                job_kws.append(k)
            if d + "实习" not in job_kws:
                job_kws.append(d + "实习")
    student_kws = list(STUDENT_KEYWORDS)
    return job_kws, student_kws


def _extract_name(text):
    first = text.splitlines()[0].strip() if text.splitlines() else ""
    if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", first):
        return first
    return ""

# ===================== 模块：resume_reader =====================

# -*- coding: utf-8 -*-
"""
py —— 只负责读取 Word 简历（正文 + 表格）
输出：dict { text, name, phone }
"""
import re
import docx


def read_resume(path):
    """读取 .docx 简历，返回原始文本与基本个人信息。
    只读取，不修改、不上传。"""
    doc = docx.Document(path)
    parts = []

    # 1) 正文段落
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            parts.append(t)

    # 2) 表格（每个单元格文本 + 行号，保证不丢失）
    for ti, table in enumerate(doc.tables):
        rows = []
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            line = " | ".join([c for c in cells if c])
            if line:
                rows.append(line)
        if rows:
            parts.append("【表格%d】" % (ti + 1))
            parts.extend(rows)

    text = "\n".join(parts)

    info = {
        "text": text,
        "name": _extract_name(text),
        "phone": _extract_phone(text),
    }
    return info


def _extract_name(text):
    # 第一行通常就是姓名（纯中文 2~4 字，无标点）
    first = text.splitlines()[0].strip() if text.splitlines() else ""
    if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", first):
        return first
    m = re.search(r"姓名[：:\s]*([\u4e00-\u9fa5]{2,4})", text)
    return m.group(1) if m else ""


def _extract_phone(text):
    m = re.search(r"1[3-9]\d{9}", text)
    return m.group(0) if m else ""

# ===================== 模块：job_parser =====================

# -*- coding: utf-8 -*-
"""
py —— 岗位字段提取（施工令第20条）
只从原始招聘页面/搜索摘要提取，缺失一律不编造。
信息等级：A=成功访问原始页面；B=页面不可访问但摘要可确认基本事实；C=模糊结果。
"""
import re
from datetime import datetime
from urllib.parse import urlparse


# ---------------------------------------------------------------------------
# 常用正则
# ---------------------------------------------------------------------------
_SALARY_K = re.compile(r"(\d+(?:\.\d+)?)\s*[-~至]\s*(\d+(?:\.\d+)?)\s*k(?:·\s*(\d+)\s*薪)?", re.I)
_SALARY_YUAN = re.compile(r"(\d+(?:\.\d+)?)\s*[-~至]\s*(\d+(?:\.\d+)?)\s*元\s*/\s*(天|月|时)")
_SALARY_SINGLE = re.compile(r"(\d+(?:\.\d+)?)\s*元\s*/\s*(天|月|时)")
_EDU = re.compile(r"(统招\s*)?(硕士|博士|本科|大专|专科|中专|学历不限|不限学历)")
_EXP = re.compile(r"((?:\d+\s*[-~至]\s*)?\d+\s*年\s*(?:以上|以下)?|经验不限|在校生|应届生|应届|无经验)")
_DISTRICT_FULL = re.compile(r"北京市?([\u4e00-\u9fa5]{2,6}(?:区|县))")
_DISTRICT_RAW = re.compile(r"北京-([\u4e00-\u9fa5]{1,6}(?:区|县)?)")
_ADDRESS = re.compile(r"北京市?[\u4e00-\u9fa5]{2,10}(?:区|县)[\u4e00-\u9fa5A-Za-z0-9·\-]{2,40}")
_PHONE_LAND = re.compile(r"(?<!\d)(?:010|0[1-9]\d{1,2})-?\d{7,8}(?!\d)")
_PHONE_MOBILE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
_PHONE_CTX = re.compile(r"电话|联系电话|座机|Tel|tel|联系|传真|联系人")
_HEADCOUNT = re.compile(r"[招聘]{1,2}(\d+)\s*人")
_DAYS_AGO = re.compile(r"(\d+)\s*天前")
_DATE_FULL = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
_WORKDAYS = re.compile(r"(?:每周|一周|每周至少)\s*(\d{1,2})\s*天")
_DURATION = re.compile(r"(\d{1,2}\s*个月|半年|一年|长期)(?:\s*以上)?")
_WELFARE = ["五险一金", "年终奖金", "带薪年假", "节日礼物", "免费班车", "餐补",
            "住房补贴", "股票期权", "弹性工作", "定期体检", "补充医疗保险", "绩效奖金"]
_COMPANY_SUFFIX = re.compile(r"(有限公司|有限责任公司|股份有限公司|集团|控股|股份公司|事务所|工作室|中心|研究院|学校|医院|政府|部队)$")

_PLATFORM_DOMAINS = {"liepin.com", "zhaopin.com", "51job.com", "zhipin.com",
                     "shixiseng.com", "yingjiesheng.com", "nowcoder.com",
                     "lagou.com", "ganji.com", "58.com", "mrcn.com"}

_LIST_PAGE_RE = re.compile(r"keyword=|/interns/?|/zhaopin/?|/search|/jobs/search|/post/index|/positions|/company/|joblist|liepin\.com/[a-z]*zp[a-z]+/?$", re.I)

# 各平台岗位详情链接形态（用于列表页发现）
JOB_URL_PATTERNS = [
    re.compile(r"https?://www\.liepin\.com/job/\d+\.shtml"),
    re.compile(r"https?://www\.zhipin\.com/job_detail/[^\"'\s]+"),
    re.compile(r"https?://www\.shixiseng\.com/intern/[^\"'\s]+"),
    re.compile(r"https?://www\.zhaopin\.com/jobs/[^\"'\s]+"),
    re.compile(r"https?://www\.nowcoder\.com/job/[^\"'\s]+"),
    re.compile(r"https?://www\.lagou\.com/jobs/\d+\.html"),
    re.compile(r"https?://www\.yingjiesheng\.com/job-\d{3}-\d{3}-\d+\.html"),
    re.compile(r"https?://jobs\.51job\.com/[^\"'\s]+"),
]


def make_job():
    return {
        "title": "", "company": "", "location_raw": "", "district": "",
        "address": "", "salary": "", "salary_unit": "", "edu_req": "",
        "exp_req": "", "accept_intern": None, "accept_student": None,
        "workdays": "", "duration": "", "duty": "", "requirement": "",
        "welfare": "", "headcount": "", "publish_date": "", "publish_days": None,
        "phone": "", "phone_source": "", "website": "", "url": "",
        "source_site": "", "source_type": "", "source_grade": "D",
        "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "distance_km": None, "distance_type": None,
        "region_tier": "unknown", "match_reasons": [], "score": 0.0,
        "excluded": False, "exclude_reason": "", "sources": [],
        "industry": "", "company_scale": "", "company_nature": "",
        "company_address": "", "company_website": "", "company_phone": "",
        "note": "",
    }


# ---------------------------------------------------------------------------
# 页面解析（A 级）
# ---------------------------------------------------------------------------
def parse_page(url, html, company_hint=""):
    """解析单岗位页面。不是岗位页返回 None。"""
    if not html:
        return None
    text = html_to_text(html, 60000)
    title_text = _title_from_html(html)
    return _parse_page_from_text(url, text, title_text, company_hint)


def parse_page_from_text(url, text, company_hint=""):
    """直接解析已转换好的纯文本（供抓取循环复用，省内存）。"""
    title_text = ""
    return _parse_page_from_text(url, text, title_text, company_hint)


def _parse_page_from_text(url, text, title_text, company_hint=""):
    if not text:
        return None
    text = text or ""
    if len(text) < 60:
        return None

    job = make_job()
    job["url"] = url
    job["source_site"] = urlparse(url).netloc
    job["fetched_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    job["title"] = _extract_title(title_text, text)
    job["company"] = _extract_company(title_text, text) or company_hint

    loc = _extract_location(text)
    job["location_raw"] = loc["raw"]
    job["district"] = loc["district"]
    job["address"] = loc["address"]

    job["salary"], job["salary_unit"] = _extract_salary(text)
    job["edu_req"] = _extract_edu(text)
    job["exp_req"] = _extract_exp(text)
    job["accept_intern"], job["accept_student"] = _extract_student_flags(text, job["title"])
    job["workdays"] = _extract_workdays(text)
    job["duration"] = _extract_duration(text)
    job["headcount"] = _extract_headcount(text)
    job["publish_date"], job["publish_days"] = _extract_publish(text)
    job["phone"], job["phone_source"] = _extract_phone(text)
    job["website"] = _extract_website(text, url)
    job["welfare"] = _extract_welfare(text)

    duty, req = _extract_duty_req(text)
    job["duty"] = duty
    job["requirement"] = req

    if _is_dead(text):
        job["excluded"] = True
        job["exclude_reason"] = "页面明确显示岗位已结束/已关闭/已截止"

    if not _is_job_page(job):
        return None
    return job


def _title_from_html(html):
    m = re.search(r"<title[^>]*>(.*?)</title>", html, re.S | re.I)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip()
    return ""


def _extract_title(title_text, text):
    head = title_text or text[:1200]
    # 形式1：猎聘 【北京 会计实习生招聘】-公司...（标题标签或正文开头）
    m = re.search(r"【([^】]+)】", head)
    if m:
        raw = m.group(1)
        raw = re.sub(r"^(全国|北京|上海|广州|深圳|杭州|成都|南京|武汉|天津|西安|苏州|郑州|青岛|济南|重庆|其他|北京上海|北京上海深圳|北京其它|北京上海广州)", "", raw)
        raw = re.sub(r"(招聘信息|招聘|信息|职位|岗位)$", "", raw)
        raw = raw.strip()
        if raw:
            return raw[:40]
    # 形式2：h1
    m = re.search(r"<h1[^>]*>(.*?)</h1>", title_text, re.S)
    if m:
        t = re.sub(r"<[^>]+>", "", m.group(1)).strip()
        if t:
            return t[:40]
    # 形式3：XXX公司招聘岗位名
    m = re.search(r"招聘([\u4e00-\u9fa5（）()·\s]{2,20})$", head)
    if m:
        t = m.group(1).strip()
        if t and len(t) <= 20:
            return t
    # 形式4：文本中「岗位：XXX」/「职位：XXX」
    m = re.search(r"(?:岗位|职位)[:：]\s*([\u4e00-\u9fa5（）()·]{2,20})", text[:3000])
    if m:
        return m.group(1)
    return ""


def _extract_company(title_text, text):
    head = title_text or text[:1500]
    # 猎聘详情页：已认证 ... · 公司名 聊一聊
    m = re.search(r"已认证[^聊]{0,80}?·\s*([^\s·]{2,40}?)\s*聊一聊", text)
    if m:
        return m.group(1).strip()
    # 标题：【..】-公司名北京招聘信息-猎聘
    m = re.search(r"】\s*-\s*([\u4e00-\u9fa5（）()]{2,40}?)(?:北京|上海|广州|深圳|全国|[\u4e00-\u9fa5]{2,10}招聘信息)?-?[^-\s]*$", head)
    if m:
        name = re.sub(r"(招聘信息|信息)$", "", m.group(1)).strip()
        if name and "猎聘" not in name and len(name) <= 40:
            return name
    # 应届生：[北京]公司名 / 岗位名
    m = re.search(r"\[[^\]]+\]([\u4e00-\u9fa5（）()]{2,40}?)\s*$", head)
    if m:
        return m.group(1).strip()
    return ""


def _extract_salary(text):
    head = text[:8000]
    m = _SALARY_K.search(head)
    if m:
        raw = "%s-%sk" % (m.group(1), m.group(2))
        if m.group(3):
            raw += "·%s薪" % m.group(3)
        return raw, "千/月"
    m = _SALARY_YUAN.search(head)
    if m:
        return "%s-%s元/%s" % (m.group(1), m.group(2), m.group(3)), "元/%s" % m.group(3)
    m = _SALARY_SINGLE.search(head)
    if m:
        return "%s元/%s" % (m.group(1), m.group(2)), "元/%s" % m.group(2)
    return "", ""


def _extract_edu(text):
    m = _EDU.search(text[:6000])
    if m:
        return m.group(0).replace(" ", "")
    return ""


def _extract_exp(text):
    m = _EXP.search(text[:6000])
    if m:
        e = m.group(1).replace(" ", "")
        # 排除“2026年”这类年份被当成经验年限
        if re.fullmatch(r"20\d{2}年", e):
            return ""
        return e
    return ""


def _extract_student_flags(text, title):
    accept_intern = None
    accept_student = None
    head = text[:6000]
    intern_hit = bool(re.search(r"实习|实习生|可实习|接受实习|实习生招聘", title + head))
    if intern_hit:
        accept_intern = True
    if re.search(r"学生可投|在校生|应届生|应届|大学生", head):
        accept_student = True
    if re.search(r"正式员工|社会招聘|劳动合同制", head) and not intern_hit:
        accept_intern = False
    return accept_intern, accept_student

def _extract_workdays(text):
    m = _WORKDAYS.search(text)
    if m:
        return "%s天/周" % m.group(1)
    return ""


def _extract_duration(text):
    m = _DURATION.search(text)
    if m:
        return m.group(1).replace(" ", "")
    return ""


def _extract_headcount(text):
    m = _HEADCOUNT.search(text[:6000])
    if m:
        return "%s人" % m.group(1)
    return ""


def _extract_publish(text):
    m = _DAYS_AGO.search(text[:8000])
    if m:
        return "%s天前更新" % m.group(1), int(m.group(1))
    m = _DATE_FULL.search(text)
    if m:
        try:
            d = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            days = (datetime.now() - d).days
            return "%s年%s月%s日" % (m.group(1), m.group(2), m.group(3)), max(0, days)
        except Exception:
            return "%s年%s月%s日" % (m.group(1), m.group(2), m.group(3)), None
    return "", None


def _extract_phone(text):
    # 固话：必须同时出现“电话/联系/座机/Tel/传真/联系人”等上下文，避免把编号当电话
    head = text[:8000]
    for m in _PHONE_LAND.finditer(head):
        i = m.start()
        if _PHONE_CTX.search(head[max(0, i - 25):i + 25]):
            return m.group(0), "招聘页面公开"
    m = _PHONE_MOBILE.search(head)
    if m:
        return m.group(0), "招聘页面公开"
    return "", ""


def _extract_website(text, page_url):
    host = urlparse(page_url).netloc
    for m in re.finditer(r"https?://([\w-]+\.)+(com|cn|net|org|gov)", text[:8000]):
        u = m.group(0)
        try:
            if urlparse(u).netloc == host:
                continue
        except Exception:
            continue
        dom = urlparse(u).netloc
        if any(dom.endswith(p) for p in _PLATFORM_DOMAINS):
            continue
        if dom in ("baidu.com", "weibo.com", "qq.com", "weixin.qq.com"):
            continue
        return u
    return ""


def _extract_welfare(text):
    found = [w for w in _WELFARE if w in text]
    return "、".join(found[:6])


def _extract_location(text):
    raw = ""
    district = ""
    address = ""
    m = _DISTRICT_RAW.search(text)
    if m:
        raw = "北京-" + m.group(1)
        d = m.group(1)
        district = _normalize_district(d)
    m = _DISTRICT_FULL.search(text)
    if m and not district:
        district = _normalize_district(m.group(1))
        raw = raw or "北京市%s" % m.group(1)
    # 摘要常见形态：北京 朝阳区（无连字符）
    if not district:
        for name in ["东城", "西城", "朝阳", "海淀", "丰台", "石景山", "通州", "大兴",
                     "顺义", "房山", "昌平", "门头沟", "平谷", "怀柔", "密云", "延庆"]:
            if re.search(r"北京\s*%s" % name, text[:6000]):
                district = name
                raw = raw or "北京-%s区" % name
                break
    if not district:
        m = _DISTRICT_FULL.search(text)
        if m:
            district = _normalize_district(m.group(1))
    m = _ADDRESS.search(text)
    if m:
        address = m.group(0)
    if "亦庄" in text[:6000] and not district:
        district = "亦庄"
        raw = raw or "北京-亦庄"
    if "经开区" in text[:6000] and not district:
        district = "经开区"
        raw = raw or "北京-经开区"
    return {"raw": raw, "district": district, "address": address}


def _normalize_district(d):
    d = re.sub(r"^(北京市?|北京市)", "", d)
    for name in ["东城", "西城", "朝阳", "海淀", "丰台", "石景山", "通州", "大兴",
                 "顺义", "房山", "昌平", "门头沟", "平谷", "怀柔", "密云", "延庆",
                 "亦庄", "经开区", "经济技术开发区"]:
        if name in d:
            return name
    return d


def _extract_duty_req(text):
    """从页面提取 岗位职责 / 任职要求（只清洗原文，不编写）。"""
    duty_lines, req_lines = [], []
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    mode = None
    section_kws = {
        "duty": ["岗位职责", "职位描述", "工作内容", "职位介绍", "岗位描述", "工作职责", "职责描述"],
        "req": ["任职要求", "岗位要求", "职位要求", "任职资格", "应聘要求", "基本要求"],
    }
    for i, line in enumerate(lines):
        low = line
        if any(k in low for k in section_kws["duty"]):
            mode = "duty"
            continue
        if any(k in low for k in section_kws["req"]):
            mode = "req"
            continue
        if mode == "duty" and any(k in low for k in section_kws["req"]):
            mode = "req"
            continue
        if mode is None:
            continue
        if len(line) < 6 or len(line) > 120:
            continue
        # 跳过明显的导航/页脚/联系方式行
        if re.match(r"^(首页|职位|投简历|聊一聊|登录|注册|下载|分享|收藏|上一篇|下一篇|相关推荐|公司介绍|工作地址|联系方式)", line):
            continue
        if mode == "duty" and len(duty_lines) < 8:
            duty_lines.append(line)
        elif mode == "req" and len(req_lines) < 8:
            req_lines.append(line)
    duty = "\n".join(duty_lines)
    req = "\n".join(req_lines)
    return duty, req


def _is_dead(text):
    head = text[:4000]
    for kw in ["已结束", "已关闭", "招聘截止", "已下线", "职位不存在", "该职位已过期", "停止招聘"]:
        if kw in head:
            return True
    return False


def _is_job_page(job):
    if not job["title"] or not job["company"]:
        return False
    if not any([job["salary"], job["location_raw"], job["edu_req"],
                job["exp_req"], job["duty"], job["requirement"]]):
        return False
    return True


# ---------------------------------------------------------------------------
# 搜索摘要解析（B 级）
# ---------------------------------------------------------------------------
def parse_snippet(item, grade="B"):
    """从搜索引擎结果摘要构造 B 级岗位记录（信息不足的字段留空）。"""
    title = item.get("title", "")
    snippet = item.get("snippet", "")
    url = item.get("url", "")
    # 列表页/搜索页/专题页不是单条岗位，信息不足不入库
    if url and _LIST_PAGE_RE.search(url):
        return None
    text = title + " " + snippet
    # 登录墙/纯导航标题：信息不可用，不入库
    if re.search(r"登录查看|请登录|登录后|查看高薪|免费注册|注册后查看", text[:200]):
        return None

    job = make_job()
    job["url"] = url
    job["source_site"] = urlparse(url).netloc if url else ""
    job["source_grade"] = grade

    m = re.search(r"【([^】]+)】", title)
    if m:
        raw = re.sub(r"^(北京|全国|上海|广州|深圳)", "", m.group(1))
        raw = re.sub(r"(招聘信息|招聘网站|招聘网|信息网|首页|平台|实时更新|网站|招聘|信息|职位|岗位)$", "", raw)
        job["title"] = raw.strip()[:40]
    if not job["title"]:
        t = re.sub(r"[【】\[\]·|｜\-_]", " ", title)
        t = re.sub(r"(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5])", "", t)  # 中文间空格
        t = re.sub(r"(招聘信息|招聘网站|招聘网|信息网|首页|平台|实时更新|网站|职位信息|求职招聘|找工作|招聘|信息|登录查看|查看高薪)", "", t)
        t = re.sub(r"\s+", " ", t).strip(" -")
        t = re.sub(r"^[【\[\]\s]*|[\s\]】]*$", "", t)
        job["title"] = t[:40]

    # 公司：通常出现在标题/摘要开头
    m = re.search(r"([\u4e00-\u9fa5（）()]{2,30}?(?:有限公司|有限责任公司|股份有限公司|集团|事务所|科技|服务|咨询|会计|财务))", text)
    if m:
        job["company"] = m.group(1)[:40]
    if not job["company"]:
        m = re.search(r"^([\u4e00-\u9fa5（）()]{2,20}?)[,，]", text)
        if m:
            job["company"] = m.group(1)
    # 摘要里出现平台名/城市+专业 之类伪公司名时清掉
    if job["company"] and (job["company"] in ("猎聘", "BOSS直聘", "前程无忧", "智联招聘", "实习僧", "应届生求职网", "牛客", "拉勾")
                           or re.fullmatch(r"(北京|上海|广州|深圳)[\u4e00-\u9fa5]{1,6}", job["company"])
                           or re.search(r"(实习|招聘|岗位|职位|就业|服务|了解|登录|平台)", job["company"])):
        job["company"] = ""

    loc = _extract_location(text)
    job["location_raw"] = loc["raw"]
    job["district"] = loc["district"]

    job["salary"], job["salary_unit"] = _extract_salary(text)
    job["edu_req"] = _extract_edu(text)
    job["exp_req"] = _extract_exp(text)
    job["accept_intern"], job["accept_student"] = _extract_student_flags(text, job["title"])
    job["publish_date"], job["publish_days"] = _extract_publish(text)
    job["note"] = snippet[:200]
    return job


def list_page_detail_urls(html, base_url, limit=3):
    """检测列表页并返回其中的岗位详情链接（有预算上限，避免无限爬取）。"""
    found = []
    host = urlparse(base_url).netloc
    for pat in JOB_URL_PATTERNS:
        for m in pat.finditer(html):
            u = m.group(0)
            if u not in found and urlparse(u).netloc:
                found.append(u)
            if len(found) >= limit:
                return found
    return found

# ===================== 模块：company_parser =====================

# -*- coding: utf-8 -*-
"""
py —— 公司信息补充（施工令第23条）
只使用岗位页面/企业公开信息；没有就写“未公开”，绝不推测。
"""
import re

_SCALE_RE = re.compile(r"(\d+[-~]\d+|10000以上|1000-\d+|500-\d+|50-\d+|1-\d+)\s*人")
_NATURE_RE = re.compile(r"(上市公司|国有企业|国企|央企|民营|私营|外资|合资|事业单位|政府机关|融资未公开|已上市|拟上市)")
_INDUSTRY_RE = re.compile(r"([\u4e00-\u9fa5·]{2,20}(?:服务|科技|互联网|金融|咨询|会计|审计|税务|商贸|制造|教育|医疗|建筑|物流|文化|传媒|软件|数据|信息|集团|公司))")


def parse_company(job, page_text, page_html=""):
    """从岗位页文本补充公司信息；job 会被原地更新。"""
    text = page_text or ""
    if job.get("company"):
        job["company"] = _clean_company_name(job["company"])

    m = _SCALE_RE.search(text)
    if m:
        job["company_scale"] = m.group(1) + "人"

    m = _NATURE_RE.search(text[:8000])
    if m:
        job["company_nature"] = m.group(1)

    # 行业：猎聘详情页常见 “某北京财务/审计/税务公司” / “财务/审计/税务” 标签
    m = re.search(r"([\u4e00-\u9fa5]{2,12}(?:[/、][\u4e00-\u9fa5]{2,8}){1,3}公司?)", text[:5000])
    if m and not re.search(r"(有限公司|集团)", m.group(1)):
        job["industry"] = m.group(1)[:30]
    if not job.get("industry"):
        m = _INDUSTRY_RE.search(text[:5000])
        if m:
            cand = m.group(1)
            if len(cand) <= 24 and not re.search(r"(招聘|登录|注册)", cand):
                job["industry"] = cand

    # 公司地址：页面公开的详细地址
    if job.get("address"):
        job["company_address"] = job["address"]
    else:
        m = re.search(r"北京市?[\u4e00-\u9fa5]{2,10}(?:区|县)[\u4e00-\u9fa5A-Za-z0-9·\-]{2,40}", text)
        if m:
            job["company_address"] = m.group(0)

    # 公司电话：优先复用岗位页公开电话
    if job.get("phone"):
        job["company_phone"] = job["phone"]

    # 官网：页面公开且非平台域名的链接
    if job.get("website"):
        job["company_website"] = job["website"]

    return job


def _clean_company_name(name):
    name = re.sub(r"^(北京|上海|广州|深圳)", "", name)
    name = re.sub(r"(北京招聘信息|招聘信息|信息)$", "", name)
    return name.strip()[:40]

# ===================== 模块：search_engine =====================

# -*- coding: utf-8 -*-
"""
py —— 搜索引擎发现（施工令第16条 SEARCH_PROVIDER）
当前可用入口：360 搜索（主）、必应 cn.bing.com（辅）。
百度/搜狗/DuckDuckGo 设计为可选项：被封/验证码时自动跳过，不影响整体。
360 结果链接为跳转页，用页面内公开的 JS 跳转地址解析出真实 URL（非绕过）。
"""
import re
import time
from bs4 import BeautifulSoup
from urllib.parse import quote_plus


REDIRECT_RE = re.compile(r'window\.location\.replace\("([^"]+)"\)')
NOSCRIPT_RE = re.compile(r"URL='([^']+)'")

# 搜索引擎拦截/验证页特征（用于检测后跳过，不绕过）
_BLOCK_URL_KW = ["qcaptcha", "noresult", "antispider", "captcha", "verify"]
_BLOCK_TEXT_KW = ["验证码", "访问过于频繁", "安全验证", "请输入验证码", "antispider"]


def resolve_360_redirect(url):
    """解析 so.com/link 跳转页，返回真实目标 URL；失败返回 None。"""
    html = fetch(url)
    if not html:
        return None
    m = REDIRECT_RE.search(html)
    if m:
        return m.group(1)
    m = NOSCRIPT_RE.search(html)
    if m:
        return m.group(1)
    return None


# ---------------------------------------------------------------------------
# 360 搜索
# ---------------------------------------------------------------------------
def search_360(query, page=1):
    """返回 [{title, url, snippet}]；被拦截返回字符串 'blocked'；失败返回 []。"""
    url = "https://www.so.com/s?q=%s&pn=%d" % (quote_plus(query), max(1, page))
    html, final_url = fetch_raw(url)
    if not html:
        return []
    if any(k in (final_url or "") for k in _BLOCK_URL_KW):
        return "blocked"
    if any(k in html[:3000] for k in _BLOCK_TEXT_KW):
        return "blocked"
    try:
        soup = BeautifulSoup(html, "html.parser")
    except Exception:
        return []
    out = []
    for res in soup.select(".res-list"):
        a = res.select_one("h3 a")
        if not a:
            continue
        title = a.get_text(" ", strip=True)
        href = a.get("href", "")
        if not href.startswith("http"):
            continue
        snippet = ""
        for sel in (".res-desc", ".res-comm-con", "p"):
            p = res.select_one(sel)
            if p:
                snippet = p.get_text(" ", strip=True)
                if snippet:
                    break
        out.append({"title": title, "url": href, "snippet": snippet, "engine": "360"})
    return out


# ---------------------------------------------------------------------------
# 必应搜索
# ---------------------------------------------------------------------------
def search_bing(query, page=1):
    """返回 [{title, url, snippet}]；失败返回 []。"""
    url = ("https://cn.bing.com/search?q=%s&count=20&first=%d&mkt=zh-CN"
           % (quote_plus(query), (max(1, page) - 1) * 10))
    html = fetch(url)
    if not html:
        return []
    try:
        soup = BeautifulSoup(html, "html.parser")
    except Exception:
        return []
    out = []
    for item in soup.select("li.b_algo"):
        a = item.select_one("h2 a")
        if not a:
            continue
        title = a.get_text(" ", strip=True)
        href = a.get("href", "")
        if not href.startswith("http"):
            continue
        snippet = ""
        p = item.select_one(".b_caption p")
        if p:
            snippet = p.get_text(" ", strip=True)
        out.append({"title": title, "url": href, "snippet": snippet, "engine": "bing"})
    return out


# ---------------------------------------------------------------------------
# 提供者注册表（被拦截即冷却，自动跳过，不影响整体）
# ---------------------------------------------------------------------------
class ProviderHub:
    def __init__(self):
        self.providers = [
            {"name": "360", "fn": search_360, "cooldown_until": 0.0},
            {"name": "bing", "fn": search_bing, "cooldown_until": 0.0},
        ]

    def _pick(self, query, forced=None):
        now = time.time()
        if forced:
            for p in self.providers:
                if p["name"] == forced and now >= p["cooldown_until"]:
                    return p
            return None
        # 政府/高校/企业官网类查询优先必应（360 对 site: 域名的覆盖有限）
        if re.search(r"site:|\b(gov|edu|24365|就业见习|毕业生就业)\b", query):
            for p in self.providers:
                if p["name"] == "bing" and now >= p["cooldown_until"]:
                    return p
        for p in self.providers:
            if now >= p["cooldown_until"]:
                return p
        return None

    def search(self, query, page=1, forced=None):
        p = self._pick(query, forced)
        if p is None:
            return []
        try:
            res = p["fn"](query, page)
        except Exception:
            res = []
        if res == "blocked":
            p["cooldown_until"] = time.time() + 180  # 冷却 3 分钟，期间不打扰该引擎
            other = self._pick(query)
            if other and other["name"] != p["name"]:
                try:
                    res = other["fn"](query, page)
                except Exception:
                    res = []
                if res == "blocked":
                    other["cooldown_until"] = time.time() + 180
                    return []
            else:
                return []
        return res if isinstance(res, list) else []


PROVIDER_HUB = ProviderHub()


def search_query(query, page=1, forced=None):
    """对外统一入口。"""
    return PROVIDER_HUB.search(query, page, forced)

# ===================== 模块：source_search =====================

# -*- coding: utf-8 -*-
"""
py —— 公开招聘来源发现（施工令第15条 三层来源）
第一层：政府/公共就业平台（大多需登录/JS/被限流，交由搜索引擎发现，直接探测失败的跳过并记录）
第二层：大型招聘平台公开招聘页（猎聘招聘专场页可完整公开抓取 = A级主源；BOSS/智联/51job/实习僧等为 JS 壳，由搜索引擎摘要提供 B 级）
第三层：企业官网/高校就业网（由搜索引擎 site: 查询发现）
本模块只负责“发现候选 URL”，字段解析交给 job_parser / company_parser。
"""
import re
from urllib.parse import urljoin


# 已验证可公开抓取的猎聘招聘专场 slug（北京）。运行时还会通过搜索引擎增量发现。
LIEPIN_SLUGS_KNOWN = {
    "会计实习": "zphuijishixi",
    "会计实习生": "zphuijishixisheng",
    "财务实习": "zpcaiwushixi",
    "财务实习生": "zpcaiwushixisheng",
    "财务助理": "zpcaiwuzhuli",
    "审计实习生": "zpshenjishixisheng",
    "出纳实习生": "zpchunashixisheng",
    "预算": "zpyusuan",
    "数据分析实习生": "zpshujufenxishixisheng",
    "数据统计": "zpshujutongji",
    "大数据": "zpdashuju",
    "数据分析": "zpshujufenxi",
    "经营分析": "zpjingyingfenxi",
}

LIEPIN_SPECIAL_RE = re.compile(r"/(?:city-bj/)?(zp[a-z0-9]+)/", re.I)
LIEPIN_JOB_RE = re.compile(r"https?://www\.liepin\.com/job/\d+\.shtml")
YJS_JOB_RE = re.compile(r"/job-\d{3}-\d{3}-\d+\.html")

# 政府/公共就业平台（直接探测用；失败即跳过，不影响整体）
GOV_SOURCES = [
    ("中国公共招聘网", "https://job.mohrss.gov.cn/"),
    ("就业在线", "https://www.jobonline.cn/"),
    ("国家人才网", "https://www.newjobs.com.cn/"),
    ("24365国家大学生就业服务平台", "https://job.ncss.cn/"),
    ("北京高校大学生就业创业信息网", "https://jobs.bjbys.net.cn/frontpage/bjbys/html/recruitmentinfoList.html"),
]


def discover_liepin_slugs(job_keywords, limit=16):
    """通过搜索引擎发现更多猎聘北京专场页 slug，合并已知 slug。"""
    slugs = list(dict.fromkeys(LIEPIN_SLUGS_KNOWN.values()))
    queries = []
    for kw in job_keywords[:10]:
        queries.append("site:liepin.com %s" % kw)
    for q in queries:
        if len(slugs) >= limit:
            break
        results = search_query(q)
        for item in results:
            title = item.get("title", "")
            url = item.get("url", "")
            if "liepin.com" not in url and "猎聘" not in title:
                continue
            real = resolve_360_redirect(url) if url.startswith("https://www.so.com/link") else url
            if real:
                m = LIEPIN_SPECIAL_RE.search(real)
                if m and m.group(1) not in slugs:
                    slugs.append(m.group(1))
            if len(slugs) >= limit:
                break
    return slugs


def liepin_special_job_urls(slugs):
    """抓取猎聘专场页，提取里面的岗位详情链接（A 级候选）。"""
    seen = set()
    for slug in slugs:
        url = "https://www.liepin.com/city-bj/%s/" % slug
        html = fetch(url)
        if not html:
            continue
        for m in LIEPIN_JOB_RE.finditer(html):
            u = m.group(0)
            if u not in seen:
                seen.add(u)
        if len(seen) >= 400:  # 预算内
            break
    return sorted(seen)


def yingjiesheng_job_urls(limit=120):
    """应届生求职网北京站：岗位详情页有阿里云WAF验证（访问限制），
    按规则跳过直接抓取；该来源改由搜索引擎摘要（B级）进入。
    保留本函数仅为读取列表页做统计参考。"""
    return []


def probe_gov_sources():
    """政府平台直接探测：能抓且含岗位信息才算数，否则记跳过原因返回空。"""
    found = []
    for name, url in GOV_SOURCES:
        html = fetch(url)
        if not html:
            continue
        text = html_to_text(html, 4000)
        if len(text) < 100:
            continue
        if re.search(r"岗位|职位|招聘", text):
            found.append((name, url, text[:2000]))
    return found

# ===================== 模块：matcher =====================

# -*- coding: utf-8 -*-
"""
py —— 岗位适合度匹配（施工令第25、26、42、43条）
输出客观“匹配说明”，不用“很好/值得去”之类主观评价。
分数仅用于程序内排序。
"""
import re

# 销售类岗位：原则上排除
SALES_TITLE_KW = ["销售", "电销", "业务员", "地推", "置业顾问", "保险",
                  "信用卡", "主播", "带货", "电话客服", "催收", "推销"]
SALES_DESC_KW = ["电话销售", "完成销售业绩", "销售指标", "销售任务", "销售顾问",
                 "陌拜", "电销", "客户资源", "销售提成", "业绩要求"]

# 明确要求 2 年以上经验：排除（施工令第42条）
HARD_EXP_RE = re.compile(r"([2-9]|[1-9]\d|[两三四五六七八九十])\s*年\s*以上")
SOFT_EXP_RE = re.compile(r"(1\s*年\s*以上|1\s*[-~至]\s*3\s*年)")


def _range_hard_exp(exp_text):
    """“X-Y年/X至Y年”范围经验：较大值>=2 且非“0-1年” → 硬排除。
    例：5-10年 → 排除；0-1年 → 保留；1-2年 → 排除。"""
    for a, b in re.findall(r"(\d{1,2})\s*[-~至]\s*(\d{1,2})\s*年", exp_text or ""):
        if int(b) >= 2 and not (int(a) == 0 and int(b) <= 1):
            return True
    return False

# 学历硬门槛
EDU_EXCLUDE_RE = re.compile(r"硕士|博士")

# 只招正式员工
FULLTIME_ONLY_RE = re.compile(r"仅限正式|只招正式|正式员工优先|只招全职|仅招全职")

# 专业层级
LEVEL1 = ["会计", "财务", "出纳", "审计", "税务", "核算", "成本会计", "预算",
          "资金", "应收", "应付", "总账", "报表", "对账", "财务分析", "会计助理",
          "财务助理", "出纳", "记账", "凭证"]
LEVEL2 = ["数据", "统计", "BI", "ERP", "经营分析", "商业分析", "数据分析",
          "数据运营", "财务数据", "excel"]
LEVEL3 = ["行政", "文员", "人事", "前台", "助理", "运营", "专员", "客服"]

# 学生资格关键词
FREE_EXP_KW = ["经验不限", "在校生", "应届", "无经验", "学生可投"]


def match_job(job, profile):
    """返回 job 的匹配结果（原地写入 excluded/score/match_reasons）。"""
    title = job.get("title", "")
    text = "%s %s %s" % (title, job.get("duty", ""), job.get("requirement", ""))
    text_low = text.lower()
    reasons = []

    # ---------------- 排除规则 ----------------
    if job.get("excluded"):
        return job

    if any(k in title for k in SALES_TITLE_KW):
        job["excluded"] = True
        job["exclude_reason"] = "岗位属于销售类（原则上排除）"
        return job
    sales_hits = [k for k in SALES_DESC_KW if k in text]
    if len(sales_hits) >= 2:
        job["excluded"] = True
        job["exclude_reason"] = "职责内容属于销售类（原则上排除）"
        return job

    exp_req = job.get("exp_req", "")
    if HARD_EXP_RE.search(exp_req) or _range_hard_exp(exp_req):
        job["excluded"] = True
        job["exclude_reason"] = "明确要求2年以上经验，与学生条件不符"
        return job
    # 任职要求正文兜底（逐句判断；句内含“优先/加分/尤佳”视为偏好而非硬要求）
    req_text = job.get("requirement", "")
    hard_sent = False
    for sent in re.split(r"[;；。\n]", req_text[:6000]):
        if HARD_EXP_RE.search(sent) or _range_hard_exp(sent):
            if not re.search(r"优先|加分|尤佳", sent):
                hard_sent = True
                break
    if hard_sent:
        job["excluded"] = True
        job["exclude_reason"] = "任职要求明确多年经验，与学生条件不符"
        return job

    edu_req = job.get("edu_req", "")
    if EDU_EXCLUDE_RE.search(edu_req):
        job["excluded"] = True
        job["exclude_reason"] = "学历门槛为硕士/博士，不符合"
        return job

    if FULLTIME_ONLY_RE.search(text[:2000]):
        job["excluded"] = True
        job["exclude_reason"] = "岗位明确只招正式员工"
        return job

    if job.get("publish_days") is not None and job["publish_days"] > 90:
        job["excluded"] = True
        job["exclude_reason"] = "发布时间超过90天（原则上不进入主列表）"
        return job

    # 地点：明确为其他城市 → 排除；未确认地点 → 保留并标记（进扩展候选）
    if is_other_city(job.get("district", ""),
                     (job.get("location_raw") or "") + " " + (job.get("address") or "") + " " + (job.get("note") or "")):
        job["excluded"] = True
        job["exclude_reason"] = "工作地点为其他城市，不在北京及合理周边"
        return job

    # ---------------- 学生资格 30 ----------------
    stu = 0
    if job.get("accept_intern") is True:
        stu = 26
        reasons.append("学生资格：明确接受实习")
    elif job.get("accept_intern") is False:
        stu = 6
        reasons.append("学生资格：页面偏向正式员工（谨慎）")
    else:
        stu = 18
        reasons.append("学生资格：未明确是否接受实习")
    if job.get("accept_student") is True:
        stu = min(30, stu + 4)
        reasons.append("学生资格：在校生/应届可投")
    if any(k in exp_req for k in FREE_EXP_KW):
        stu = min(30, stu + 4)
        reasons.append("学生资格：经验不限/接受学生")
    if SOFT_EXP_RE.search(exp_req):
        stu = max(0, stu - 14)
        reasons.append("学生资格：要求1年以上经验（与学生条件可能有出入）")
    if edu_req and ("本科" in edu_req):
        stu = max(0, stu - 4)
        reasons.append("学历：岗位要求本科，学生为高职/专科（可能不符）")
    elif edu_req and ("大专" in edu_req or "专科" in edu_req or "学历不限" in edu_req):
        reasons.append("学历：符合（岗位要求大专/专科/学历不限）")
    elif not edu_req:
        reasons.append("学历：岗位未注明要求")
    else:
        reasons.append("学历：岗位要求%s（以实际招聘为准）" % edu_req)

    # ---------------- 专业匹配 25 ----------------
    pro = 0
    if any(k in title for k in LEVEL1):
        pro = 25
        reasons.append("专业匹配：岗位名称与会计/财务方向高度相关")
    elif any(k in title for k in LEVEL2):
        pro = 22
        reasons.append("专业匹配：岗位名称与数据/财务数据方向相关")
    elif any(k in title for k in LEVEL3):
        pro = 8
        reasons.append("专业匹配：岗位属于通用型岗位（非财务专业核心岗）")
    elif any(k in text for k in LEVEL1):
        pro = 20
        reasons.append("专业匹配：职责内容含会计/财务要素")
    elif any(k in text for k in LEVEL2):
        pro = 17
        reasons.append("专业匹配：职责内容含数据/统计要素")
    else:
        pro = 4
        reasons.append("专业匹配：未发现会计/财务/数据要素")

    # ---------------- 技能匹配 15 ----------------
    skill_points = {"Excel": 6, "数据透视表": 3, "用友U8": 3, "财联": 2,
                    "财务软件": 2, "SQL": 3, "Python": 3, "脚本": 1,
                    "AI工具": 1, "Windows系统维护": 1}
    sk = 0
    matched = []
    for skill in profile.get("skills", []):
        pts = skill_points.get(skill, 1)
        if skill.lower() in text_low:
            sk = min(15, sk + pts)
            matched.append(skill)
    if matched:
        reasons.append("技能匹配：%s" % "、".join(matched))
    else:
        reasons.append("技能匹配：页面未明确列出简历中技能")

    # ---------------- 距离 8（辅助排序，不压过岗位匹配） ----------------
    cm = 0
    dkm = job.get("distance_km")
    if dkm is not None:
        if dkm <= 5:
            cm = 8
        elif dkm <= 10:
            cm = 7
        elif dkm <= 20:
            cm = 5
        elif dkm <= 30:
            cm = 3
        else:
            cm = 1
        if dkm == 0.0 and "区中心" in (job.get("distance_type") or ""):
            reasons.append("距离：同区（区中心估算，未到街道级）")
        else:
            reasons.append("距离搜索中心：%s公里（%s）" % (dkm, job.get("distance_type") or "直线距离"))
    else:
        cm = 0
        tier = job.get("region_tier", "unknown")
        if tier in ("same", "preferred"):
            reasons.append("距离：未计算（地点与起点同区/偏好区）")
        elif tier == "bj":
            reasons.append("距离：未计算（地点在北京市内）")
        else:
            reasons.append("距离：未计算")

    # ---------------- 信息完整度 5 ----------------
    fields = [job.get("salary"), job.get("edu_req"), job.get("exp_req"),
              job.get("duty"), job.get("location_raw"), job.get("publish_date"),
              job.get("phone")]
    present = sum(1 for f in fields if f)
    info_score = round(5.0 * present / len(fields), 1)

    # ---------------- 来源可靠性 5 ----------------
    grade = job.get("source_grade", "D")
    src_score = {"A": 5, "B": 4, "C": 2, "D": 1}.get(grade, 1)

    score = round(stu + pro + sk + cm + info_score + src_score, 1)
    job["score"] = score
    job["match_reasons"] = reasons[:8]
    return job

# ===================== 模块：dedup =====================

# -*- coding: utf-8 -*-
"""
py —— 岗位去重（施工令第24条）
同一公司同一岗位在不同平台出现：只保留一个主记录，来源合并显示。
"""
import re


def norm_company(name):
    n = (name or "").strip().lower()
    n = re.sub(r"(有限公司|有限责任公司|股份有限公司|股份公司|集团|控股|责任公司)$", "", n)
    n = re.sub(r"[\s·|｜\-_]+", "", n)
    return n


def norm_title(title):
    t = (title or "").strip().lower()
    t = re.sub(r"[\s!！?？。，,·\-_【】\[\]（）()]+", "", t)
    t = re.sub(r"^(招聘|诚聘|急聘|高薪|北京|全国)", "", t)
    return t


def dedup_jobs(jobs):
    """按 URL 去重，再按 公司+岗位 合并多平台记录。返回去重后的列表。"""
    by_url = {}
    for j in jobs:
        if j.get("url") and j["url"] not in by_url:
            by_url[j["url"]] = j

    merged = {}
    order = []
    for j in by_url.values():
        key = (norm_company(j.get("company", "")), norm_title(j.get("title", "")))
        if not key[0] or not key[1]:
            merged[j["url"]] = j
            order.append(j["url"])
            continue
        if key in merged:
            old = merged[key]
            _merge(old, j)
        else:
            merged[key] = j
            order.append(key)

    return [merged[k] for k in order if k in merged]


def _merge(main_job, other):
    """把 other 的来源/链接并入 main_job（保留信息更完整的记录）。"""
    if other.get("source_site") and other["source_site"] not in main_job.get("sources", []):
        main_job.setdefault("sources", []).append(other["source_site"])
    if other.get("url") and other["url"] != main_job.get("url") and other["url"] not in main_job.get("sources", []):
        main_job.setdefault("sources", []).append(other["url"])
    # 等级取更优
    grade_order = {"A": 3, "B": 2, "C": 1, "D": 0}
    if grade_order.get(other.get("source_grade", "D"), 0) > grade_order.get(main_job.get("source_grade", "D"), 0):
        main_job["source_grade"] = other["source_grade"]
    # 字段补缺
    for field in ("salary", "salary_unit", "edu_req", "exp_req", "duty",
                  "requirement", "location_raw", "address", "phone",
                  "publish_date", "headcount", "workdays", "duration", "welfare"):
        if not main_job.get(field) and other.get(field):
            main_job[field] = other[field]

# ===================== 模块：ranking =====================

# -*- coding: utf-8 -*-
"""
py —— 排序与主/扩展区分（最终版）
排序依据：距离近 > 综合得分（学生资格/专业/技能/信息完整度/来源）> 区域关系 > 来源等级。
距离未计算且地点未确认/周边区域的岗位进“扩展候选区”，如实标注。
核心原则：离搜索中心越近，优先级越高。
"""


def region_order(job):
    return region_rank(job.get("region_tier", "unknown"))


def grade_order(job):
    return {"A": 3, "B": 2, "C": 1, "D": 0}.get(job.get("source_grade", "D"), 0)


def info_count(job):
    return sum(1 for f in (job.get("salary"), job.get("duty"),
                           job.get("location_raw"), job.get("publish_date")) if f)


def is_main_candidate(job):
    """主推荐：距离已计算，或区域关系合理（同区/偏好区/北京市内）。"""
    if job.get("distance_km") is not None:
        return True
    return job.get("region_tier") in ("same", "preferred", "bj")


def split_and_rank(jobs, profile):
    """jobs: 已匹配未排除的岗位列表。
    返回 (main_list, expansion_list)，均按 距离→得分 排序。"""
    main_list, expansion_list = [], []
    for j in jobs:
        if is_main_candidate(j):
            main_list.append(j)
        else:
            expansion_list.append(j)

    def _key(j):
        dkm = j.get("distance_km")
        return (-j.get("score", 0),
                -grade_order(j),
                -info_count(j),
                -region_order(j),
                dkm if dkm is not None else 99999)

    main_list.sort(key=_key)
    expansion_list.sort(key=_key)
    return main_list, expansion_list


def finalize(main_list, expansion_list, target):
    """输出最终列表：主区优先，不足时从扩展区补充（过滤信息过少的噪声岗位），补充项明确标注。"""
    main = main_list[:target]
    if len(main) < target:
        noise_re = re.compile(r"招聘网|招聘信息|求职招聘|找工作|登录查看|国家大学生|财政部|人员统一|公众服务|年度初级|履行助理|(^|\s)未公开(\s|$)|^了解")
        def _clean(j):
            return (j.get("title") or "") + " " + (j.get("company") or "")
        pick = [j for j in expansion_list if (j.get("company") or "").strip() and not noise_re.search(_clean(j))]
        need = target - len(main)
        for j in pick[:need]:
            j["note"] = (j.get("note") or "") + "；扩展候选：距离未计算或地点未确认"
            main.append(j)
    return main

# ===================== 模块：word_export =====================

# -*- coding: utf-8 -*-
"""
py —— 生成《适合你的岗位.docx》（施工令第30条）
结构：基本信息 / 搜索结果概况 / 推荐岗位 / 扩展候选岗位 / 未完全确认信息
只输出真实抓取到的内容，缺失字段一律写“未公开/未能从公开来源确认”。
"""
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

SEP = "----------------------------------"


def _set_font(run, size=10.5, bold=False, color=None, name="微软雅黑"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name)
    if color:
        run.font.color.rgb = RGBColor(*color)


def _para(doc, text, size=10.5, bold=False, color=None, align=None, space_after=4):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    _set_font(p.add_run(text), size, bold, color)
    return p


def _job_block(doc, idx, job):
    """输出单个岗位条目"""
    _para(doc, SEP, size=9, color=(120, 120, 120))
    company = job.get("company") or "未公开"
    _para(doc, "%02d｜%s" % (idx, company), size=12, bold=True)
    _para(doc, "岗位：%s" % (job.get("title") or "未公开"))

    reasons = job.get("match_reasons") or []
    if reasons:
        _para(doc, "匹配说明：" + "；".join(reasons), size=10)

    loc = job.get("location_raw") or job.get("address") or "未公开"
    _para(doc, "工作地点：%s" % loc)

    dkm = job.get("distance_km")
    if dkm is not None:
        if dkm == 0.0 and "区中心" in (job.get("distance_type") or ""):
            _para(doc, "距离搜索中心：同区（区中心估算，未到街道级）")
        else:
            _para(doc, "距离搜索中心：%s 公里" % dkm)
        _para(doc, "距离类型：%s" % (job.get("distance_type") or "直线距离"))
    else:
        _para(doc, "距离搜索中心：未计算")
        _para(doc, "距离类型：未计算")

    if job.get("salary"):        sal = job["salary"]
        unit = job.get("salary_unit") or ""
        if unit and unit not in sal:
            sal = "%s（%s）" % (sal, unit)
        _para(doc, "薪资：%s" % sal)
    else:
        _para(doc, "薪资：未公开")
    if job.get("workdays"):
        _para(doc, "每周：%s" % job["workdays"])
    if job.get("duration"):
        _para(doc, "实习周期：%s" % job["duration"])
    if job.get("edu_req"):
        _para(doc, "学历要求：%s" % job["edu_req"])
    if job.get("exp_req"):
        _para(doc, "经验要求：%s" % job["exp_req"])
    if job.get("headcount"):
        _para(doc, "招聘人数：%s" % job["headcount"])
    if job.get("publish_date"):
        _para(doc, "发布/更新：%s" % job["publish_date"])

    if job.get("duty"):
        _para(doc, "岗位职责：", bold=True)
        for line in job["duty"].split("\n")[:6]:
            if line.strip():
                _para(doc, "  " + line.strip(), size=10)
    if job.get("requirement"):
        _para(doc, "任职要求：", bold=True)
        for line in job["requirement"].split("\n")[:6]:
            if line.strip():
                _para(doc, "  " + line.strip(), size=10)
    if job.get("welfare"):
        _para(doc, "福利：%s" % job["welfare"])

    _para(doc, "公司信息：", bold=True)
    _para(doc, "公司名称：%s" % (job.get("company") or "未公开"), size=10)
    _para(doc, "地址：%s" % (job.get("company_address") or "未公开"), size=10)
    _para(doc, "电话：%s" % (job.get("phone") or "未公开"), size=10)
    _para(doc, "官网：%s" % (job.get("website") or "未公开"), size=10)
    for label, key in (("行业", "industry"), ("企业性质", "company_nature"), ("公司规模", "company_scale")):
        v = job.get(key)
        if v:
            _para(doc, "%s：%s" % (label, v), size=10)

    src = job.get("source_site") or "未公开"
    if job.get("sources"):
        src += "（另有来源：" + "、".join(job["sources"]) + "）"
    _para(doc, "来源：%s" % src, size=10)
    _para(doc, "原始链接：%s" % (job.get("url") or "未公开"), size=9, color=(70, 110, 180))
    _para(doc, "信息等级：%s" % (job.get("source_grade") or "D"), size=10)
    if job.get("note"):
        _para(doc, "备注：%s" % job["note"], size=9, color=(120, 120, 120))


def export_word(output_path, profile, stats, main_jobs, expansion_jobs):
    doc = docx.Document()

    # 页面默认字体
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.element.get_or_add_rPr()
    rFonts = style.element.rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "微软雅黑")
    style.font.size = Pt(10.5)

    # ---------- 第一部分：基本信息 ----------
    _para(doc, "《适合你的岗位》", size=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    _para(doc, "生成日期：%s" % stats.get("generated_at", ""))
    _para(doc, "搜索中心：%s" % (stats.get("start_point") or "未填写"))
    _para(doc, "排序原则：先看岗位与实习条件匹配度，再参考来源等级、信息完整度、区域关系和距离；估算距离仅作辅助")
    _para(doc, "学历：%s" % (profile.get("edu_level") or "未确认"))
    _para(doc, "专业：%s" % (profile.get("major") or "未识别"))
    _para(doc, "求职方向：%s" % (" / ".join(profile.get("direction") or ["未明确"]) or "未明确"))
    skills = profile.get("skills") or []
    _para(doc, "技能：%s" % ("、".join(skills) if skills else "未识别"))
    if profile.get("student_status"):
        _para(doc, "学生身份：%s" % profile["student_status"])
    _para(doc, "说明：本文件全部岗位信息来自公开招聘页面/搜索引擎公开摘要，缺失信息一律标注“未公开”，不虚构任何电话、地址、薪资与距离。", size=9, color=(120, 120, 120))
    _para(doc, "")

    # ---------- 第二部分：搜索结果概况 ----------
    _para(doc, "《搜索结果概况》", size=15, bold=True, space_after=6)
    ext_in_word = [j for j in expansion_jobs if j not in main_jobs]
    rows = [
        ("搜索来源数量", str(stats.get("sources_count", 0))),
        ("搜索区域", stats.get("regions", "")),
        ("搜索原则", stats.get("principle", "先看岗位匹配度，再参考来源、信息完整度、区域和距离")),
        ("搜索关键词数量", str(stats.get("query_count", 0))),
        ("访问的搜索入口", stats.get("engines", "")),
        ("原始候选数量", str(stats.get("discovered", 0))),
        ("去重后数量", str(stats.get("deduped", 0))),
        ("过滤后数量", str(stats.get("filtered", 0))),
        ("排除岗位数量", str(stats.get("excluded", 0))),
        ("最终有效数量（主推荐+扩展候选）", str(len(main_jobs) + len(ext_in_word))),
        ("扩展候选数量", str(len(ext_in_word))),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Table Grid"
    for i, (k, v) in enumerate(rows):
        c0, c1 = table.rows[i].cells
        _set_font(c0.paragraphs[0].add_run(k), 10.5, True)
        _set_font(c1.paragraphs[0].add_run(v), 10.5)
    _para(doc, "")

    # ---------- 第三部分：推荐岗位 ----------
    _para(doc, "《推荐岗位》", size=15, bold=True, space_after=4)
    if not any(j.get("distance_km") is not None for j in main_jobs):
        _para(doc, "说明：距离由高德定位（配置 AMAP_KEY 后启用）或按北京各区中心估算得出，Word 中均明确标注距离类型；无法定位的岗位写“距离未计算”。未虚构任何公里数。", size=9, color=(120, 120, 120))
    for i, job in enumerate(main_jobs, 1):
        _job_block(doc, i, job)
        if i < len(main_jobs):
            _para(doc, "")

    # ---------- 第四部分：扩展候选岗位 ----------
    if expansion_jobs:
        doc.add_page_break()
        _para(doc, "《扩展候选岗位》", size=15, bold=True, space_after=4)
        _para(doc, "以下岗位：距离未计算或工作地点未确认/在周边区域，仅供扩展参考。", size=9, color=(120, 120, 120))
        for i, job in enumerate(expansion_jobs, 1):
            _job_block(doc, i, job)
            if i < len(expansion_jobs):
                _para(doc, "")

    # ---------- 第五部分：未完全确认信息 ----------
    doc.add_page_break()
    _para(doc, "《未完全确认信息》", size=15, bold=True, space_after=6)
    all_jobs = main_jobs + expansion_jobs
    items = {
        "电话未公开": sum(1 for j in all_jobs if not j.get("phone")),
        "薪资未公开": sum(1 for j in all_jobs if not j.get("salary")),
        "距离未计算": sum(1 for j in all_jobs if j.get("distance_km") is None),
        "学历要求未明确": sum(1 for j in all_jobs if not j.get("edu_req")),
        "发布时间未确认": sum(1 for j in all_jobs if not j.get("publish_date")),
        "详细地址未公开": sum(1 for j in all_jobs if not (j.get("address") or j.get("company_address"))),
    }
    for k, v in items.items():
        _para(doc, "%s：%d 个岗位" % (k, v))
    _para(doc, "")
    _para(doc, "以上信息均如实标注，未做任何推测补充。", size=9, color=(120, 120, 120))

    temp_path = output_path + ".tmp.docx"
    try:
        doc.save(temp_path)
        os.replace(temp_path, output_path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
    return output_path

# ===================== 模块：main =====================

# -*- coding: utf-8 -*-
"""
一键找实习.py —— 实习岗位雷达 V2 主流程
读取脚本所在文件夹中的简历.docx → 画像 → 搜索矩阵 → 多来源发现 → 抓取/摘要 → 清洗 → 匹配 → 去重 → 排序 → 生成同文件夹中的适合你的岗位.docx

用法：
    python 一键找实习.py            正式模式（生成 Word）
    python 一键找实习.py --search-test   搜索测试模式（只报告真实发现情况，不生成 Word）
"""
import os
import re
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime


APP_DIR = os.path.dirname(os.path.abspath(__file__))
RESUME_PATH = os.path.join(APP_DIR, "简历.docx")
OUTPUT_PATH = os.path.join(APP_DIR, "适合你的岗位.docx")

# 抓取预算（防止无限爬取）
MAX_TOTAL_FETCH = 900      # 总抓取次数
MAX_EXTRA_FETCH = 100      # 列表页扩展详情抓取上限
MAX_DIRECT_FETCH = 60      # 高德地理编码最多调用次数（仅配置 Key 时生效，避免烧额度；无 Key 时本地估算不限次）

JOB_HINT_KW = ["实习", "招聘", "岗位", "职位", "会计", "财务", "出纳", "审计",
               "税务", "助理", "数据", "专员", "校招", "应届", "核算", "报表",
               "对账", "预算", "资金", "应收", "应付", "总账"]
STRONG_JOB_KW = ["实习", "招聘", "岗位", "职位", "校招", "应届", "招聘信息"]
HINT_DOMAINS = ("liepin.com", "zhaopin.com", "51job.com", "zhipin.com",
                "shixiseng.com", "yingjiesheng.com", "nowcoder.com", "lagou.com",
                "ncss.cn", "mohrss.gov.cn", "jobonline.cn", "newjobs.com.cn",
                "bjbys.net.cn", "gov.cn", "edu.cn")
SKIP_TITLE_KW = ["文库", "知道", "百科", "视频", "作文", "心得体会", "报告范文",
                 "论文", "工资待遇", "薪资水平", "招聘网首页", "官网", "下载", "攻略"]

_extra_fetch_count = 0
_fetch_count = 0
_seen_urls = set()
_seen_job_urls = set()
_print_lock = threading.Lock()

# ---- 单文件版附加设施 ----
# 真实控制台（双击）保持系统 Unicode 输出；重定向/管道时用 UTF-8 防止乱码/崩溃
if not sys.stdout.isatty():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

_TIME_LIMIT_SECONDS = 45 * 60          # 最大运行时间 45 分钟（施工令30）
_T0 = time.time()

def _timeout_now():
    """达到时间上限即停止新搜索，保证最终一定生成 Word。"""
    return (time.time() - _T0) >= _TIME_LIMIT_SECONDS

def _timeout_notice():
    print("  [超时] 已达 %d 分钟运行上限，停止继续搜索，进入整理阶段。"
          % (_TIME_LIMIT_SECONDS // 60), flush=True)



def log(msg):
    with _print_lock:
        print(msg, flush=True)


# ---------------------------------------------------------------------------
# 0. 搜索起点（用户每次输入，本次搜索中心）
# ---------------------------------------------------------------------------
def prompt_start_point():
    """提示用户输入实习搜索起点；支持命令行传入通用地点参数。"""
    if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        return sys.argv[1].strip()
    print("请输入你的实习搜索起点：")
    print("例如：北京市某区某地铁站，或某城市某商业区")
    try:
        val = input("> ").strip()
    except EOFError:
        val = ""
    if not val:
        val = (START_POINT or "").strip()
    if not val or "请填写" in val:
        val = "北京市"
        print("未输入起点，暂以北京市作为搜索范围。")
    return val


# ---------------------------------------------------------------------------
# 1. 搜索矩阵
# ---------------------------------------------------------------------------
def build_queries(profile, start_point):
    """搜索区域以用户本次输入的搜索起点动态生成：起点所在区 → 街道词 → 邻居区 → 北京 → 其他区。"""
    job_kws = list(dict.fromkeys(profile.get("job_keywords") or []))
    platforms = ["shixiseng.com", "zhaopin.com", "51job.com", "liepin.com",
                 "zhipin.com", "yingjiesheng.com", "nowcoder.com", "lagou.com"]

    start_district = extract_district(start_point)
    street_word = extract_street_word(start_point)
    neighbors = BJ_NEIGHBORS.get(start_district, [])[:3] if start_district else []
    # 优先级：起点所在区 > 邻居区 > 北京 > 其他区
    priority_regions = [r for r in ([start_district] + neighbors) if r]
    other_regions = [r for r in BJ_DISTRICTS if r not in priority_regions
                     and r != "经开区"]

    core = job_kws[:12]
    queries = []
    seen = set()

    def add(q):
        q = re.sub(r"\s+", " ", q).strip()
        if q and q not in seen:
            seen.add(q)
            queries.append(q)

    # 主批次（按价值排序：平台定向 → 起点区组合 → 街道词 → 北京 → 其他区 → 政府）
    for kw in core[:8]:
        for dom in platforms:
            add("site:%s %s" % (dom, kw))
    for kw in core[:10]:
        for stu in ("在校生", "大专", "应届"):
            for reg in priority_regions[:4]:
                add("%s %s %s" % (kw, stu, reg))
    if street_word and street_word != start_district:
        for kw in core[:6]:
            add("%s %s" % (kw, street_word))
            add("%s 实习 %s" % (kw, street_word))
    for kw in core[:8]:
        add("%s 北京" % kw)
    for reg in other_regions[:8]:
        for kw in core[:3]:
            add("%s 实习 %s" % (kw, reg))
    for q in ("就业见习 北京 会计", "北京 就业见习 财务", "北京毕业生就业创业服务平台 实习",
              "北京高校 大学生 实习 招聘 会计", "24365 实习 北京 会计",
              "北京 大学生 就业 实习 财务", "site:gov.cn 北京 会计 实习 招聘"):
        add(q)

    queries = queries[: MAX_QUERIES]

    # 扩展批次（主批次不足时启用：周边城市）
    ext_queries = []
    for kw in job_kws[12:20]:
        for stu in ("实习", "在校生"):
            for reg in ("燕郊", "廊坊", "北京"):
                q = "%s %s %s" % (kw, stu, reg)
                if q not in seen:
                    seen.add(q)
                    ext_queries.append(q)
    for kw in job_kws[12:20]:
        for dom in platforms[:4]:
            q = "site:%s %s" % (dom, kw)
            if q not in seen:
                seen.add(q)
                ext_queries.append(q)
    return queries, ext_queries


# ---------------------------------------------------------------------------
# 2. 搜索引擎发现
# ---------------------------------------------------------------------------
def engine_discover(queries, callback=None):
    """对每个查询跑搜索引擎，收集候选（resolve 360 跳转）。"""
    candidates = []
    for i, q in enumerate(queries, 1):
        if _timeout_now():
            _timeout_notice()
            break
        # 隔两个查询切一次必应，降低单引擎压力
        forced = "bing" if (i % 3 == 0) else None
        results = search_query(q, forced=forced)
        got = 0
        if isinstance(results, list):
            for item in results:
                title = item.get("title", "")
                engine = item.get("engine", "")
                if any(k in title for k in SKIP_TITLE_KW):
                    continue
                # 360 摘要质量高，弱过滤；必应中文结果多为资讯，需强过滤
                if engine == "bing":
                    url0 = item.get("url", "")
                    if not (any(k in title for k in STRONG_JOB_KW) or
                            any(d in url0 for d in HINT_DOMAINS)):
                        continue
                else:
                    if not any(k in title for k in JOB_HINT_KW):
                        continue
                url = item.get("url", "")
                if url.startswith("https://www.so.com/link"):
                    if got >= 3:
                        continue
                    real = resolve_360_redirect(url)
                    if not real:
                        continue
                    url = real
                    got += 1
                if url in _seen_urls or not url.startswith("http"):
                    continue
                _seen_urls.add(url)
                candidates.append({"url": url, "title": title,
                                   "snippet": item.get("snippet", ""),
                                   "src": "搜索引擎-" + item.get("engine", ""),
                                   "grade": "B", "company_hint": ""})
        if callback:
            callback(i, len(queries), q, len(candidates))
    return candidates


# ---------------------------------------------------------------------------
# 3. 直接来源（A 级）
# ---------------------------------------------------------------------------
def direct_discover(profile):
    candidates = []
    # 猎聘招聘专场
    print("正在发现猎聘公开招聘专场页……", flush=True)
    slugs = discover_liepin_slugs(profile.get("job_keywords") or [])
    print("  猎聘专场页：%d 个（%s）" % (len(slugs), ", ".join(slugs[:8]) + ("…" if len(slugs) > 8 else "")), flush=True)
    for u in liepin_special_job_urls(slugs):
        if u not in _seen_urls:
            _seen_urls.add(u)
            candidates.append({"url": u, "title": "", "snippet": "",
                               "src": "猎聘专场", "grade": "A", "company_hint": ""})
    # 应届生求职网
    print("正在读取应届生求职网北京站……", flush=True)
    for u in yingjiesheng_job_urls():
        if u not in _seen_urls:
            _seen_urls.add(u)
            candidates.append({"url": u, "title": "", "snippet": "",
                               "src": "应届生求职网", "grade": "A", "company_hint": ""})
    # 政府平台直接探测（记录跳过原因，不阻塞）
    gov = probe_gov_sources()
    print("政府/公共就业平台直接探测：可用 %d 个，其余因登录/JS/访问限制跳过（交由搜索引擎发现）" % len(gov), flush=True)
    return candidates


# ---------------------------------------------------------------------------
# 4. 抓取与解析
# ---------------------------------------------------------------------------
def process_candidates(candidates, profile, start_point):
    """抓取候选 → A级(页面) / B级(摘要) → 公司/地点/距离/匹配。返回 jobs。"""
    global _fetch_count, _extra_fetch_count
    jobs = []
    parsed_urls = set()

    def work(item):
        global _fetch_count, _extra_fetch_count
        url = item["url"]
        if _fetch_count >= MAX_TOTAL_FETCH:
            return None
        with _print_lock:
            _fetch_count += 1
        html = fetch(url)
        job = None
        if html:
            text = html_to_text(html, 60000)
            job = parse_page_from_text(url, text, item.get("company_hint", ""))
            if job:
                job["source_grade"] = "A"
                job["source_type"] = item.get("src", "")
                parse_company(job, text)
                # 列表页扩展（预算内）
                global _extra_fetch_count
                if _extra_fetch_count < MAX_EXTRA_FETCH and len(text) < 20000:
                    extra = list_page_detail_urls(html, url, limit=3)
                    extras = []
                    for eu in extra:
                        if eu not in parsed_urls and eu not in _seen_job_urls:
                            parsed_urls.add(eu)
                            _seen_job_urls.add(eu)
                            with _print_lock:
                                _extra_fetch_count += 1
                            extras.append(eu)
                    return job, extras
                return job, []
        if job is None:
            # 抓取失败/非岗位页 → B 级摘要
            if item.get("snippet") or item.get("title"):
                job = parse_snippet(item, grade="B")
                if job and job.get("title"):
                    job["source_type"] = item.get("src", "")
                    return job, []
        return None, []

    # A 级直接候选优先
    ordered = sorted(candidates, key=lambda c: 0 if c["grade"] == "A" else 1)
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {}
        pending_extra = []
        for item in ordered:
            if _timeout_now():
                _timeout_notice()
                break
            if _fetch_count >= MAX_TOTAL_FETCH:
                break
            f = pool.submit(work, item)
            futures[f] = item
            # 控制并发节奏：攒够 MAX_WORKERS*4 个结果再继续
            if len(futures) >= MAX_WORKERS * 4:
                _drain(futures, jobs, parsed_urls, pending_extra, pool)
        _drain(futures, jobs, parsed_urls, pending_extra, pool, last=True)

        # 列表页扩展出来的详情（有预算）
        for extra_url in pending_extra:
            if _timeout_now() or _fetch_count >= MAX_TOTAL_FETCH:
                break
            f = pool.submit(work, {"url": extra_url, "title": "", "snippet": "",
                                   "src": "列表页扩展", "grade": "A", "company_hint": ""})
            futures[f] = {"url": extra_url, "title": "", "snippet": "",
                          "src": "列表页扩展", "grade": "A", "company_hint": ""}
        _drain(futures, jobs, parsed_urls, [], pool, last=True)

    # 地点 → 距离 → 匹配
    dist = ROUTE
    dist_calls = 0
    for job in jobs:
        job["region_tier"] = classify(
            job.get("district", ""),
            (job.get("location_raw") or "") + " " + (job.get("address") or ""),
            profile.get("location_pref") or [],
            start_point)
        if (job.get("address") or job.get("district")) and (not dist.available or dist_calls < MAX_DIRECT_FETCH):
            dist_calls += 1
            res = dist.estimate(start_point, job.get("address"), job.get("district"))
            if res:
                job["distance_km"] = res[0]
                job["distance_type"] = res[1]
        match_job(job, profile)
    return jobs


def _drain(futures, jobs, parsed_urls, extra_out, pool, last=False):
    for f in list(futures.keys()):
        try:
            job, extras = f.result()
        except Exception:
            continue
        if job and job.get("url") not in parsed_urls:
            parsed_urls.add(job["url"])
            jobs.append(job)
        if extras:
            extra_out.extend(extras)
        futures.pop(f, None)


# ---------------------------------------------------------------------------
# 5. 主流程
# ---------------------------------------------------------------------------
def main():
    search_test = "--search-test" in sys.argv

    print("=" * 40)
    print("          一键找实习")
    print("=" * 40)
    print()

    # 第一步：输入搜索起点（本次搜索中心）
    start_point = prompt_start_point()
    print("OK 搜索中心：%s" % start_point)
    print()

    if not os.path.exists(RESUME_PATH):
        print("错误：未找到 %s" % RESUME_PATH)
        print("请把简历（.docx）放到该位置后重新运行。")
        _wait_exit()
        return 1

    # 第一阶段：读取简历
    print("[1/9] 读取简历……", flush=True)
    resume = read_resume(RESUME_PATH)
    if not resume["text"]:
        print("错误：简历内容为空。")
        _wait_exit()
        return 1

    # 第二阶段：画像
    print("[2/9] 分析个人条件……", flush=True)
    prof = analyze(resume["text"])
    print("OK 学校：%s（%s）" % (prof["school"] or "未识别", prof["school_type"] or "层次未确认"))
    print("OK 学历：%s ｜ 专业：%s ｜ 身份：%s" % (prof["edu_level"], prof["major"] or "未识别", prof["student_status"]))
    print("OK 技能：%s" % ("、".join(prof["skills"]) if prof["skills"] else "未识别"))
    print("OK 求职方向：%s" % (" / ".join(prof["direction"]) or "未明确"))
    print("正在建立求职画像……")
    print("正在建立搜索区域……")

    # 第三阶段：搜索矩阵（以用户输入的搜索中心为中心）
    print("[3/9] 生成搜索矩阵……", flush=True)
    queries, ext_queries = build_queries(prof, start_point)
    print("  主批次查询：%d 个关键词组合；扩展批次：%d 个" % (len(queries), len(ext_queries)))

    # 第四阶段：直接来源发现
    print("[4/9] 发现公开招聘来源……", flush=True)
    candidates = direct_discover(prof)
    print("  直接来源候选：%d 条" % len(candidates), flush=True)

    # 第五阶段：搜索引擎发现
    print("[5/9] 搜索引擎检索（%d 组）……" % len(queries), flush=True)

    def prog(i, n, q, total):
        print("  [搜索 %d/%d] 关键词：%s ｜ 累计候选：%d" % (i, n, q, total), flush=True)

    eng_cands = engine_discover(queries, prog)
    print("  搜索引擎候选：%d 条（去重后）" % len(eng_cands), flush=True)

    # 不足时启用扩展批次
    if len(candidates) + len(eng_cands) < 80 and ext_queries:
        print("  候选不足，启用扩展搜索批次（%d 组）……" % len(ext_queries), flush=True)
        more = engine_discover(ext_queries)
        eng_cands.extend(more)
        print("  扩展批次候选：%d 条" % len(more), flush=True)

    all_candidates = candidates + eng_cands
    print("  候选合计：%d 条" % len(all_candidates), flush=True)

    # 第六阶段：抓取/解析/匹配
    print("[6/9] 岗位清洗与匹配……", flush=True)
    jobs = process_candidates(all_candidates, prof, start_point)

    # 第七阶段：先去重，再过滤（统计口径：原始候选 >= 去重后 >= 过滤后 >= 最终输出）
    print("[7/9] 去重与过滤……", flush=True)
    deduped = dedup_jobs(jobs)
    valid = [j for j in deduped if not j.get("excluded")]
    print("  抓取解析：%d 个候选岗位；去重后：%d 个；排除：%d 个；过滤后：%d 个" % (
        len(jobs), len(deduped), len(deduped) - len(valid), len(valid)), flush=True)

    # 第八阶段：排序
    print("[8/9] 排序……", flush=True)
    main_list, expansion_list = split_and_rank(valid, prof)
    final_jobs = finalize(main_list, expansion_list, TARGET_JOBS)
    print("  主推荐：%d 个；扩展候选：%d 个" % (len(main_list), len(expansion_list)), flush=True)

    # 统计
    domains = set()
    for j in deduped:
        if j.get("source_site"):
            domains.add(j["source_site"])
    stats = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "start_point": start_point,
        "principle": "先看岗位匹配度，再参考来源、信息完整度、区域和距离",
        "sources_count": len(domains),
        "regions": "、".join(priority_regions_text(prof, start_point)),
        "query_count": len(queries) + (len(ext_queries) if ext_queries else 0),
        "engines": "360搜索、必应",
        "discovered": len(all_candidates),
        "deduped": len(deduped),
        "filtered": len(valid),
        "excluded": len(deduped) - len(valid),
    }

    # 第九阶段：输出
    if search_test:
        print()
        print("========== 搜索测试报告（未生成 Word） ==========")
        print("原始候选：%d" % stats["discovered"])
        print("去重后：%d" % stats["deduped"])
        print("过滤后：%d" % stats["filtered"])
        print("最终可输出：%d 个（上限 %d）" % (len(final_jobs), TARGET_JOBS))
        print("A级（成功读取原始页面）：%d 个" % sum(1 for j in deduped if j.get("source_grade") == "A"))
        print("B级（搜索摘要确认）：%d 个" % sum(1 for j in deduped if j.get("source_grade") == "B"))
        print("来源域名数：%d" % stats["sources_count"])
        if not ROUTE.available:
            print("距离：未配置高德 Key，按北京各区中心估算直线距离；无法定位的标注“未计算”（不虚构）")
        print("================================================")
        _wait_exit()
        return 0

    print("[9/9] 生成 Word……", flush=True)
    try:
        export_word(OUTPUT_PATH, prof, stats, final_jobs[:TARGET_JOBS], expansion_list)
        print()
        print("完成：")
        print("原始候选：%d 个" % stats["discovered"])
        print("去重后：%d 个" % stats["deduped"])
        print("过滤后：%d 个" % stats["filtered"])
        in_main_ids = set(id(j) for j in final_jobs[:TARGET_JOBS])
        ext_in_word = [j for j in expansion_list if id(j) not in in_main_ids]
        print("最终输出 %d 个有效岗位（主推荐 %d + 扩展候选 %d）" % (
            len(final_jobs[:TARGET_JOBS]) + len(ext_in_word),
            len(final_jobs[:TARGET_JOBS]), len(ext_in_word)))
        print("Word：%s" % OUTPUT_PATH)
        try:
            os.startfile(OUTPUT_PATH)
        except Exception:
            pass
        print()
        if ext_in_word:
            print("另有 %d 个岗位在《扩展候选岗位》部分（距离未计算/地点未确认）" % len(ext_in_word))
    except PermissionError:
        print()
        print("生成 Word 失败：%s 正被 Word 打开。" % OUTPUT_PATH)
        print("请先关闭已打开的《适合你的岗位.docx》窗口，然后重新运行本程序。")
        print("岗位数据已全部就绪，关闭后重跑即可生成。")
        _wait_exit()
        return 1
    except Exception as e:
        print("生成 Word 失败：%s" % e)
        _wait_exit()
        return 1

    _wait_exit()
    return 0


def priority_regions_text(prof, start_point):
    start_d = extract_district(start_point)
    parts = []
    if start_d:
        parts.append(start_d)
    for r in (BJ_NEIGHBORS.get(start_d, [])[:3] if start_d else []):
        if r not in parts:
            parts.append(r)
    if not parts:
        parts = ["大兴", "亦庄", "通州"]
    return parts[:6]


def _wait_exit():
    try:
        input("\n按回车退出。")
    except EOFError:
        pass


if __name__ == "__main__":
    sys.exit(main())