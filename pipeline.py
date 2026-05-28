#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SmileBuild 落札情報パイプライン
================================
毎日 GitHub Actions で実行され、以下を行う：

  P1  関東地整／PPI から入札・契約結果を取得         -> fetch_results()
  P2  入札調書の明細から「入札参加者」を取得          -> fetch_bidders()
  P3  経審/許可データと突合して従業員数・所在地を補強  -> enrich()
  蓄積  全期間を history.json に貯める（履歴・年間集計の素）
  出力  ツール用の data.json を生成（会社集計はツール側で実施）
  通知  前回からの「新着の該当案件」をメール送信

【重要・キャリブレーション】
  役所サイトの実ファイル構造は独特で、中身を見ずに完璧に抜くのは不可能です。
  実データから値を抜く 3 つの関数（fetch_results / fetch_bidders / enrich）の
  中の "★CALIBRATE★" の箇所だけ、実ファイルを 1 度開いて合わせてください。
  それ以外（蓄積・集計・data.json生成・差分検知・メール）は完成しています。

  まず動作確認は:  python pipeline.py --sample      (ダミーで全工程を通す)
  実運用は:        python pipeline.py                (実取得)
"""

import os, sys, json, hashlib, smtplib, datetime, re, time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ---- 任意依存（実取得時のみ必要。--sample では不要）----
try:
    import requests
    from bs4 import BeautifulSoup
except Exception:
    requests = None
    BeautifulSoup = None

# =========================================================
#  設定（GitHub Secrets / 環境変数で上書き可）
# =========================================================
ROOT      = os.path.dirname(os.path.abspath(__file__))
HISTORY   = os.path.join(ROOT, "history.json")     # 全期間の蓄積
DATA_JSON = os.path.join(ROOT, "data.json")        # ツールが読む最新データ
STATE     = os.path.join(ROOT, "state.json")       # 前回通知済みID

# 通知の条件（＝弘晃のターゲット）。Secrets で調整可。
NOTIFY_MIN_YEN = int(os.environ.get("NOTIFY_MIN_YEN", 100_000_000))   # 1億
NOTIFY_MAX_YEN = int(os.environ.get("NOTIFY_MAX_YEN", 200_000_000))   # 2億
NOTIFY_PREFS   = [p for p in os.environ.get("NOTIFY_PREFS", "").split(",") if p]  # 空=全部

# メール（GitHub Secrets に登録）
SMTP_HOST = os.environ.get("SMTP_HOST", "")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")
MAIL_FROM = os.environ.get("MAIL_FROM", SMTP_USER)
MAIL_TO   = os.environ.get("MAIL_TO", "")

KTR_BASE  = "https://www.ktr.mlit.go.jp"

# ---- データ源：国交省 全地方整備局（工事のみ）----
# PPI(入札情報サービス)が全整備局を横断する公式の統合検索。工事専用の検索画面がある：
#   工事検索  https://www.i-ppi.jp/ippi/SearchServices/web/Koji/Kokoku/Search.aspx
# PPIは.aspxフォームのため取得には Playwright が必要（USE_PPI=1 で有効化）。
USE_PPI       = os.environ.get("USE_PPI", "0") == "1"
PPI_KOJI_URL  = "https://www.i-ppi.jp/ippi/SearchServices/web/Koji/Kokoku/Search.aspx"  # ★工事専用★
LOOKBACK_DAYS = int(os.environ.get("LOOKBACK_DAYS", "45"))   # 何日分さかのぼるか

# PPIを使わない場合：各地方整備局の入札結果ページを巡回。
# ★CALIBRATE★ 実URL/年度ディレクトリは年度替わりで変わるため、稼働時に最新を設定。
BUREAUS = [
    ("北海道開発局",     "https://www.hkd.mlit.go.jp/"),
    ("東北地方整備局",   "https://www.thr.mlit.go.jp/"),
    ("関東地方整備局",   "https://www.ktr.mlit.go.jp/nyuusatu/index00000028.html"),
    ("北陸地方整備局",   "https://www.hrr.mlit.go.jp/"),
    ("中部地方整備局",   "https://www.cbr.mlit.go.jp/"),
    ("近畿地方整備局",   "https://www.kkr.mlit.go.jp/"),
    ("中国地方整備局",   "https://www.cgr.mlit.go.jp/"),
    ("四国地方整備局",   "https://www.skr.mlit.go.jp/"),
    ("九州地方整備局",   "https://www.qsr.mlit.go.jp/"),
    ("沖縄総合事務局",   "https://www.dc.ogb.go.jp/"),
]
# 特定の整備局だけにしたい場合は Secrets BUREAUS_ONLY="関東,東北" を設定（空=全部）
BUREAUS_ONLY = [b for b in os.environ.get("BUREAUS_ONLY", "").split(",") if b]

TODAY = datetime.date.today().isoformat()

# ---- 工事のみ抽出（測量・建設コンサルタント等の"業務"を除外）----
_GYOMU_RE = re.compile(
    r"業務|委託|コンサル|設計(?!施工)|測量|調査|点検|診断|計画策定|検討|"
    r"監督支援|積算|技術審査|補償|用地|資料作成|データ作成|システム|"
    r"観測|解析|試験|研究|清掃|運転|運用|警備|賃貸|借[上り]|購入|物品|役務")
_KOJI_RE = re.compile(
    r"工事|工区|舗装|護岸|改良|拡幅|補修|補強|築堤|樋管|樋門|橋梁|橋りょう|"
    r"トンネル|ずい道|法面|のり面|電線共同溝|無電柱|区画線|防護柵|擁壁|"
    r"盛土|切土|浚渫|河道|耐震|塗装|遮音|排水|管渠|構築|新設|設置工")
def is_works(name="", kind="", category=""):
    """工事=True / 測量・建設コンサルタント等の業務=False。"""
    cat = str(category or "")
    if re.search(r"工事", cat) and not re.search(r"業務|コンサル|委託", cat):
        return True
    if re.search(r"業務|コンサル|委託", cat):
        return False
    blob = "%s %s %s" % (cat, kind, name)
    if re.search(r"設計施工|設計・施工|DB方式", blob):   # 設計施工一括は工事
        return True
    if _GYOMU_RE.search(blob):
        return False
    if _KOJI_RE.search(blob):
        return True
    return True   # 工事検索前提。判定不能は工事扱い（業務語があれば上で除外済み）

def log(*a):
    print("[%s]" % datetime.datetime.now().strftime("%H:%M:%S"), *a, flush=True)

def case_id(c):
    key = "%s|%s|%s|%s|%s" % (c.get("date"), c.get("name"), c.get("company"), c.get("org"), c.get("amount"))
    return hashlib.md5(key.encode("utf-8")).hexdigest()[:12]

def pref_of(s):
    if not s: return ""
    m = re.search(r"(東京都|北海道|(?:京都|大阪)府|(?:神奈川|和歌山|鹿児島)県|..県)", str(s))
    return m.group(1) if m else ""

def city_of(s):
    if not s: return ""
    p = pref_of(s); s = str(s)
    rest = s[s.find(p)+len(p):] if p else s
    m = re.match(r"\s*([^\d０-９]{1,10}?[市区町村郡])", rest)
    return m.group(1) if m else ""

# =========================================================
#  P1  入札・契約結果の取得
# =========================================================
def fetch_results(sample=False):
    """落札結果のリストを返す。各要素は data モデル(dict)。工事のみ。"""
    if sample:
        cases = _sample_cases()
    elif USE_PPI:
        cases = _fetch_ppi()
    else:
        cases = _fetch_bureaus()

    # 工事のみに限定（測量・建設コンサルタント等の業務を除外）
    before = len(cases)
    cases = [c for c in cases if is_works(c.get("name", ""), c.get("kind", ""), c.get("category", ""))]
    if before:
        log("工事フィルタ:", before, "→", len(cases), "件（業務・コンサルを除外）")
    log("P1 取得:", len(cases), "件")
    return cases

def _fetch_bureaus():
    """全地方整備局の入札結果ページを巡回して落札（工事）を集める。"""
    if requests is None:
        log("requests/bs4 が無いため取得スキップ。`pip install -r requirements.txt`")
        return []
    targets = [(n, u) for (n, u) in BUREAUS if (not BUREAUS_ONLY or any(k in n for k in BUREAUS_ONLY))]
    log("巡回対象:", len(targets), "整備局",
        ("（"+",".join(BUREAUS_ONLY)+"に限定）") if BUREAUS_ONLY else "（全部）")
    cases = []
    for bureau, url in targets:
        try:
            html = requests.get(url, timeout=30,
                                headers={"User-Agent": "SmileBuild-bot/1.0 (info@smile-build.com)"}).text
            soup = BeautifulSoup(html, "html.parser")
            # ★CALIBRATE★ 各整備局の入札結果テーブル（または結果Excelへのリンク）に合わせて調整。
            #   多くは「工事」「業務」が別ページ/別ファイルで公表されるため、工事側のみを辿るのが理想。
            #   どうしても混在する場合も、後段の is_works() で業務は除外される。
            n0 = len(cases)
            for tr in soup.select("table tr"):
                tds = [td.get_text(strip=True) for td in tr.select("td")]
                if len(tds) < 4:
                    continue
                rec = _blank_rec()
                rec.update({
                    "date":    _norm_date(tds[0]),
                    "name":    tds[1],
                    "org":     (tds[2] if len(tds) > 2 and tds[2] else bureau),
                    "company": tds[3] if len(tds) > 3 else "",
                    "amount":  _norm_yen(tds[4]) if len(tds) > 4 else None,
                    "rate":    _norm_num(tds[5]) if len(tds) > 5 else None,
                    "bureau":  bureau,
                })
                if rec["company"] and rec["name"]:
                    rec["kind"] = classify_kind(rec["name"])
                    cases.append(rec)
            log("  %s: %d件" % (bureau, len(cases) - n0))
            time.sleep(1.0)  # 礼儀
        except Exception as e:
            log("  %s 取得エラー: %s" % (bureau, e))
    return cases

def _fetch_ppi():
    """PPI(入札情報サービス)の『工事』検索で全整備局横断に取得（Playwright）。
       ★CALIBRATE★ フォーム項目（発注機関=全整備局/開札日範囲/工事種別）と
       結果テーブルのセレクタを、実画面に合わせて1度だけ調整する。"""
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        log("Playwright未導入。`pip install playwright && playwright install chromium`。USE_PPI=0で各整備局巡回に切替可")
        return []
    cases = []
    since = (datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS))
    try:
        with sync_playwright() as p:
            br = p.chromium.launch(headless=True)
            pg = br.new_page()
            pg.goto(PPI_KOJI_URL, timeout=60000)   # ★工事専用URL（業務は別画面なので源で除外）
            # ★CALIBRATE★ ここで検索条件を入力：
            #   - 発注機関：国土交通省（全地方整備局）にチェック/選択
            #   - 開札日 from = since.isoformat()  / to = TODAY
            #   - 区分：落札結果
            #   例) pg.fill("#kaisatsu_from", since.strftime("%Y/%m/%d"))
            #       pg.click("#searchButton")
            #   結果テーブルを行ごとに読み、必要なら次ページへ。
            #   for row in pg.query_selector_all("table#result tr"): ...
            #       rec = _blank_rec(); rec.update({...}); cases.append(rec)
            br.close()
    except Exception as e:
        log("PPI取得エラー:", e)
    return cases

def _blank_rec():
    return {
        "date": "", "name": "", "org": "", "company": "",
        "amount": None, "rate": None, "kind": "", "period": "",
        "pref": "", "city": "", "employees": None,
        "category": "", "bureau": "",
        "docs": [], "bidders": [], "detail_url": "",
    }

KIND_RULES = [
    ("舗装", ["舗装"]), ("河川改修", ["河川", "築堤", "護岸", "樋管", "樋門"]),
    ("トンネル", ["トンネル", "ずい道"]), ("橋梁補修", ["橋", "橋梁"]),
    ("電線共同溝", ["電線共同溝", "無電柱"]), ("法面工", ["法面", "のり面"]),
    ("道路改良", ["道路", "改良", "拡幅"]), ("維持修繕", ["維持", "修繕", "補修"]),
]
def classify_kind(name):
    for label, keys in KIND_RULES:
        if any(k in (name or "") for k in keys):
            return label
    return "その他"

# =========================================================
#  P2  入札調書の明細 → 入札参加者
# =========================================================
def fetch_bidders(case):
    """その案件に応札した会社のリスト [{company, pref, amount}] を返す。"""
    url = case.get("detail_url")
    if not url or requests is None:
        return []
    bidders = []
    try:
        html = requests.get(url, timeout=30,
                            headers={"User-Agent": "SmileBuild-bot/1.0"}).text
        soup = BeautifulSoup(html, "html.parser")
        # ★CALIBRATE★ 入札調書の「参加業者・入札額」の表に合わせて取り出す
        for tr in soup.select("table tr"):
            tds = [td.get_text(strip=True) for td in tr.select("td")]
            if len(tds) < 2:
                continue
            name = tds[0]
            amt  = _norm_yen(tds[1])
            if name and not re.search(r"業者名|商号|会社", name):
                bidders.append({"company": name, "pref": pref_of(name), "amount": amt})
        time.sleep(1.0)  # 礼儀: 1秒空ける
    except Exception as e:
        log("fetch_bidders エラー:", e)
    return bidders

# =========================================================
#  P3  経審/許可データ突合 → 従業員数・所在地
# =========================================================
_KEISHIN = None
def _load_keishin():
    """経審/建設業許可の会社マスタを {正規化会社名: {pref, employees}} で返す。
       ★CALIBRATE★ 実データ(経審結果CSV等)を keishin.csv として置き、列を合わせる。"""
    global _KEISHIN
    if _KEISHIN is not None:
        return _KEISHIN
    _KEISHIN = {}
    path = os.path.join(ROOT, "keishin.csv")
    if not os.path.exists(path):
        log("keishin.csv 無し → 従業員数/所在地の突合はスキップ（任意）")
        return _KEISHIN
    import csv
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            name = _normalize_co(row.get("商号") or row.get("会社名") or "")
            if not name:
                continue
            emp = row.get("職員数") or row.get("従業員数") or ""
            addr = row.get("所在地") or row.get("住所") or ""
            _KEISHIN[name] = {
                "pref": pref_of(addr),
                "city": city_of(addr),
                "employees": int(re.sub(r"\D", "", emp)) if re.search(r"\d", emp) else None,
            }
    log("経審マスタ:", len(_KEISHIN), "社")
    return _KEISHIN

def _normalize_co(s):
    # 会社名の表記ゆれ吸収（突合用）。完全一致は難しいので取りこぼしは出る。
    s = str(s or "")
    s = re.sub(r"[（(].*?[）)]", "", s)
    s = s.replace("株式会社", "").replace("有限会社", "").replace("(株)", "").replace("（株）", "")
    return re.sub(r"\s|　", "", s)

# ---- gBizINFO（経済産業省・無料API）で 所在地＋従業員数 を取得 ----
# 無料トークンを https://info.gbiz.go.jp/ で発行し、Secrets/環境変数 GBIZ_TOKEN に設定。
GBIZ_TOKEN = os.environ.get("GBIZ_TOKEN", "")
def _gbiz_lookup(name):
    if not (GBIZ_TOKEN and requests and name):
        return None
    try:
        r = requests.get("https://info.gbiz.go.jp/hojin/v1/hojin",
                         params={"name": name, "limit": 1},
                         headers={"X-hojinInfo-api-token": GBIZ_TOKEN, "Accept": "application/json"},
                         timeout=20)
        if r.status_code != 200:
            return None
        infos = (r.json() or {}).get("hojin-infos") or []
        if not infos:
            return None
        h = infos[0]
        loc = h.get("location") or ""
        emp = h.get("employee_number")
        return {"pref": pref_of(loc), "city": city_of(loc),
                "employees": int(emp) if (emp not in (None, "")) else None}
    except Exception as e:
        log("gBiz lookup失敗:", name, e)
        return None

def enrich(cases):
    """所在地（県・市）と従業員数を自動補完。
       1) gBizINFO API（GBIZ_TOKEN があれば）  2) keishin.csv（あれば）の順で照合。"""
    ks = _load_keishin()
    cache = {}
    def lookup(name):
        key = _normalize_co(name)
        if not key:
            return None
        if key in cache:
            return cache[key]
        info = _gbiz_lookup(name)
        if GBIZ_TOKEN:
            time.sleep(0.3)   # APIへの礼儀（レート制限回避）
        if not info and ks:
            info = ks.get(key)
        cache[key] = info
        return info
    hit = 0
    for c in cases:
        info = lookup(c.get("company"))
        if info:
            if not c.get("pref") and info.get("pref"): c["pref"] = info["pref"]
            if not c.get("city") and info.get("city"): c["city"] = info["city"]
            if c.get("employees") is None and info.get("employees") is not None: c["employees"] = info["employees"]
            hit += 1
        for b in c.get("bidders", []):
            bi = lookup(b.get("company"))
            if bi:
                b["pref"] = b.get("pref") or bi.get("pref", "")
                b["city"] = b.get("city") or bi.get("city", "")
                b["employees"] = bi.get("employees")
    log("所在地・従業員 補完:", hit, "/", len(cases),
        "（gBizトークン:", "有" if GBIZ_TOKEN else "無", "/ 経審マスタ:", len(ks), "社）")

# =========================================================
#  蓄積・出力・通知
# =========================================================
def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return default

def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)

def merge_history(history_cases, new_cases):
    by_id = {c["id"]: c for c in history_cases}
    added = []
    for c in new_cases:
        c["id"] = case_id(c)
        if c["id"] not in by_id:
            by_id[c["id"]] = c
            added.append(c)
        else:
            by_id[c["id"]].update({k: v for k, v in c.items() if v})  # 後勝ちで補強
    merged = sorted(by_id.values(), key=lambda x: x.get("date") or "", reverse=True)
    return merged, added

def matches_target(c):
    a = c.get("amount")
    if a is None or not (NOTIFY_MIN_YEN <= a <= NOTIFY_MAX_YEN):
        return False
    if NOTIFY_PREFS and c.get("pref") and c["pref"] not in NOTIFY_PREFS:
        return False
    return True

def send_email(new_targets, total_new):
    if not (SMTP_HOST and MAIL_TO):
        log("メール設定が無いため送信スキップ（新着該当 %d件）" % len(new_targets))
        return
    if not new_targets:
        log("新着の該当案件なし → メール送らず")
        return
    rows = ""
    for c in new_targets:
        rows += ("<tr><td style='padding:6px 10px;border-bottom:1px solid #eee'>%s</td>"
                 "<td style='padding:6px 10px;border-bottom:1px solid #eee;font-weight:700;color:#1b2e3c'>%s</td>"
                 "<td style='padding:6px 10px;border-bottom:1px solid #eee'>%s</td>"
                 "<td style='padding:6px 10px;border-bottom:1px solid #eee'>%s</td>"
                 "<td style='padding:6px 10px;border-bottom:1px solid #eee;text-align:right'>%s万円</td></tr>") % (
                 c.get("date",""), c.get("company",""), c.get("pref",""), c.get("name",""),
                 "{:,}".format(int((c.get("amount") or 0)/10000)))
    body = """
    <div style="font-family:sans-serif;color:#1b2e3c">
      <h2 style="color:#2d7a4f;margin:0 0 4px">本日の新着ターゲット {n} 件</h2>
      <p style="color:#5a6a72;font-size:13px;margin:0 0 14px">
        条件：落札 {lo}〜{hi}万円{pf}　／　本日の取得 {tot} 件中、条件に合致した新着のみ抜粋</p>
      <table style="border-collapse:collapse;font-size:13px;min-width:560px">
        <tr style="background:#1b2e3c;color:#fff">
          <th style="padding:7px 10px;text-align:left">落札日</th><th style="padding:7px 10px;text-align:left">落札者</th>
          <th style="padding:7px 10px;text-align:left">県</th><th style="padding:7px 10px;text-align:left">工事名</th>
          <th style="padding:7px 10px;text-align:right">落札金額</th></tr>
        {rows}
      </table>
      <p style="color:#5a6a72;font-size:12px;margin-top:14px">詳細・過去履歴・入札参加ランキングはツールで確認 → 落札情報フィルタ</p>
    </div>""".format(n=len(new_targets), tot=total_new, rows=rows,
                     lo="{:,}".format(int(NOTIFY_MIN_YEN/10000)), hi="{:,}".format(int(NOTIFY_MAX_YEN/10000)),
                     pf=("／"+ "・".join(NOTIFY_PREFS) if NOTIFY_PREFS else ""))
    msg = MIMEMultipart("alternative")
    msg["Subject"] = "【SmileBuild】本日の新着ターゲット %d件（%s）" % (len(new_targets), TODAY)
    msg["From"], msg["To"] = MAIL_FROM, MAIL_TO
    msg.attach(MIMEText(body, "html", "utf-8"))
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as s:
        s.starttls()
        if SMTP_USER:
            s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(MAIL_FROM, [a.strip() for a in MAIL_TO.split(",")], msg.as_string())
    log("メール送信:", len(new_targets), "件 →", MAIL_TO)

# ---- 取得値の正規化 ----
def _norm_date(s):
    if not s: return ""
    s = re.sub(r"[年/]", "-", str(s)).replace("月", "-").replace("日", "")
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3))) if m else ""
def _norm_yen(s):
    if s is None: return None
    t = re.sub(r"[^\d]", "", str(s));  return int(t) if t else None
def _norm_num(s):
    m = re.search(r"\d+(\.\d+)?", str(s or ""));  return float(m.group()) if m else None

# =========================================================
#  メイン
# =========================================================
def main(sample=False):
    log("=== START (sample=%s) ===" % sample)

    # 1. 取得
    new_cases = fetch_results(sample=sample)

    # 2. P2 参加者（detail_url がある案件のみ）
    if not sample:
        for c in new_cases:
            if c.get("detail_url"):
                c["bidders"] = fetch_bidders(c)

    # 3. P3 補強
    enrich(new_cases)

    # 4. 蓄積
    history = load_json(HISTORY, {"cases": []}).get("cases", [])
    merged, added = merge_history(history, new_cases)
    save_json(HISTORY, {"updated": TODAY, "cases": merged})
    log("蓄積:", len(merged), "件（新規", len(added), "件）")

    # 5. ツール用 data.json（直近2年に絞ると軽い。全件にしたければ merged をそのまま）
    cutoff = (datetime.date.today() - datetime.timedelta(days=730)).isoformat()
    recent = [c for c in merged if (c.get("date") or "") >= cutoff] or merged
    save_json(DATA_JSON, {"updated": TODAY, "cases": recent})
    log("data.json 出力:", len(recent), "件")

    # 6. 差分通知（前回未通知 かつ 条件合致）
    state = load_json(STATE, {"notified": []})
    notified = set(state["notified"])
    new_targets = [c for c in added if matches_target(c) and c["id"] not in notified]
    send_email(new_targets, len(added))
    state["notified"] = list(notified | {c["id"] for c in added if matches_target(c)})
    save_json(STATE, state)

    log("=== DONE ===")

# =========================================================
#  ダミーデータ（--sample 用：全工程を通して動作確認するため）
# =========================================================
def _sample_cases():
    base = datetime.date.today()
    cos = [("東関建設(株)","千葉県",58),("房総土木(株)","千葉県",34),("関東道路(株)","神奈川県",72),
           ("常陸建設(株)","茨城県",41),("多摩川組(株)","東京都",63),("上総建設(株)","千葉県",88)]
    orgs = ["大宮国道事務所","横浜国道事務所","千葉国道事務所","常陸河川国道事務所"]
    out = []
    for i in range(10):
        w = cos[i % len(cos)]
        amt = (80 + i*18) * 1_000_000
        d = (base - datetime.timedelta(days=i*7)).isoformat()
        bidders = [{"company": w[0], "pref": w[1], "employees": w[2], "amount": amt}]
        for j in range(2):
            b = cos[(i+j+1) % len(cos)]
            bidders.append({"company": b[0], "pref": b[1], "employees": b[2], "amount": int(amt*1.05)})
        out.append({
            "date": d, "name": "国道%d号 ○○%s工事" % (16+i, ["道路改良","河川改修","舗装"][i%3]),
            "org": orgs[i % len(orgs)], "company": w[0], "pref": w[1], "employees": w[2],
            "amount": amt, "rate": 90.0 + (i % 9), "kind": classify_kind("道路改良"),
            "period": d + "〜" + (base + datetime.timedelta(days=200)).isoformat(),
            "docs": [{"label": "入札調書", "url": KTR_BASE + "/"}], "bidders": bidders, "detail_url": "",
        })
    # 業務（測量・コンサル等）も混在させる → 工事フィルタで除外されるのを確認できる
    for k, gname in enumerate(["一般国道○○号 橋梁点検業務", "△△川 河川測量業務",
                               "□□地区 道路詳細設計業務", "管内 用地補償コンサルタント業務"]):
        w = cos[k % len(cos)]
        r = _blank_rec()
        r.update({"date": (base - datetime.timedelta(days=k*5)).isoformat(), "name": gname,
                  "org": orgs[k % len(orgs)], "company": w[0], "pref": w[1], "employees": w[2],
                  "amount": (30 + k*12) * 1_000_000, "rate": 92.0 + k, "category": "業務"})
        out.append(r)
    return out


if __name__ == "__main__":
    main(sample=("--sample" in sys.argv))
