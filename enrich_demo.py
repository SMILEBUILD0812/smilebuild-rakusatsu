#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
所在地・従業員数の自動取得デモ（gBizINFO 経済産業省・無料API）
================================================================
これは「落札くん」に所在地（県・市）と従業員数が自動で入るのを"見る"ためのスクリプト。

なぜブラウザ単体でやらないか：
  gBizINFO はトークンが必要で、ブラウザから直接叩くと CORS で弾かれ、
  トークンも画面に露出して危険。だから取得は手元(Python)で行い、
  結果を「落札くん」に渡して表示する。

使い方（5分）:
  1) https://info.gbiz.go.jp/ で無料の API トークンを発行（メールで届く）
  2) 取得したい会社名を companies.txt に1行ずつ書く（例：落札した会社名）
       例)
         株式会社○○建設
         △△土木株式会社
  3) 実行:
        GBIZ_TOKEN=あなたのトークン  python enrich_demo.py
     （Windowsなら:  set GBIZ_TOKEN=... の後  python enrich_demo.py）
  4) 画面に 所在地・従業員数 が表示され、master_gbiz.json が出力される
  5) その master_gbiz.json を「落札くん」にドラッグ＆ドロップ
     → 取り込んだ落札データの所在地・従業員数が自動で埋まる

※ gBizINFO は約400万法人を収録。ただし従業員数は公開がある会社のみ入る（全社は埋まらない）。
"""

import os, sys, json, re, time
try:
    import requests
except Exception:
    requests = None

TOKEN = os.environ.get("GBIZ_TOKEN", "")
HERE  = os.path.dirname(os.path.abspath(__file__))

def pref_of(s):
    m = re.search(r"(東京都|北海道|(?:京都|大阪)府|(?:神奈川|和歌山|鹿児島)県|..県)", str(s or ""))
    return m.group(1) if m else ""
def city_of(s):
    s = str(s or ""); p = pref_of(s)
    rest = s[s.find(p)+len(p):] if p else s
    m = re.match(r"\s*([^\d０-９]{1,10}?[市区町村郡])", rest)
    return m.group(1) if m else ""
def norm_co(s):
    s = str(s or "")
    s = re.sub(r"[（(].*?[）)]", "", s)
    s = re.sub(r"株式会社|有限会社|\(株\)|（株）|\(有\)|（有）|合同会社|協同組合", "", s)
    return re.sub(r"\s|　", "", s)

def lookup(name):
    r = requests.get("https://info.gbiz.go.jp/hojin/v1/hojin",
                     params={"name": name, "limit": 1},
                     headers={"X-hojinInfo-api-token": TOKEN, "Accept": "application/json"},
                     timeout=20)
    if r.status_code != 200:
        return {"error": "HTTP %s" % r.status_code}
    infos = (r.json() or {}).get("hojin-infos") or []
    if not infos:
        return None
    h = infos[0]; loc = h.get("location") or ""
    emp = h.get("employee_number")
    return {"matched": h.get("name"), "loc": loc, "pref": pref_of(loc), "city": city_of(loc),
            "employees": int(emp) if emp not in (None, "") else None}

def main():
    if requests is None:
        print("requests が必要です：  pip install requests"); return
    if not TOKEN:
        print("GBIZ_TOKEN が未設定です。https://info.gbiz.go.jp/ で無料トークンを発行し、")
        print("  GBIZ_TOKEN=xxxx python enrich_demo.py  の形で実行してください。"); return

    path = os.path.join(HERE, "companies.txt")
    if os.path.exists(path):
        names = [l.strip() for l in open(path, encoding="utf-8") if l.strip()]
    else:
        names = ["鹿島建設株式会社", "前田建設工業株式会社", "株式会社安藤・間"]  # 例（companies.txt が無い場合）
        print("companies.txt が無いので例の会社名で実行します。\n")

    master = {}
    print("%-28s %-12s %-8s" % ("会社名", "所在地", "従業員数"))
    print("-" * 56)
    for nm in names:
        try:
            info = lookup(nm); time.sleep(0.3)
        except Exception as e:
            info = {"error": str(e)}
        if not info:
            print("%-28s %s" % (nm, "（gBizに該当なし）")); continue
        if info.get("error"):
            print("%-28s %s" % (nm, "エラー: " + info["error"])); continue
        loc = (info["pref"] + info["city"]) or info["loc"] or "—"
        emp = info["employees"]
        print("%-28s %-12s %-8s" % (nm, loc, (str(emp)+"人") if emp is not None else "—"))
        master[norm_co(nm)] = {"pref": info["pref"], "city": info["city"], "employees": emp}

    out = os.path.join(HERE, "master_gbiz.json")
    json.dump({"master": master}, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("\n→ %s を書き出しました。これを「落札くん」にドラッグすると所在地・従業員数が埋まります。" % out)

if __name__ == "__main__":
    main()
