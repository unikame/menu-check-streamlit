#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
メニュー違反チェック（月非依存・汎用版）

run_check.py で7月/8月に対して個別実装していたロジックを、
アップロードされたワークブックのシート名から月を自動検出して
何ヶ月分でもまとめて処理できるように一般化したモジュール。

前提とするシート構成（アップロードされる1つの.xlsxの中に、月ごとに以下がある想定）:
  - "{月}月使用食材"      : 必須。食材データ（商品ID/商品名/レシピ名/食材数量 等の列を含む）
  - "{月}月昼夕..."        : 推奨。昼夜ペアのメニュー名一覧
                              （列構成: A=昼普通/夜普通ラベル, B=日付(昼行のみ), 以降5ポジション分の
                                「タグ,メニュー名」または「メニュー名のみ」の列が並ぶ）
  - "{月}月昼"             : 任意。調理法列付きの昼メニュー（No.8/12/19/20 等の参考実装に使用）

夜（夕）の食材データは、同じワークブック内に「{月}月_夜食材」のようなシートがあればそれを使い、
無ければ別ファイルでアップロードされたCSV（ファイル名に「{月}月」と「夜食材」を含むもの）を使う。
どちらも無い月は、夜データに依存するルールをスキップする（結果に注記を残す）。
"""
import os
import re
import sys
import datetime
from collections import Counter

import pandas as pd
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, 'scripts'))
import check_menu as cm  # noqa: E402

WD_JP = ['月', '火', '水', '木', '金', '土', '日']
cols_std = ['日付', '曜日', 'No', 'ルール', '該当箇所', '理由', '修正提案', '重要度']

BASE_SEASONING_KW = ['だし', '醤油', '塩こしょう', '片栗粉', '上白糖', 'コンソメ', 'ウスターソース',
                      'ごま油', '味覇', 'シャンタン', 'みりん', '料理酒', '食塩', 'こしょう', '胡椒',
                      '酢', 'マヨネーズ', 'ケチャップ', 'ソース', 'たれ', 'ダレ', 'あん', 'スパイス',
                      '昆布茶', 'スープ', '味噌', '砂糖', 'サラダ油', '唐辛子', 'わさび',
                      '辛子', 'からし', 'ラー油', '豆板醤', 'オイスター', 'ポン酢', 'カレー粉', '小麦粉']
SOY_KW = ['豆腐', 'がんも', 'おから', '卯の花', 'うの花', '豆乳', '高野豆腐', '厚揚げ', '油揚げ', '生揚げ', '湯葉', '大豆', '納豆']
FISH_KW = ['鮭', 'サーモン', 'さば', 'サバ', 'あじ', 'アジ', 'ぶり', 'ブリ', '白身魚', 'カレイ', 'かれい',
           'まぐろ', 'マグロ', '鮪', 'さわら', 'サワラ', 'たら', 'タラ', 'いわし', 'イワシ', 'さんま', 'サンマ',
           'ほっけ', 'ホッケ', '金目鯛', '鯛', 'タイ', 'アブラカレイ', 'ししゃも', 'シシャモ']
NG_WORDS = ['餅', 'ヤングコーン', 'パスタこんにゃく', '旨辛ジャン', '麻婆']
HEALTH_KW = ['アジ', 'あじ', 'イワシ', 'いわし', 'サバ', 'さば', 'サンマ', 'さんま', 'マグロ', 'まぐろ',
             'レモン', '酢の物', 'ほうれん草', 'あさり', 'アサリ', '豆乳', 'ひじき']
EAT_NG = {
    'イカ': '硬さ(歯で嚙み切れない)', 'ビーフジャーキー': '硬さ(歯で嚙み切れない)',
    'こんにゃく': '弾力が強すぎる', 'ナタデココ': '弾力が強すぎる', 'ホルモン': '弾力が強すぎる',
    'ごぼう': '繊維質', 'セロリ': '繊維質', '牛すじ': '繊維質',
    'もち': '口に張り付く', '海苔': '口に張り付く', 'ウエハース': '口に張り付く',
    '鶏むね': 'パサつく', '焼き魚': 'パサつく', '茹でたまご': 'パサつく', 'ゆで卵': 'パサつく',
    'おから': 'ぼそぼそする', '玄米': 'ぼそぼそする', 'スコーン': 'ぼそぼそする',
    '春巻き': '口腔内を傷つける', 'チップス': '口腔内を傷つける', '乾燥した小魚': '口腔内を傷つける', '田作り': '口腔内を傷つける',
}
LOOKALIKE_PAIRS = [
    (['海老カツ', '海老タラカツ'], 8),
    (['塩じゃがコロッケ', 'ビーフ入りコロッケ'], 8),
    (['さつまいもコロッケ', 'かぼちゃコロッケ'], 8),
    (['ガツンとジューシーメンチ', 'ビーフ入りメンチカツ'], 8),
]
VEG_TIERS = [
    (1, ['ほうれん草']),
    (1, ['インゲンカット', 'ミニミニブロッコリー', 'オクラ', 'ささがきごぼう', '大根乱切り', '冷凍かぶ', '竹の子千切り', 'スナップピース']),
]
FISH_FD_ONLY = ['いわしの梅煮', 'マスの塩焼き', 'サーモン塩焼き', 'ぶりのみぞれ', 'さばの味噌煮',
                'さわらの西京焼き', 'タラの香草焼き', 'あじみりん焼き', 'あじの塩焼き']
NUTRI_BOUNDS = {'昼': {'kcal': (415, 455), 'salt_max': 3.8, 'protein_min': 12},
                '夜': {'kcal': (245, 275), 'salt_max': 3.0, 'protein_min': 12}}


def is_base_seasoning(prod):
    return any(k in str(prod) for k in BASE_SEASONING_KW)


def is_soy(name):
    return any(k in str(name) for k in SOY_KW)


def is_fish(name):
    return any(k in str(name) for k in FISH_KW)


def is_health(name):
    return any(k in str(name) for k in HEALTH_KW)


def canon_name(name):
    for group, _ in LOOKALIKE_PAIRS:
        for g in group:
            if g in name:
                return group[0]
    return cm.norm_recipe(name)


class MenuData:
    """1つのアップロードから抽出した、月ごとのデータをまとめて保持する"""

    def __init__(self):
        self.months = []                 # [7, 8, ...]
        self.shoku = {}                  # month -> DataFrame (昼, build_day_index済み)
        self.shoku_night = {}            # month -> DataFrame (夜, build_day_index済み) ※無ければキー無し
        self.menu_lunch_tagged = {}      # month -> (df, pos_def) 調理法/タグ付き昼メニュー ※無ければキー無し
        self.rows = []                   # [(date, weekday_jp, slot, pos, name), ...] 全月分
        self.warnings = []               # 注記（画面に表示する）
        self.date_range = None


def _find_month_sheets(sheet_names, suffix_regex):
    result = {}
    for sn in sheet_names:
        m = re.match(r'^(\d{1,2})月' + suffix_regex, sn)
        if m:
            month = int(m.group(1))
            result.setdefault(month, sn)
    return result


def _detect_lunchdinner_layout(ws, date_col=2, label_col=1, sample_rows=30):
    """列構成が「タグ,名前」のペア型(7月昼夕比較: 名前列=4,6,8,10,12、5品分で列Lまで使う)か、
    「名前のみ」の単列型(8月昼夕: 名前列=4,5,6,7,8、5品分で列Hまでしか使わない)かを判定する。
    ペア型は5品目の名前が列12(L)付近まで伸びるのに対し、単列型は列8(H)までしか埋まらない
    ため、列9〜13(I〜M)に値があるかどうかを判定材料にする（奇偶列だけを見ると
    単列型でも偶然半分が奇数列に当たってしまい誤判定するため使わない）。"""
    cur_date = None
    paired_votes = 0
    straight_votes = 0
    checked = 0
    for r in range(1, ws.max_row + 1):
        if checked >= sample_rows:
            break
        label = ws.cell(row=r, column=label_col).value
        if label is None:
            continue
        d = ws.cell(row=r, column=date_col).value
        if isinstance(d, datetime.datetime):
            cur_date = d
        if cur_date is None:
            continue
        vals_3_8 = [ws.cell(row=r, column=c).value for c in range(3, 9)]
        if sum(1 for v in vals_3_8 if v not in (None, '')) == 0:
            continue
        checked += 1
        beyond8 = any(ws.cell(row=r, column=c).value not in (None, '') for c in range(9, 14))
        if beyond8:
            paired_votes += 1
        else:
            straight_votes += 1
    if paired_votes >= straight_votes:
        return {4: 'メイン', 6: 'サブ', 8: '副菜1', 10: '副菜2', 12: 'サラダ'}
    return {4: 'メイン', 5: 'サブ', 6: '副菜1', 7: '副菜2', 8: 'サラダ'}


def _parse_lunch_dinner_sheet(ws, date_col=2, label_col=1):
    """「N月昼夕...」系シートを (date, weekday_jp, slot, pos, name) のリストにする。
    列構成は 7月昼夕比較(タグ,名前 の2列セット×5) と 8月昼夕(名前のみ×5) の両パターンに対応。"""
    pos_map = _detect_lunchdinner_layout(ws, date_col, label_col)
    rows = []
    cur_date = None
    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=label_col).value
        if label is None:
            continue
        d = ws.cell(row=r, column=date_col).value
        if isinstance(d, datetime.datetime):
            cur_date = d
        if cur_date is None:
            continue
        slot = '夜' if '夜' in str(label) else ('昼' if '昼' in str(label) else str(label))
        for c, pos in pos_map.items():
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                rows.append((cur_date.date(), WD_JP[cur_date.weekday()], slot, pos, v.strip()))
    return rows


def load_workbook_data(xlsx_path, night_csv_paths=None):
    """xlsx_path: メインのメニューワークブック。
    night_csv_paths: {month: csv_path} 夜食材CSV（任意）。"""
    night_csv_paths = night_csv_paths or {}
    data = MenuData()

    xl = pd.ExcelFile(xlsx_path)
    wb_raw = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet_names = xl.sheet_names

    shoku_sheets = _find_month_sheets(sheet_names, r'使用食材$')
    lunchdinner_sheets = _find_month_sheets(sheet_names, r'昼夕')
    lunch_tagged_sheets = _find_month_sheets(sheet_names, r'昼$')
    night_sheets = _find_month_sheets(sheet_names, r'.*夜食材$')

    months = sorted(set(shoku_sheets) | set(lunchdinner_sheets))
    if not months:
        raise ValueError('「N月使用食材」「N月昼夕...」形式のシートが見つかりませんでした。シート名をご確認ください。')
    data.months = months

    for month in months:
        if month in shoku_sheets:
            df = pd.read_excel(xl, shoku_sheets[month], header=0)
            data.shoku[month] = cm.build_day_index(df)
        else:
            data.warnings.append(f'{month}月：「{month}月使用食材」シートが見つからず、食材ベースの判定をスキップしました')

        if month in night_sheets:
            df = pd.read_excel(xl, night_sheets[month], header=0)
            data.shoku_night[month] = cm.build_day_index(df)
        elif month in night_csv_paths:
            try:
                df = pd.read_csv(night_csv_paths[month])
                data.shoku_night[month] = cm.build_day_index(df)
            except Exception as e:
                data.warnings.append(f'{month}月：夜食材CSVの読み込みに失敗しました（{e}）')
        else:
            data.warnings.append(f'{month}月：夜（夕）の食材データが見つからず、昼夜合算が必要な一部ルールをスキップしました')

        if month in lunchdinner_sheets:
            ws = wb_raw[lunchdinner_sheets[month]]
            data.rows += _parse_lunch_dinner_sheet(ws)
        else:
            data.warnings.append(f'{month}月：「{month}月昼夕...」シートが見つからず、No.24/28/29の判定をスキップしました')

        if month in lunch_tagged_sheets:
            df = pd.read_excel(xl, lunch_tagged_sheets[month], header=None)
            data.menu_lunch_tagged[month] = df

    all_dates = [r[0] for r in data.rows]
    if all_dates:
        data.date_range = pd.date_range(min(all_dates), max(all_dates))
    else:
        # rowsが無い場合はshokuの日付から範囲を作る
        mds = []
        for month, shoku in data.shoku.items():
            for md in shoku['md'].dropna().unique():
                try:
                    mds.append(datetime.date(2000 + (0 if month >= 1 else 0), int(md[0]), int(md[1])))
                except Exception:
                    pass
        if mds:
            data.date_range = pd.date_range(min(mds), max(mds))
    return data


def _shoku_for(data, month, night=False):
    d = data.shoku_night if night else data.shoku
    return d.get(month)


def raw_dish_names(data, date):
    names = set()
    month = date.month
    md = (date.month, date.day)
    for shoku in [data.shoku.get(month), data.shoku_night.get(month)]:
        if shoku is None:
            continue
        sub = shoku[(shoku['md'] == md) & (shoku['isDX'])]
        for _, r in sub.iterrows():
            prod = str(r['商品名'])
            qty = r.get('食材数量')
            if pd.isna(qty) or qty == 0 or cm.is_noise(prod):
                continue
            rn = str(r['レシピ名'])
            if rn and '備品' not in rn:
                names.add(rn)
    return names


def dishes_products(data, date):
    """(レシピ名 -> [商品名,...]) を昼夜合算で返す（基礎調味料含む）"""
    month = date.month
    md = (date.month, date.day)
    out = {}
    for shoku in [data.shoku.get(month), data.shoku_night.get(month)]:
        if shoku is None:
            continue
        sub = shoku[(shoku['md'] == md) & (shoku['isDX'])]
        for _, r in sub.iterrows():
            prod = str(r['商品名'])
            qty = r.get('食材数量')
            if pd.isna(qty) or qty == 0 or cm.is_noise(prod):
                continue
            out.setdefault(str(r['レシピ名']), []).append(prod)
    return out


def _min_gap_check(data, match_fn, min_gap, rule_no, rule_name, severity='中'):
    """スペーシング型：min_gap日以内の再使用はNG（例：かにのふわふわ5日以上空ける）"""
    dr = data.date_range
    dates_with = []
    for d in dr:
        names = raw_dish_names(data, d)
        hit = [n for n in names if match_fn(n)]
        if hit:
            dates_with.append((d, hit))
    viol = []
    for i in range(1, len(dates_with)):
        d0, _ = dates_with[i - 1]
        d1, h1 = dates_with[i]
        gap = (d1 - d0).days
        if gap <= min_gap:
            viol.append({
                '日付': d1.strftime('%-m/%-d'), '曜日': WD_JP[d1.weekday()], 'No': rule_no, 'ルール': rule_name,
                '該当箇所': f'前回{d0.strftime("%-m/%-d")} → 今回{d1.strftime("%-m/%-d")}:{h1[0][:16]}',
                '理由': f'{gap}日しか空いていない（要{min_gap + 1}日以上）',
                '修正提案': '使用日をずらす', '重要度': severity,
            })
    return pd.DataFrame(viol)


def _max_gap_check(data, match_fn, max_gap, rule_no, rule_name, severity='中'):
    """頻度型：max_gap日を超えて使用が無いとNG（例：魚メニューは3日に1回以上）"""
    dr = data.date_range
    dates_with = []
    for d in dr:
        names = raw_dish_names(data, d)
        hit = [n for n in names if match_fn(n)]
        if hit:
            dates_with.append((d, hit))
    viol = []
    if dates_with:
        gap0 = (dates_with[0][0] - dr[0]).days
        if gap0 > max_gap:
            viol.append({
                '日付': dr[0].strftime('%-m/%-d'), '曜日': WD_JP[dr[0].weekday()], 'No': rule_no,
                'ルール': rule_name + '（期間開始〜初回）',
                '該当箇所': f'{dr[0].strftime("%-m/%-d")}〜{dates_with[0][0].strftime("%-m/%-d")}',
                '理由': f'{gap0}日間使用なし', '修正提案': '追加を検討', '重要度': severity,
            })
    for i in range(1, len(dates_with)):
        d0, _ = dates_with[i - 1]
        d1, h1 = dates_with[i]
        gap = (d1 - d0).days
        if gap > max_gap:
            viol.append({
                '日付': d1.strftime('%-m/%-d'), '曜日': WD_JP[d1.weekday()], 'No': rule_no, 'ルール': rule_name,
                '該当箇所': f'前回{d0.strftime("%-m/%-d")} → 今回{d1.strftime("%-m/%-d")}:{h1[0][:16]}',
                '理由': f'{gap}日間使用なし（上限{max_gap}日）',
                '修正提案': '間隔内に追加', '重要度': severity,
            })
    return pd.DataFrame(viol)


# ---------------- 各ルールの判定関数（run_check.py の最終版ロジックを月非依存化） ----------------

def check_rule1(data, positions=('メイン', 'サブ')):
    """No.1: メイン/サブ商材（酷似品含む）を1週間以上空けて使用（昼のみ判定：従来通り昼メニューのメイン/サブを対象）"""
    if not data.rows:
        return pd.DataFrame()
    by_date = {}
    for (d, wd, slot, pos, name) in data.rows:
        if slot == '昼' and pos in positions:
            by_date.setdefault(d, {})[pos] = name
    days_sorted = sorted(by_date.items())
    last_seen = {}
    viol = []
    for d, posmap in days_sorted:
        for p, name in posmap.items():
            key = canon_name(name)
            if not key:
                continue
            if key in last_seen:
                prev_date, prev_p, prev_name = last_seen[key]
                gap = (d - prev_date).days
                if 0 < gap <= 7:
                    viol.append({
                        '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[d.weekday()], 'No': 1,
                        'ルール': 'メイン/サブ商材（酷似品含む）を1週間空けず再使用',
                        '該当箇所': f'{p}:{name[:20]}（前回 {prev_date.strftime("%-m/%-d")} {prev_p}:{prev_name[:16]}）',
                        '理由': f'{gap}日しか空いていない（要8日以上）',
                        '修正提案': '該当日か次回使用日をずらす', '重要度': '高',
                    })
            last_seen[key] = (d, p, name)
    return pd.DataFrame(viol)


def check_rule3_5(data):
    """No.3(挽肉重複) / No.5(鶏豚牛重複、No.34のハンバーグ例外込み)（昼のみ判定）"""
    if not data.rows:
        return pd.DataFrame(), pd.DataFrame()
    by_date = {}
    for (d, wd, slot, pos, name) in data.rows:
        if slot == '昼' and pos in ('メイン', 'サブ'):
            by_date.setdefault(d, {})[pos] = name
    v3, v5 = [], []
    for d, posmap in sorted(by_date.items()):
        nm_m, nm_s = posmap.get('メイン', ''), posmap.get('サブ', '')
        if not nm_m or not nm_s:
            continue
        gm, gs = cm.group_from_name(nm_m), cm.group_from_name(nm_s)
        if gm != gs:
            continue
        if gm == 'ひき肉系':
            is_exception = ('豆腐ハンバーグ' in nm_m and 'ハンバーグ' in nm_s and '豆腐ハンバーグ' not in nm_s) or \
                            ('豆腐ハンバーグ' in nm_s and 'ハンバーグ' in nm_m and '豆腐ハンバーグ' not in nm_m)
            if is_exception:
                continue
            v3.append({
                '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[d.weekday()], 'No': 3,
                'ルール': '挽肉商材がメイン・サブメインで同日重複',
                '該当箇所': f'メイン:{nm_m[:16]} / サブ:{nm_s[:16]}', '理由': f'両方「{gm}」',
                '修正提案': 'メインかサブの系統を変える', '重要度': '高',
            })
        if gm in ('鶏肉系', '豚肉系', '牛肉系'):
            v5.append({
                '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[d.weekday()], 'No': 5,
                'ルール': '鶏豚牛が同日でメイン・サブメインに重複（枠をずらす）',
                '該当箇所': f'メイン:{nm_m[:16]} / サブ:{nm_s[:16]}', '理由': f'両方「{gm}」',
                '修正提案': 'メインかサブの系統を変える', '重要度': '高',
            })
    return pd.DataFrame(v3), pd.DataFrame(v5)


def check_rule4_36(data):
    """No.4(コロッケ1日空け) / No.36(通常⇔クリームは連日可、の例外込み)"""
    dr = data.date_range
    hits = []
    for d in dr:
        names = raw_dish_names(data, d)
        for n in names:
            if 'コロッケ' in n:
                cat = 'クリーム' if 'クリーム' in n else '通常'
                hits.append((d, n, cat))
    hits.sort(key=lambda x: x[0])
    viol = []
    last_seen = {}
    for d, n, cat in hits:
        if cat in last_seen:
            pd0, pn0 = last_seen[cat]
            gap = (d - pd0).days
            if gap <= 1:
                viol.append({
                    '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[d.weekday()], 'No': 4,
                    'ルール': f'コロッケ({cat})が連日重複',
                    '該当箇所': f'{n[:16]}（前回 {pd0.strftime("%-m/%-d")} {pn0[:16]}）',
                    '理由': f'{gap}日しか空いていない', '修正提案': '使用日をずらす', '重要度': '中',
                })
        last_seen[cat] = (d, n)
    return pd.DataFrame(viol)


def check_rule6(data):
    """No.6: 1食内（昼/夜は別々）で同一食材が複数レシピに重複使用（基礎調味料除外）"""
    dr = data.date_range
    viol = []
    for d in dr:
        month = d.month
        for meal, shoku in [('昼', data.shoku.get(month)), ('夜', data.shoku_night.get(month))]:
            if shoku is None:
                continue
            md = (d.month, d.day)
            sub = shoku[(shoku['md'] == md) & (shoku['isDX'])]
            prod_recipes = {}
            for _, r in sub.iterrows():
                prod = str(r['商品名'])
                qty = r.get('食材数量')
                if pd.isna(qty) or qty == 0 or cm.is_noise(prod) or is_base_seasoning(prod):
                    continue
                recipe = cm.norm_recipe(r['レシピ名'])
                if not recipe or '備品' in recipe:
                    continue
                prod_recipes.setdefault(prod, set()).add(recipe)
            for prod, recipes in prod_recipes.items():
                if len(recipes) >= 2:
                    viol.append({
                        '日付': d.strftime('%-m/%-d'), '曜日': meal, 'No': 6,
                        'ルール': '1食内で同一食材が複数レシピに重複使用',
                        '該当箇所': f'[{meal}] {prod[:26]} → ' + ' / '.join(list(recipes)[:4]),
                        '理由': f'{len(recipes)}レシピで使用', '修正提案': 'いずれかを別食材に', '重要度': '中',
                    })
    return pd.DataFrame(viol)


def check_rule7(data):
    """No.7: 大豆系商材は同日（昼夜合算）内で複数使用NG"""
    dr = data.date_range
    viol = []
    for d in dr:
        names = raw_dish_names(data, d)
        soy = sorted(n for n in names if is_soy(n))
        if len(soy) >= 2:
            viol.append({
                '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[d.weekday()], 'No': 7,
                'ルール': '大豆系商材が同日内で複数使用（半日未満）',
                '該当箇所': ' / '.join(soy), '理由': f'同日に大豆系が{len(soy)}品',
                '修正提案': '一方を翌日以降にずらす', '重要度': '中',
            })
    return pd.DataFrame(viol)


def check_rule12(data):
    """No.12: 当日揚げ調理は3品まで（調理法タグがある月のみ判定可）"""
    viol = []
    pos_def7 = [('メイン', 2, 3, 4), ('サブ', 5, 6, 7), ('副菜①', 8, 9, 10), ('副菜②', 11, 12, 13), ('サラダ', 14, 15, 16)]
    for month, menu in data.menu_lunch_tagged.items():
        if menu.shape[1] <= 16:
            continue  # 調理法列が無い形式（8月昼のような3列無しパターン）は対象外
        for i in range(menu.shape[0]):
            d = menu.iloc[i, 0]
            if pd.isna(d):
                continue
            try:
                date = pd.to_datetime(d)
            except Exception:
                continue
            fried = []
            for nm_, cmcol, nmcol, smcol in pos_def7:
                if menu.shape[1] > cmcol and pd.notna(menu.iloc[i, cmcol]):
                    cook = str(menu.iloc[i, cmcol]).strip()
                    if cook == '揚':
                        fried.append(nm_)
            if len(fried) >= 4:
                viol.append({
                    '日付': date.strftime('%-m/%-d'), '曜日': WD_JP[date.weekday()], 'No': 12,
                    'ルール': '当日揚げ調理が3品を超過', '該当箇所': '/'.join(fried),
                    '理由': f'揚げ物が{len(fried)}品（上限3品）', '修正提案': '1品を煮/和え等に', '重要度': '中',
                })
    return pd.DataFrame(viol)


def check_rule14(data):
    """No.14: 75歳以上向け栄養素基準（月平均・75歳基準）"""
    viol = []
    for month, shoku in data.shoku.items():
        if not all(c in shoku.columns for c in ['カロリー', 'たんぱく質', '食塩相当量']):
            continue
        sub = shoku[shoku['isDX']]
        daily = sub.groupby('md')[['カロリー', 'たんぱく質', '食塩相当量']].sum().reset_index()
        if not len(daily):
            continue
        bounds = NUTRI_BOUNDS['昼']
        avg = daily[['カロリー', 'たんぱく質', '食塩相当量']].mean()
        label = f'{month}月'
        if avg['カロリー'] < bounds['kcal'][0] or avg['カロリー'] > bounds['kcal'][1]:
            pos = '下限未達' if avg['カロリー'] < bounds['kcal'][0] else '上限超過'
            viol.append({'日付': label, '曜日': '昼', 'No': 14, 'ルール': f'エネルギー月平均が{pos}',
                         '該当箇所': f'{label}昼・月平均', '理由': f'{avg["カロリー"]:.0f}kcal（基準{bounds["kcal"][0]}-{bounds["kcal"][1]}kcal）',
                         '修正提案': '全体のメニュー量・構成を見直す', '重要度': '高'})
        if avg['食塩相当量'] > bounds['salt_max']:
            viol.append({'日付': label, '曜日': '昼', 'No': 14, 'ルール': '食塩相当量の月平均が上限超過',
                         '該当箇所': f'{label}昼・月平均', '理由': f'{avg["食塩相当量"]:.2f}g（上限{bounds["salt_max"]}g）',
                         '修正提案': '調味料量を見直す', '重要度': '中'})
        if avg['たんぱく質'] < bounds['protein_min']:
            viol.append({'日付': label, '曜日': '昼', 'No': 14, 'ルール': 'たんぱく質の月平均が下限未達',
                         '該当箇所': f'{label}昼・月平均', '理由': f'{avg["たんぱく質"]:.1f}g（下限{bounds["protein_min"]}g）',
                         '修正提案': 'たんぱく質を含む食材を増やす', '重要度': '中'})
    return pd.DataFrame(viol)


def check_rule15(data):
    """No.15: 週に1回、健康食材を使用する"""
    return _max_gap_check(data, is_health, 7, 15, '健康食材の使用間隔が週1回を下回る')


def check_rule21(data):
    """No.21: 禁止食材・調味料の使用禁止"""
    pattern = '|'.join(NG_WORDS)
    viol = []
    for month in data.months:
        for label, shoku in [('昼', data.shoku.get(month)), ('夜', data.shoku_night.get(month))]:
            if shoku is None:
                continue
            hit = shoku[shoku['商品名'].astype(str).str.contains(pattern, na=False) |
                        shoku['レシピ名'].astype(str).str.contains(pattern, na=False)].copy()
            for (md, recipe), grp in hit.groupby(['md', 'レシピ名'], sort=False):
                if pd.isna(md):
                    continue
                m, d = md
                matched_words = set()
                matched_prods = []
                for _, r in grp.iterrows():
                    for w in NG_WORDS:
                        if w in str(r['商品名']) or w in str(recipe):
                            matched_words.add(w)
                            if w in str(r['商品名']):
                                matched_prods.append(str(r['商品名']))
                matched_prods = list(dict.fromkeys(matched_prods))[:3]
                viol.append({
                    '日付': f'{m}/{d}', '曜日': label, 'No': 21, 'ルール': '禁止食材・調味料の使用',
                    '該当箇所': f'{str(recipe)[:22]}' + (f' → {"/".join(matched_prods)[:30]}' if matched_prods else ''),
                    '理由': f'禁止ワード「{"/".join(sorted(matched_words))}」に該当',
                    '修正提案': '代替食材/調味料に変更', '重要度': '高',
                })
    return pd.DataFrame(viol)


def check_rule22(data):
    """No.22: 魚メニューは3日に1回（昼夕併せて）"""
    return _max_gap_check(data, is_fish, 3, 22, '魚メニューの間隔が3日を超過')


def check_rule23(data):
    """No.23: 食べにくさチェックリスト該当（参考実装）"""
    dr = data.date_range
    viol = []
    seen = set()
    for d in dr:
        names = raw_dish_names(data, d)
        for n in names:
            for kw, reason in EAT_NG.items():
                if kw == 'イカ':
                    hit = bool(re.search(r'(?<!ス)イカ', n))
                else:
                    hit = kw in n
                if hit:
                    key = (d.strftime('%-m/%-d'), n, kw)
                    if key in seen:
                        continue
                    seen.add(key)
                    viol.append({
                        '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[d.weekday()], 'No': 23,
                        'ルール': '食べにくさチェックリスト該当（要確認）',
                        '該当箇所': f'{n[:20]}（{kw}）', '理由': f'NG例「{kw}」に該当＝{reason}',
                        '修正提案': '商品開発部のたべやすさ基準で再確認', '重要度': '低（参考実装・人による最終判断が必要）',
                    })
    return pd.DataFrame(viol)


def check_rule25(data):
    """No.25: かぼちゃは週1回以上、同一曜日は4週間空ける"""
    dr = data.date_range
    kabocha_dates = []
    for d in dr:
        names = raw_dish_names(data, d)
        hit = [n for n in names if 'かぼちゃ' in n]
        if hit:
            kabocha_dates.append((d, hit[0]))
    viol = []
    for i in range(1, len(kabocha_dates)):
        d0, _ = kabocha_dates[i - 1]
        d1, n1 = kabocha_dates[i]
        gap = (d1 - d0).days
        if gap > 7:
            viol.append({
                '日付': d1.strftime('%-m/%-d'), '曜日': WD_JP[d1.weekday()], 'No': 25,
                'ルール': 'かぼちゃの使用間隔が週1回を下回る',
                '該当箇所': f'前回{d0.strftime("%-m/%-d")} → 今回{d1.strftime("%-m/%-d")}:{n1[:16]}',
                '理由': f'{gap}日間かぼちゃなし', '修正提案': '間の週にかぼちゃメニューを追加', '重要度': '中',
            })
    by_weekday = {}
    for d, n in kabocha_dates:
        wd = d.weekday()
        if wd in by_weekday:
            pd0, pn0 = by_weekday[wd]
            gap = (d - pd0).days
            if gap <= 28:
                viol.append({
                    '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[wd], 'No': 25,
                    'ルール': 'かぼちゃが同一曜日で4週間以内に再使用',
                    '該当箇所': f'前回{pd0.strftime("%-m/%-d")}:{pn0[:14]} → 今回{d.strftime("%-m/%-d")}:{n[:14]}',
                    '理由': f'同一曜日で{gap}日しか空いていない（要29日以上）', '修正提案': '曜日をずらすか間隔を空ける', '重要度': '中',
                })
        by_weekday[wd] = (d, n)
    return pd.DataFrame(viol)


def check_rule26(data):
    """No.26: かにのふわふわは5日以上空けて使用"""
    return _min_gap_check(data, lambda n: 'かに' in n and 'ふわふわ' in n, 5, 26,
                           'かにのふわふわが5日以内に再使用', severity='低')


def check_rule30(data):
    """No.30: 野菜使用の間隔（FDメニュールール（野菜）反映・参考実装）"""
    dr = data.date_range
    viol = []

    def veg_products_on(date):
        month = date.month
        md = (date.month, date.day)
        prods = set()
        for shoku in [data.shoku.get(month), data.shoku_night.get(month)]:
            if shoku is None:
                continue
            sub = shoku[(shoku['md'] == md) & (shoku['isDX'])]
            for _, r in sub.iterrows():
                qty = r.get('食材数量')
                if pd.isna(qty) or qty == 0:
                    continue
                prods.add(str(r['商品名']))
        return prods

    for min_gap, kws in VEG_TIERS:
        for kw in kws:
            dates_with = []
            for d in dr:
                prods = veg_products_on(d)
                hit = [p for p in prods if kw in p]
                if hit:
                    dates_with.append((d, hit[0]))
            for i in range(1, len(dates_with)):
                d0, n0 = dates_with[i - 1]
                d1, n1 = dates_with[i]
                gap = (d1 - d0).days
                if gap <= min_gap:
                    viol.append({
                        '日付': d1.strftime('%-m/%-d'), '曜日': WD_JP[d1.weekday()], 'No': 30,
                        'ルール': '野菜(FDメニュールール)の使用間隔違反',
                        '該当箇所': f'{n1[:20]}（前回{d0.strftime("%-m/%-d")}）',
                        '理由': f'{gap}日しか空いていない（要{min_gap + 1}日以上）',
                        '修正提案': '使用日をずらす', '重要度': '低（参考実装）',
                    })
    return pd.DataFrame(viol)


def check_rule27(data):
    """No.27: FD専用商材（魚弁当のみ、確定5商材）は平日に入れる（参考実装）"""
    dr = data.date_range
    viol = []
    for d in dr:
        names = raw_dish_names(data, d)
        for n in names:
            for kw in FISH_FD_ONLY:
                if kw in n:
                    wd = d.weekday()
                    if wd >= 5:
                        viol.append({
                            '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[wd], 'No': 27,
                            'ルール': 'FD専用商材（魚弁当）は平日に入れる',
                            '該当箇所': n[:30],
                            '理由': f'FD専用魚商材「{kw}」が休日（{WD_JP[wd]}）に使用されている',
                            '修正提案': '平日の枠に振り替える', '重要度': '中（参考実装・魚弁当のみ）',
                        })
    return pd.DataFrame(viol)


def check_rule28(data):
    """No.28: 本日の魚料理は平日の夜に採用する"""
    viol = []
    for (d, wd, slot, pos, name) in data.rows:
        if '本日の魚料理' in name:
            problems = []
            if slot != '夜':
                problems.append(f'{slot}に使用（要:夜）')
            if wd in ('土', '日'):
                problems.append(f'{wd}曜（休日）に使用（要:平日）')
            if problems:
                viol.append({
                    '日付': d.strftime('%-m/%-d'), '曜日': wd, 'No': 28,
                    'ルール': '本日の魚料理は平日の夜に採用する',
                    '該当箇所': name[:30],
                    '理由': ' / '.join(problems),
                    '修正提案': '平日の夜枠に振り替える', '重要度': '中',
                })
    return pd.DataFrame(viol)


def check_rule29(data):
    """No.29: おまかせメニューを昼・夜月2回以上採用"""
    omakase_count = Counter()
    seen_months = set()
    for (d, wd, slot, pos, name) in data.rows:
        seen_months.add(d.month)
        if 'おまかせ' in name:
            omakase_count[(d.month, slot)] += 1
    viol = []
    for month in sorted(seen_months):
        for slot in ['昼', '夜']:
            cnt = omakase_count.get((month, slot), 0)
            if cnt < 2:
                viol.append({
                    '日付': f'{month}月', '曜日': '-', 'No': 29,
                    'ルール': 'おまかせメニューを昼・夜月2回以上採用',
                    '該当箇所': f'{month}月{slot}',
                    '理由': f'おまかせメニューが{cnt}回のみ（月2回以上必要）',
                    '修正提案': f'{slot}のおまかせ枠を追加する', '重要度': '中',
                })
    return pd.DataFrame(viol)


def check_rule24(data):
    """No.24: 白和えは副菜に分類し、サラダ使用時は酢を混ぜる"""
    viol = []
    for (d, wd, slot, pos, name) in data.rows:
        if '白和え' in name:
            if pos == 'サラダ' and '酢' not in name:
                viol.append({
                    '日付': d.strftime('%-m/%-d'), '曜日': wd, 'No': 24,
                    'ルール': '白和えは副菜に分類し、サラダ使用時は酢を混ぜる',
                    '該当箇所': f'{slot}:{pos}:{name[:20]}',
                    '理由': 'サラダ枠での白和え使用だが「酢」の明記がない',
                    '修正提案': '酢を加える、または副菜枠に変更する', '重要度': '低',
                })
            elif pos not in ('副菜1', '副菜2', 'サラダ'):
                viol.append({
                    '日付': d.strftime('%-m/%-d'), '曜日': wd, 'No': 24,
                    'ルール': '白和えは副菜に分類し、サラダ使用時は酢を混ぜる',
                    '該当箇所': f'{slot}:{pos}:{name[:20]}',
                    '理由': f'白和えが副菜以外（{pos}）で使用されている',
                    '修正提案': '副菜枠に変更する', '重要度': '低',
                })
    return pd.DataFrame(viol)


def check_rule31(data):
    """No.31: マッシュ系調理、メニュー名に明記があれば同日2品以上NG"""
    dr = data.date_range
    viol = []
    for d in dr:
        names = raw_dish_names(data, d)
        mash = sorted(n for n in names if 'マッシュ' in n)
        if len(mash) >= 2:
            viol.append({
                '日付': d.strftime('%-m/%-d'), '曜日': WD_JP[d.weekday()], 'No': 31,
                'ルール': 'メニュー名に「マッシュ」と明記された商材が同日2品以上',
                '該当箇所': ' / '.join(mash), '理由': f'マッシュ系が{len(mash)}品',
                '修正提案': '一方を別の調理法名に', '重要度': '低',
            })
    return pd.DataFrame(viol)


ALL_RULES = [
    ('No.1 メイン/サブ商材1週間ルール（酷似商材含む）', check_rule1),
    ('No.3/5 挽肉・鶏豚牛のメイン/サブ重複', check_rule3_5),
    ('No.4/36 コロッケ間隔', check_rule4_36),
    ('No.6 同一食材の複数レシピ重複（1食内）', check_rule6),
    ('No.7 大豆系商材の同日重複', check_rule7),
    ('No.12 揚げ物3品まで', check_rule12),
    ('No.14 栄養素基準（月平均）', check_rule14),
    ('No.15 健康食材 週1回以上', check_rule15),
    ('No.21 禁止食材・調味料', check_rule21),
    ('No.22 魚メニュー3日に1回', check_rule22),
    ('No.23 食べにくさチェックリスト', check_rule23),
    ('No.24 白和えの分類', check_rule24),
    ('No.25 かぼちゃ週1回・同曜日4週間', check_rule25),
    ('No.26 かにのふわふわ5日以上空ける', check_rule26),
    ('No.27 FD専用商材（魚弁当）は平日', check_rule27),
    ('No.28 本日の魚料理は平日の夜', check_rule28),
    ('No.29 おまかせメニュー月2回以上', check_rule29),
    ('No.30 野菜使用間隔（FDメニュールール）', check_rule30),
    ('No.31 マッシュ系同日重複', check_rule31),
]


def run_all_checks(xlsx_path, night_csv_paths=None):
    data = load_workbook_data(xlsx_path, night_csv_paths)
    frames = []
    for label, fn in ALL_RULES:
        try:
            result = fn(data)
        except Exception as e:
            data.warnings.append(f'{label} の判定中にエラーが発生しスキップしました（{e}）')
            continue
        if isinstance(result, tuple):
            for r in result:
                if len(r):
                    frames.append(r.reindex(columns=cols_std))
        else:
            if len(result):
                frames.append(result.reindex(columns=cols_std))
    if frames:
        combined = pd.concat(frames, ignore_index=True)
        combined['No'] = combined['No'].astype(int)
        combined = combined.sort_values(['No', '日付']).reset_index(drop=True)
    else:
        combined = pd.DataFrame(columns=cols_std)
    n_days = len(set(r[0] for r in data.rows)) if data.rows else 0
    summary = {
        'months': data.months,
        'n_days': n_days,
        'total': len(combined),
        'by_rule': combined['No'].value_counts().sort_index().to_dict() if len(combined) else {},
        'warnings': data.warnings,
    }
    return combined, summary


def write_report(combined, n_days, out_path):
    cm.write_report(combined, n_days, out_path)
