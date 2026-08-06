#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
やどかり弁当 メニュー違反チェック（Streamlit画面）

menu_check_core.run_all_checks() を呼び出して結果を表示するだけの薄いUI層。
判定ロジックは一切ここに書かない（ルールの追加・修正は menu_check_core.py 側で行う）。

■ この版での変更点（ユーザー指定）
  ・「重要度」は廃止（列・絞り込み・色分けをすべて削除）
  ・並び順は日付順のみ（月次集計行はその月の先頭）
  ・出力列は 日付 / 昼夜 / No / ルール / 該当箇所 / 理由 / 修正提案
  ・修正提案は商材名ではなくレシピ（メニュー）名

■ 必要なファイル
  必須：メニューワークブック（新構成ルール確認.xlsx など）
  推奨：食材CSV（7月昼.csv / 7月夜.csv / 8月昼.csv ...）※商品ID単位の判定に必須級
  推奨：野菜マスタ_テンプレート.xlsx（No.9 / No.17）
  推奨：調味料.csv（No.8 / No.19）
  推奨：食材データ.xlsx（No.12 当日揚げ / No.21 禁止食材 / No.14 栄養価）
"""
import os
import re
import tempfile

import pandas as pd
import streamlit as st

# チェックエンジンのモジュール名はリポジトリによって menu_checker.py / menu_check_core.py の
# どちらの場合もあるため、両方を試す（どちらか片方があれば動く）。
try:
    import menu_checker as mc
except ImportError:  # pragma: no cover
    import menu_check_core as mc

st.set_page_config(page_title='やどかり弁当 メニューチェック', page_icon='🍱', layout='wide')

# ----------------------------------------------------------------------------
# ユーティリティ
# ----------------------------------------------------------------------------
TMP_DIR = tempfile.mkdtemp(prefix='menu_check_')


def save_upload(uploaded, subdir=''):
    """アップロードされたファイルを一時ディレクトリに保存し、そのパスを返す。
    menu_check_core 側のローダーはすべて『ファイルパス』を受け取る設計のため、
    UploadedFile をそのまま渡さずここで実体化する。"""
    if uploaded is None:
        return None
    d = os.path.join(TMP_DIR, subdir) if subdir else TMP_DIR
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, uploaded.name)
    with open(path, 'wb') as f:
        f.write(uploaded.getbuffer())
    return path


def parse_month_slot(filename):
    """ファイル名から (月, '昼'|'夜') を推定する。'7月昼.csv' → (7, '昼')。
    推定できない場合は (None, None)。"""
    name = os.path.basename(str(filename))
    m = re.search(r'(\d{1,2})\s*月', name)
    month = int(m.group(1)) if m else None
    if '昼' in name:
        slot = '昼'
    elif '夜' in name or '夕' in name:
        slot = '夜'
    else:
        slot = None
    return month, slot


def date_sort_key(v):
    """menu_check_core と同じ並べ替えキー（画面側の並べ替え・月抽出用）"""
    return mc._date_sort_key(v)


# ----------------------------------------------------------------------------
# 入力UI
# ----------------------------------------------------------------------------
st.title('🍱 やどかり弁当 メニュー違反チェック')
st.caption('メニュー構成ルールに照らして違反候補を抽出し、レシピ単位の代替え案を提示します。')

with st.expander('① ファイルをアップロード', expanded=True):
    col_l, col_r = st.columns(2)

    with col_l:
        menu_xlsx = st.file_uploader(
            'メニューワークブック（必須）',
            type=['xlsx'],
            help='「N月使用食材」「N月昼夕…」「N月栄養価」シートを含むファイル（例：新構成ルール確認.xlsx）',
        )
        day_csvs = st.file_uploader(
            '食材CSV（昼・夜／複数選択可・推奨）',
            type=['csv'],
            accept_multiple_files=True,
            help='例：7月昼.csv / 7月夜.csv / 8月昼.csv / 8月夜.csv。'
                 'ファイル名から月と昼夜を自動判定します。',
        )

    with col_r:
        veg_master = st.file_uploader(
            '野菜マスタ_テンプレート.xlsx（No.9 / No.17）', type=['xlsx'])
        seasoning_csv = st.file_uploader(
            '調味料.csv（No.8 / No.19）', type=['csv'])
        shoku_data = st.file_uploader(
            '食材データ.xlsx（No.12 / No.21 / No.14）', type=['xlsx'])

use_ai = st.checkbox(
    'No.1（メイン/サブ商材の酷似判定）にAIを使う（Claude Haiku）',
    value=False,
    help='ANTHROPIC_API_KEY が Secrets または環境変数に設定されている場合のみ有効です。',
)

run = st.button('チェックを実行', type='primary')

# ----------------------------------------------------------------------------
# 実行
# ----------------------------------------------------------------------------
if run:
    if menu_xlsx is None:
        st.error('メニューワークブック（.xlsx）をアップロードしてください。')
        st.stop()

    with st.spinner('チェック中…'):
        xlsx_path = save_upload(menu_xlsx)
        veg_path = save_upload(veg_master)
        seasoning_path = save_upload(seasoning_csv)
        fried_path = save_upload(shoku_data)

        day_csv_paths, night_csv_paths, unknown_csvs = {}, {}, []
        for f in (day_csvs or []):
            month, slot = parse_month_slot(f.name)
            path = save_upload(f, 'daycsv')
            if month is None or slot is None:
                unknown_csvs.append(f.name)
                continue
            day_csv_paths[(month, slot)] = path
            if slot == '夜':
                night_csv_paths[month] = path

        ai_client = mc.get_anthropic_client() if use_ai else None
        if use_ai and ai_client is None:
            st.warning('ANTHROPIC_API_KEY が見つからないため、AI判定なしで実行します。')

        try:
            combined, summary = mc.run_all_checks(
                xlsx_path,
                night_csv_paths=night_csv_paths or None,
                day_csv_paths=day_csv_paths or None,
                veg_master_path=veg_path,
                seasoning_csv_path=seasoning_path,
                fried_master_path=fried_path,
                ai_client=ai_client,
            )
        except Exception as e:  # noqa: BLE001
            st.error(f'チェック中にエラーが発生しました：{e}')
            st.stop()

        out_path = os.path.join(TMP_DIR, '弁当メニューチェック_代替案付き.xlsx')
        mc.write_report(combined, summary['n_days'], out_path)

    st.session_state['result'] = {
        'combined': combined,
        'summary': summary,
        'out_path': out_path,
        'unknown_csvs': unknown_csvs,
    }

# ----------------------------------------------------------------------------
# 結果表示
# ----------------------------------------------------------------------------
result = st.session_state.get('result')
if result:
    combined = result['combined']
    summary = result['summary']
    out_path = result['out_path']

    months = '、'.join(f'{m}月' for m in summary['months'])
    st.success(f'対象月: {months} / 検査日数: {summary["n_days"]}日 / 違反候補: {summary["total"]}件')

    if result['unknown_csvs']:
        st.warning(
            '次のCSVは月と昼夜を判定できなかったため使用していません：'
            + '、'.join(result['unknown_csvs'])
            + '（ファイル名を「7月昼.csv」のような形式にしてください）'
        )

    if summary['warnings']:
        with st.expander(f'注記・スキップした判定（{len(summary["warnings"])}件）'):
            for w in summary['warnings']:
                st.write('・' + w)

    # ---- ルール別の件数カード ----
    by_rule = summary['by_rule']
    if by_rule:
        st.subheader('ルール別の検出件数')
        cards = [('総違反候補', summary['total'])]
        cards += [(f'No.{no}', cnt) for no, cnt in sorted(by_rule.items())]
        per_row = 8
        for i in range(0, len(cards), per_row):
            cols = st.columns(per_row)
            for col, (label, value) in zip(cols, cards[i:i + per_row]):
                help_txt = mc.RULE_NAME_JP.get(
                    int(label[3:]) if label.startswith('No.') else -1, None)
                col.metric(label, value, help=help_txt)

    st.subheader('違反候補リスト')

    if not len(combined):
        st.info('違反候補は見つかりませんでした。')
    else:
        view = combined.copy()

        # ---- 絞り込み（重要度は廃止。No と 昼夜 と 月 のみ）----
        f1, f2, f3 = st.columns([2, 1, 1])
        with f1:
            rule_opts = sorted(view['No'].unique())
            sel_rules = st.multiselect(
                'ルール（No.）で絞り込み',
                rule_opts,
                default=rule_opts,
                format_func=lambda n: f'No.{n} {mc.RULE_NAME_JP.get(n, "")}',
            )
        with f2:
            slot_opts = [s for s in ['昼', '夜', '昼夜', '-'] if s in set(view['昼夜'])]
            sel_slots = st.multiselect('昼夜', slot_opts, default=slot_opts)
        with f3:
            month_opts = sorted({date_sort_key(v)[0] for v in view['日付']})
            sel_months = st.multiselect(
                '月', month_opts, default=month_opts, format_func=lambda m: f'{m}月')

        view = view[view['No'].isin(sel_rules)]
        view = view[view['昼夜'].isin(sel_slots)]
        view = view[[date_sort_key(v)[0] in sel_months for v in view['日付']]]

        # ---- 並びは日付順のみ（同日内は 昼 → 昼夜 → 夜 → No）----
        if len(view):
            slot_order = {'昼': 0, '昼夜': 1, '夜': 2, '-': 3}
            view = view.assign(
                _k=view['日付'].map(date_sort_key),
                _s=view['昼夜'].map(lambda x: slot_order.get(x, 9)),
            ).sort_values(['_k', '_s', 'No']).drop(columns=['_k', '_s']).reset_index(drop=True)

        st.caption(f'{len(view)}件を表示中（全{len(combined)}件）')
        st.dataframe(
            view,
            use_container_width=True,
            hide_index=True,
            column_config={
                '日付': st.column_config.TextColumn('日付', width='small'),
                '昼夜': st.column_config.TextColumn('昼夜', width='small'),
                'No': st.column_config.NumberColumn('No', width='small', format='%d'),
                'ルール': st.column_config.TextColumn('ルール', width='medium'),
                '該当箇所': st.column_config.TextColumn('該当箇所', width='large'),
                '理由': st.column_config.TextColumn('理由', width='medium'),
                '修正提案': st.column_config.TextColumn('修正提案（代替えメニュー）', width='large'),
            },
        )

    # ---- ダウンロード ----
    with open(out_path, 'rb') as f:
        st.download_button(
            'Excelレポートをダウンロード（違反候補リスト／サマリー／チェックフロー）',
            data=f.read(),
            file_name='弁当メニューチェック_代替案付き.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    with st.expander('このチェックについて'):
        st.markdown(
            """
- 判定は可能な限りキーワード推測ではなく、**実際のマスタデータ（商品ID / レシピID）** で行っています。
- **重要度は付けていません。** 結果は日付順に並びます（月次集計の行はその月の先頭）。
- **修正提案は商材名ではなくレシピ（メニュー）名**で、その日時点で最も長く使われていないものを選びます。
- 対象外のルール：No.2（見た目酷似）、No.23（食べにくさ）はチェック対象から除外中。
  No.13 / No.16 / No.32〜35 / No.37 は未対応です。
- 食材CSV（昼・夜）を渡さないと、商品ID単位で判定するルール（No.1 / 10 / 12 / 18 / 19 / 20 / 21 / 30）が
  スキップされます。
            """
        )
else:
    st.info('ファイルをアップロードして「チェックを実行」を押してください。')
