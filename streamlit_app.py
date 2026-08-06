#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
やどかり弁当 メニュー違反チェック（Streamlit画面）

menu_checker.run_all_checks() を呼び出して結果を表示するだけの薄いUI層。
判定ロジックは一切ここに書かない（ルールの追加・修正は menu_checker.py 側で行う）。

■ この版での変更点（ユーザー指定）
  ・「重要度」は廃止（列・絞り込み・色分けをすべて削除）
  ・並び順は日付順のみ（月次集計行はその月の先頭）
  ・修正提案は商材名ではなくレシピ（メニュー）名
  ・3つのマスタ（野菜マスタ／調味料／食材データ）は毎回アップロードせず、
    Googleスプレッドシートから自動取得する（下の MASTER_SHEET_URLS を設定）

■ 毎回アップロードするファイル
  必須：メニューワークブック（新構成ルール確認.xlsx など）
  推奨：食材CSV（7月昼.csv / 7月夜.csv …）※CSVがある月だけが判定対象になる
"""
import io
import os
import re
import tempfile
import urllib.error
import urllib.request

import pandas as pd
import streamlit as st

# チェックエンジンのモジュール名はリポジトリによって menu_checker.py / menu_check_core.py の
# どちらの場合もあるため、両方を試す（どちらか片方があれば動く）。
try:
    import menu_checker as mc
except ImportError:  # pragma: no cover
    import menu_check_core as mc

# ============================================================================
# ★ 3つのマスタの参照先URL ★
#   ・Googleスプレッドシートの編集画面URL（.../edit#gid=0）でも、
#     Googleドライブのファイル共有URL（.../file/d/xxx/view）でもそのまま貼れます。
#     内部でダウンロード用URLに変換します。
#   ・いずれも共有設定を「リンクを知っている全員」→「閲覧者」にしてください。
#     （非公開のままだとログインページが返ってきて読み込めません）
#   ・空文字にすると、そのマスタは画面からの手動アップロードになります。
# ============================================================================
MASTER_SHEET_URLS = {
    # 野菜マスタ_テンプレート（No.9 色の2日連続 / No.17 赤黄緑）
    # 複数タブを使う可能性があるため、ブック全体をxlsxとして取得します（gidは不要）。
    'veg': 'https://docs.google.com/spreadsheets/d/11taIDQoH-ibT1pJ9ewUkBK7IcAEWhThmX4kEPVD-vrY/edit?gid=519185715#gid=519185715',
    # 調味料（No.8 食材と調味料の被り / No.19 同じ調味料のみの味付け）
    # URLのgidで指定されたタブをCSVとして取得します。
    'seasoning': 'https://docs.google.com/spreadsheets/d/1B_N6oLULtCHT_s9Ob6HkQ6Y-6IlU0i2wJ6RBR_Ukufs/edit?gid=952926231#gid=952926231',
    # 食材データ（No.12 当日揚げ / No.21 禁止食材 / No.14 レシピ別栄養価）
    # 「調理法（当日揚げ）」「禁止食材・調味料該当」など複数タブを使うため、
    # ブック全体をxlsxとして取得します（gidは不要）。
    'shoku': 'https://docs.google.com/spreadsheets/d/1bPju1fjNDDV59eCFrzStjGFoM0RhVzzrPk_MMLu20v0/edit?gid=1718859394#gid=1718859394',
}

MASTER_META = {
    # 野菜マスタは、URLのgidで指定されたタブをCSVとして取得する
    # （タブが複数あるため、ブック全体のxlsxだと別タブを読んでしまう恐れがある）
    'veg':       ('野菜マスタ', 'csv',  '野菜マスタ_テンプレート.csv', 'No.9 / No.17'),
    'seasoning': ('調味料',     'csv',  '調味料.csv',                  'No.8 / No.19'),
    'shoku':     ('食材データ', 'xlsx', '食材データ.xlsx',             'No.12 / No.21 / No.14'),
}

st.set_page_config(page_title='やどかり弁当 メニューチェック', page_icon='🍱', layout='wide')

TMP_DIR = tempfile.mkdtemp(prefix='menu_check_')

_SHEET_ID_RE = re.compile(r'/spreadsheets/d/([A-Za-z0-9_-]+)')
_DRIVE_FILE_RE = re.compile(r'drive\.google\.com/file/d/([A-Za-z0-9_-]+)')
_GID_RE = re.compile(r'[#&?]gid=(\d+)')


# ----------------------------------------------------------------------------
# ユーティリティ
# ----------------------------------------------------------------------------
def save_upload(uploaded, subdir=''):
    """アップロードされたファイルを一時ディレクトリに保存し、そのパスを返す。
    menu_checker 側のローダーはすべて『ファイルパス』を受け取る設計のため、
    UploadedFile をそのまま渡さずここで実体化する。"""
    if uploaded is None:
        return None
    d = os.path.join(TMP_DIR, subdir) if subdir else TMP_DIR
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, uploaded.name)
    with open(path, 'wb') as f:
        f.write(uploaded.getbuffer())
    return path


def gsheet_export_url(url, fmt='xlsx'):
    """共有URLを、ダウンロード用のURLに変換する。次の2種類に対応する。
    ・Googleスプレッドシート（.../spreadsheets/d/<id>/edit#gid=123）
        fmt='xlsx' はブック全体（全タブ）、fmt='csv' は gid で指定した1タブのみ。
    ・Googleドライブ上のファイル（.../file/d/<id>/view）
        アップロードされたCSV/XLSXをそのままダウンロードする。
    すでにexport形式のURLを渡した場合もそのまま扱える。"""
    s = str(url)
    m = _DRIVE_FILE_RE.search(s)
    if m:
        return f'https://drive.google.com/uc?export=download&id={m.group(1)}'
    m = _SHEET_ID_RE.search(s)
    if not m:
        raise ValueError('GoogleスプレッドシートまたはGoogleドライブのURLではありません')
    sheet_id = m.group(1)
    if fmt == 'csv':
        g = _GID_RE.search(s)
        gid = g.group(1) if g else '0'
        return f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}'
    return f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'


def _download(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=60) as res:
        data = res.read()
    head = data[:200].lstrip().lower()
    if head.startswith(b'<!doctype html') or head.startswith(b'<html'):
        raise PermissionError('HTMLが返ってきました（共有設定が不十分な可能性があります）')
    if not data.strip():
        raise ValueError('中身が空でした')
    return data


def gviz_csv_url(url):
    """スプレッドシートの1タブをCSVで取得する代替エンドポイント。
    /export?format=csv が空を返すことがあるため、そのフォールバックに使う。"""
    m = _SHEET_ID_RE.search(str(url))
    if not m:
        raise ValueError('GoogleスプレッドシートのURLではありません')
    g = _GID_RE.search(str(url))
    gid = g.group(1) if g else '0'
    return (f'https://docs.google.com/spreadsheets/d/{m.group(1)}'
            f'/gviz/tq?tqx=out:csv&gid={gid}')


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_master_bytes(url, fmt):
    """マスタの中身をダウンロードして bytes で返す（1時間キャッシュ）。
    CSV取得は /export が空を返すことがあるため、gviz エンドポイントにフォールバックする。
    共有設定が不十分な場合、Googleはファイルではなくログイン用のHTMLを返すので、
    それを検出して分かりやすいエラーにする。"""
    try:
        return _download(gsheet_export_url(url, fmt))
    except (PermissionError, ValueError, urllib.error.HTTPError) as first_err:
        # スプレッドシートのCSV：/export が空を返すことがあるので gviz を試す
        if fmt == 'csv' and _SHEET_ID_RE.search(str(url)):
            try:
                return _download(gviz_csv_url(url))
            except Exception:  # noqa: BLE001
                pass
        # ドライブ上のファイル：uc?export=download が確認画面を返すことがあるので
        # 新しい配信ドメインを試す
        m = _DRIVE_FILE_RE.search(str(url))
        if m:
            try:
                return _download(
                    f'https://drive.usercontent.google.com/download'
                    f'?id={m.group(1)}&export=download')
            except Exception:  # noqa: BLE001
                pass
        raise PermissionError(
            f'読み込めませんでした（{first_err}）。共有設定を'
            '「リンクを知っている全員（閲覧者）」にしてください。') from first_err


def resolve_master(key):
    """マスタ1つ分を解決して (パス, 状態メッセージ, 成否) を返す。
    URLが設定されていればスプレッドシートから取得し、一時ファイルに書き出す。"""
    label, fmt, filename, rules = MASTER_META[key]
    url = (MASTER_SHEET_URLS.get(key) or '').strip()
    if not url:
        return None, f'{label}：URL未設定（手動アップロードしてください）', False
    try:
        data = fetch_master_bytes(url, fmt)
    except PermissionError as e:
        return None, f'{label}：{e}', False
    except (urllib.error.URLError, ValueError, TimeoutError) as e:
        return None, f'{label}：取得に失敗しました（{e}）', False
    path = os.path.join(TMP_DIR, filename)
    if fmt == 'csv' and data[:2] == b'PK':
        # ドライブ上の実体がxlsxだった場合でも読めるよう、CSVに変換して保存する
        # （調味料マスタのローダーは pd.read_csv で読む前提のため）
        try:
            pd.read_excel(io.BytesIO(data)).to_csv(path, index=False)
        except Exception as e:  # noqa: BLE001
            return None, f'{label}：Excel形式のため変換を試みましたが失敗しました（{e}）', False
    else:
        with open(path, 'wb') as f:
            f.write(data)
    return path, f'{label}：オンラインのマスタから取得済み（{rules}）', True


def parse_month_slot(filename):
    """ファイル名から (月, '昼'|'夜') を推定する。'7月昼.csv' → (7, '昼')。"""
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
    """menu_checker と同じ並べ替えキー（画面側の並べ替え・月抽出用）"""
    return mc._date_sort_key(v)


# ----------------------------------------------------------------------------
# 入力UI
# ----------------------------------------------------------------------------
st.title('🍱 やどかり弁当 メニュー違反チェック')
st.caption('メニュー構成ルールに照らして違反候補を抽出し、レシピ単位の代替え案を提示します。')

# ---- マスタ（スプレッドシート参照）----
master_paths, master_ok = {}, {}
with st.expander('① マスタ（スプレッドシート参照）', expanded=False):
    c1, c2 = st.columns([4, 1])
    with c2:
        if st.button('マスタを再読込'):
            fetch_master_bytes.clear()
            st.rerun()
    for key in MASTER_META:
        path, msg, ok = resolve_master(key)
        master_paths[key], master_ok[key] = path, ok
        (st.success if ok else st.warning)(msg)

    missing = [MASTER_META[k][0] for k in MASTER_META if not master_ok[k]]
    if missing:
        st.markdown('**取得できなかったマスタを手動でアップロード**')
        m1, m2, m3 = st.columns(3)
        manual = {
            'veg': m1.file_uploader('野菜マスタ_テンプレート.xlsx', type=['xlsx', 'csv'],
                                     disabled=master_ok['veg']),
            'seasoning': m2.file_uploader('調味料.csv', type=['csv'],
                                           disabled=master_ok['seasoning']),
            'shoku': m3.file_uploader('食材データ.xlsx', type=['xlsx'],
                                       disabled=master_ok['shoku']),
        }
        for key, up in manual.items():
            if up is not None:
                master_paths[key] = save_upload(up, 'master')

# ---- 毎回アップロードするファイル ----
with st.expander('② メニュー・食材データをアップロード', expanded=True):
    col_l, col_r = st.columns(2)
    with col_l:
        menu_xlsx = st.file_uploader(
            'メニューワークブック（必須）',
            type=['xlsx'],
            help='「N月使用食材」「N月昼夕…」「N月栄養価」シートを含むファイル（例：新構成ルール確認.xlsx）',
        )
    with col_r:
        day_csvs = st.file_uploader(
            '食材CSV（昼・夜／複数選択可・推奨）',
            type=['csv'],
            accept_multiple_files=True,
            help='例：7月昼.csv / 7月夜.csv。ファイル名から月と昼夜を自動判定します。'
                 'CSVをアップロードした月だけが判定対象になります。',
        )

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
                veg_master_path=master_paths.get('veg'),
                seasoning_csv_path=master_paths.get('seasoning'),
                fried_master_path=master_paths.get('shoku'),
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
- **食材CSVをアップロードした月だけが判定対象**になります。7月分だけ入れれば7月分だけ出力されます。
- マスタ（野菜／調味料／食材データ）はスプレッドシートから自動取得し、1時間キャッシュします。
  スプレッドシートを更新した直後は「マスタを再読込」を押してください。
- 対象外のルール：No.2（見た目酷似）、No.23（食べにくさ）。No.13 / No.16 / No.32〜35 / No.37 は未対応です。
            """
        )
else:
    st.info('ファイルをアップロードして「チェックを実行」を押してください。')
