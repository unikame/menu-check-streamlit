#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高齢者向け弁当 メニュー違反チェック（Streamlit画面）

menu_checker.run_all_checks() を呼び出して結果を表示するだけの薄いUI層。
判定ロジックは一切ここに書かない（ルールの追加・修正は menu_checker.py 側で行う）。

■ 画面構成
  サイドバー：マスタの接続状態 / ファイルのアップロード / 実行ボタン
  メイン    ：KPI（対象月・検査日数・違反候補数）＋ タブ（違反一覧 / ルール別 / 注記）

■ 毎回アップロードするファイル
  必須：メニューワークブック（新構成ルール確認.xlsx など）
  推奨：食材CSV（7月昼.csv / 7月夜.csv …）※CSVがある月だけが判定対象になる
  マスタ3種はGoogleスプレッドシートから自動取得（MASTER_SHEET_URLS）
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
#   ・いずれも共有設定を「リンクを知っている全員」→「閲覧者」にしてください。
#   ・空文字にすると、そのマスタは画面からの手動アップロードになります。
# ============================================================================
MASTER_SHEET_URLS = {
    # 野菜マスタ_テンプレート（No.9 色の2日連続 / No.17 赤黄緑）
    'veg': 'https://docs.google.com/spreadsheets/d/11taIDQoH-ibT1pJ9ewUkBK7IcAEWhThmX4kEPVD-vrY/edit?gid=519185715#gid=519185715',
    # 調味料（No.8 食材と調味料の被り / No.19 同じ調味料のみの味付け）
    'seasoning': 'https://docs.google.com/spreadsheets/d/1B_N6oLULtCHT_s9Ob6HkQ6Y-6IlU0i2wJ6RBR_Ukufs/edit?gid=952926231#gid=952926231',
    # 食材データ（No.12 当日揚げ / No.21 禁止食材 / No.14 レシピ別栄養価）
    # 複数タブを使うため、ブック全体をxlsxとして取得します（gidは不要）。
    'shoku': 'https://docs.google.com/spreadsheets/d/1bPju1fjNDDV59eCFrzStjGFoM0RhVzzrPk_MMLu20v0/edit?gid=1718859394#gid=1718859394',
}

MASTER_META = {
    'veg':       ('野菜マスタ', 'csv',  '野菜マスタ_テンプレート.csv', 'No.9 / No.17'),
    'seasoning': ('調味料',     'csv',  '調味料.csv',                  'No.8 / No.19'),
    'shoku':     ('食材データ', 'xlsx', '食材データ.xlsx',             'No.12 / No.21 / No.14'),
}

ACCENT = '#B4622F'

st.set_page_config(page_title='メニュー違反チェック', page_icon='🍱', layout='wide')

st.markdown(
    """
    <style>
      /* 余白を詰めて情報密度を上げる */
      .block-container { padding-top: 2.2rem; padding-bottom: 2rem; max-width: 1500px; }
      /* 見出し */
      h1 { font-size: 1.45rem !important; font-weight: 700 !important;
           letter-spacing: .01em; margin-bottom: .1rem !important; }
      h2 { font-size: 1.05rem !important; font-weight: 600 !important; }
      h3 { font-size: .92rem !important; font-weight: 600 !important; }
      .app-sub { color:#6B7280; font-size:.8rem; margin:0 0 1.1rem 0; }
      /* KPI カード */
      div[data-testid="stMetric"] {
        background:#FFFFFF; border:1px solid #E5E7EB; border-radius:8px;
        padding:.6rem .85rem;
      }
      div[data-testid="stMetricLabel"] p {
        font-size:.72rem !important; color:#6B7280 !important; font-weight:500 !important;
      }
      div[data-testid="stMetricValue"] {
        font-size:1.45rem !important; font-weight:650 !important; color:#111827 !important;
      }
      /* タブ */
      button[data-baseweb="tab"] { font-size:.85rem !important; padding:.35rem .1rem !important; }
      /* 表・入力まわりの文字 */
      div[data-testid="stDataFrame"] { font-size:.82rem; }
      section[data-testid="stSidebar"] { font-size:.85rem; }
      section[data-testid="stSidebar"] h2 { font-size:.95rem !important; }
      .stCaption, div[data-testid="stCaptionContainer"] p { font-size:.75rem !important; }
      /* マスタ接続状態の行 */
      .ms-ok, .ms-ng { font-size:.78rem; padding:.18rem 0; }
      .ms-ok { color:#15803D; } .ms-ng { color:#B45309; }
    </style>
    """,
    unsafe_allow_html=True,
)

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
    """共有URLをダウンロード用URLに変換する（スプレッドシート／ドライブのファイル）。"""
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
        return (f'https://docs.google.com/spreadsheets/d/{sheet_id}'
                f'/export?format=csv&gid={g.group(1) if g else "0"}')
    return f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'


def gviz_csv_url(url):
    """1タブをCSVで取得する代替エンドポイント（/export が空を返す場合のフォールバック）。"""
    m = _SHEET_ID_RE.search(str(url))
    if not m:
        raise ValueError('GoogleスプレッドシートのURLではありません')
    g = _GID_RE.search(str(url))
    return (f'https://docs.google.com/spreadsheets/d/{m.group(1)}'
            f'/gviz/tq?tqx=out:csv&gid={g.group(1) if g else "0"}')


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


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_master_bytes(url, fmt):
    """マスタの中身をダウンロードして bytes で返す（1時間キャッシュ）。
    /export が空を返す場合は gviz、ドライブのファイルは別ドメインにフォールバックする。"""
    try:
        return _download(gsheet_export_url(url, fmt))
    except (PermissionError, ValueError, urllib.error.HTTPError) as first_err:
        if fmt == 'csv' and _SHEET_ID_RE.search(str(url)):
            try:
                return _download(gviz_csv_url(url))
            except Exception:  # noqa: BLE001
                pass
        m = _DRIVE_FILE_RE.search(str(url))
        if m:
            try:
                return _download(f'https://drive.usercontent.google.com/download'
                                 f'?id={m.group(1)}&export=download')
            except Exception:  # noqa: BLE001
                pass
        raise PermissionError(
            f'読み込めませんでした（{first_err}）。共有設定を'
            '「リンクを知っている全員（閲覧者）」にしてください。') from first_err


def resolve_master(key):
    """マスタ1つ分を解決して (パス, 状態メッセージ, 成否) を返す。"""
    label, fmt, filename, rules = MASTER_META[key]
    url = (MASTER_SHEET_URLS.get(key) or '').strip()
    if not url:
        return None, f'{label}：URL未設定', False
    try:
        data = fetch_master_bytes(url, fmt)
    except PermissionError as e:
        return None, f'{label}：{e}', False
    except (urllib.error.URLError, ValueError, TimeoutError) as e:
        return None, f'{label}：取得に失敗（{e}）', False
    path = os.path.join(TMP_DIR, filename)
    if fmt == 'csv' and data[:2] == b'PK':
        # 実体がxlsxでも読めるようCSVに変換して保存する
        try:
            pd.read_excel(io.BytesIO(data)).to_csv(path, index=False)
        except Exception as e:  # noqa: BLE001
            return None, f'{label}：Excel形式の変換に失敗（{e}）', False
    else:
        with open(path, 'wb') as f:
            f.write(data)
    return path, f'{label}（{rules}）', True


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
    return mc._date_sort_key(v)


# ----------------------------------------------------------------------------
# サイドバー（入力）
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown('## 入力')

    menu_xlsx = st.file_uploader('メニューワークブック（必須）', type=['xlsx'])
    day_csvs = st.file_uploader(
        '食材CSV（昼・夜／複数可）', type=['csv'], accept_multiple_files=True,
        help='例：7月昼.csv / 7月夜.csv。ファイル名から月と昼夜を判定します。'
             'アップロードした月だけが判定対象になります。')

    run = st.button('チェックを実行', type='primary', use_container_width=True)

    st.divider()
    st.markdown('## マスタ')

    master_paths, master_ok = {}, {}
    for key in MASTER_META:
        path, msg, ok = resolve_master(key)
        master_paths[key], master_ok[key] = path, ok
        cls = 'ms-ok' if ok else 'ms-ng'
        icon = '●' if ok else '▲'
        st.markdown(f'<div class="{cls}">{icon} {msg}</div>', unsafe_allow_html=True)

    if st.button('マスタを再読込', use_container_width=True):
        fetch_master_bytes.clear()
        st.rerun()

    if not all(master_ok.values()):
        with st.expander('取得できないマスタを手動で指定'):
            manual = {
                'veg': st.file_uploader('野菜マスタ', type=['xlsx', 'csv'],
                                        disabled=master_ok['veg']),
                'seasoning': st.file_uploader('調味料', type=['csv'],
                                              disabled=master_ok['seasoning']),
                'shoku': st.file_uploader('食材データ', type=['xlsx'],
                                          disabled=master_ok['shoku']),
            }
            for key, up in manual.items():
                if up is not None:
                    master_paths[key] = save_upload(up, 'master')

    st.divider()
    use_ai = st.checkbox('No.1の酷似判定にAIを使う', value=False,
                         help='ANTHROPIC_API_KEY が設定されている場合のみ有効です。')

# ----------------------------------------------------------------------------
# ヘッダ
# ----------------------------------------------------------------------------
st.markdown('# 高齢者向け弁当 メニュー違反チェック')
st.markdown(
    '<p class="app-sub">メニュー構成ルールに照らして違反候補を抽出し、'
    'レシピ単位の代替え案を提示します。</p>',
    unsafe_allow_html=True)

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
        'combined': combined, 'summary': summary,
        'out_path': out_path, 'unknown_csvs': unknown_csvs,
    }

# ----------------------------------------------------------------------------
# 結果表示
# ----------------------------------------------------------------------------
result = st.session_state.get('result')

if not result:
    st.info('左のサイドバーでファイルを指定し、「チェックを実行」を押してください。')
    st.stop()

combined = result['combined']
summary = result['summary']
out_path = result['out_path']
by_rule = summary['by_rule']

# ---- KPI ----
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric('違反候補', f'{summary["total"]:,}')
k2.metric('対象月', '、'.join(f'{m}月' for m in summary['months']) or '—')
k3.metric('検査日数', f'{summary["n_days"]}日')
k4.metric('該当ルール数', f'{len(by_rule)}件')
k5.metric('注記', f'{len(summary["warnings"])}件')

if result['unknown_csvs']:
    st.warning('月と昼夜を判定できなかったCSV：' + '、'.join(result['unknown_csvs'])
               + '（「7月昼.csv」のような名前にしてください）')

st.write('')
tab_list, tab_rule, tab_note = st.tabs(['違反一覧', 'ルール別', '注記・出力'])

# ---- 違反一覧 ----
with tab_list:
    if not len(combined):
        st.success('違反候補は見つかりませんでした。')
    else:
        view = combined.copy()
        f1, f2, f3 = st.columns([3, 1, 1])
        with f1:
            rule_opts = sorted(view['No'].unique())
            sel_rules = st.multiselect(
                'ルール', rule_opts, default=rule_opts,
                format_func=lambda n: f'No.{n} {mc.RULE_NAME_JP.get(n, "")}')
        with f2:
            slot_opts = [s for s in ['昼', '夜', '昼夜', '-'] if s in set(view['昼夜'])]
            sel_slots = st.multiselect('昼夜', slot_opts, default=slot_opts)
        with f3:
            month_opts = sorted({date_sort_key(v)[0] for v in view['日付']})
            sel_months = st.multiselect('月', month_opts, default=month_opts,
                                        format_func=lambda m: f'{m}月')

        view = view[view['No'].isin(sel_rules)]
        view = view[view['昼夜'].isin(sel_slots)]
        view = view[[date_sort_key(v)[0] in sel_months for v in view['日付']]]

        if len(view):
            slot_order = {'昼': 0, '昼夜': 1, '夜': 2, '-': 3}
            view = view.assign(
                _k=view['日付'].map(date_sort_key),
                _s=view['昼夜'].map(lambda x: slot_order.get(x, 9)),
            ).sort_values(['_k', '_s', 'No']).drop(columns=['_k', '_s']).reset_index(drop=True)

        st.caption(f'{len(view):,} / {len(combined):,} 件（日付順）')
        st.dataframe(
            view, use_container_width=True, hide_index=True, height=560,
            column_config={
                '日付': st.column_config.TextColumn('日付', width='small'),
                '昼夜': st.column_config.TextColumn('昼夜', width='small'),
                'No': st.column_config.NumberColumn('No', width='small', format='%d'),
                'ルール': st.column_config.TextColumn('ルール', width='medium'),
                '該当箇所': st.column_config.TextColumn('該当箇所', width='large'),
                '理由': st.column_config.TextColumn('理由', width='medium'),
                '修正提案': st.column_config.TextColumn('代替えメニュー', width='large'),
            })

# ---- ルール別 ----
with tab_rule:
    if not by_rule:
        st.info('検出はありません。')
    else:
        rows = [{'No': f'No.{no}', 'ルール': mc.RULE_NAME_JP.get(no, ''), '件数': cnt}
                for no, cnt in sorted(by_rule.items(), key=lambda x: -x[1])]
        rdf = pd.DataFrame(rows)
        st.dataframe(
            rdf, use_container_width=True, hide_index=True,
            height=min(560, 38 * len(rdf) + 40),
            column_config={
                'No': st.column_config.TextColumn('No', width='small'),
                'ルール': st.column_config.TextColumn('ルール', width='medium'),
                '件数': st.column_config.ProgressColumn(
                    '件数', format='%d', min_value=0,
                    max_value=int(max(by_rule.values())), width='large'),
            })
        st.caption('件数の多い順。判定内容の詳細は menu_checker.py のルール定義を参照してください。')

# ---- 注記・出力 ----
with tab_note:
    c_dl, _ = st.columns([1, 2])
    with c_dl:
        with open(out_path, 'rb') as f:
            st.download_button('Excelレポートをダウンロード', data=f.read(),
                               file_name='弁当メニューチェック_代替案付き.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.'
                                    'spreadsheetml.sheet',
                               use_container_width=True)

    if summary['warnings']:
        st.markdown('##### 注記・スキップした判定')
        for w in summary['warnings']:
            st.markdown(f'- {w}')
    else:
        st.caption('注記はありません。')

    st.markdown('##### このチェックについて')
    st.markdown(
        '- 判定は可能な限りキーワード推測ではなく、実際のマスタデータ（商品ID / レシピID）で行っています。\n'
        '- 重要度は付けていません。結果は日付順に並びます（月次集計の行はその月の先頭）。\n'
        '- 代替え案は商材名ではなくレシピ（メニュー）名で、その日時点で最も長く使われていないものを選びます。\n'
        '- 食材CSVをアップロードした月だけが判定対象です。\n'
        '- マスタは1時間キャッシュします。更新直後はサイドバーの「マスタを再読込」を押してください。\n'
        '- 対象外：No.2（見た目酷似）／No.23（食べにくさ）。未対応：No.13 / No.16 / No.32〜35 / No.37。')
