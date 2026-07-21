#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
メニュー違反チェック（Streamlit版）
Streamlit Community Cloud (share.streamlit.io) での公開を想定。
"""
import io
import os
import re
import tempfile

import pandas as pd
import streamlit as st

import menu_checker as mc

st.set_page_config(page_title="やどかり弁当 メニューチェック", layout="wide")

SEVERITY_ORDER = {"高": 0, "中": 1, "低": 2}


def severity_rank(v):
    v = str(v)
    for k, r in SEVERITY_ORDER.items():
        if v.startswith(k):
            return r
    return 3


st.title("やどかり弁当 メニュー違反チェック")
st.caption("37項目のメニュー構成ルールに照らして自動チェックします。")

with st.expander("アップロードするファイルの形式", expanded=False):
    st.markdown(
        """
        - メニューワークブック(.xlsx)には、月ごとに **`{月}月使用食材`** シートと
          **`{月}月昼夕...`**（昼夜のメニュー名一覧）シートを含めてください。
          （例: `7月使用食材`, `7月昼夕比較`, `8月使用食材`, `8月昼夕`）
        - 夜（夕）の食材データがワークブック内に無い場合は、下の欄で月ごとにCSVを追加できます
          （ファイル名に「7月」のように月を含めてください）。
        - **対象外のルール**: No.8・17・19・20 は特定月固有の「調理法/調味料タグ」列に依存するため、
          この汎用版では判定対象外です。No.9・10・13・16・18・33・35 は追加マスタデータが無く未実装です。
        """
    )

main_file = st.file_uploader("① メニューワークブック（必須・.xlsx）", type=["xlsx"])
night_files = st.file_uploader(
    "② 夜（夕）食材CSV（任意・複数選択可、ファイル名に月を含めてください）",
    type=["csv"], accept_multiple_files=True,
)

run = st.button("チェックを実行", type="primary", disabled=(main_file is None))

if run and main_file is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        main_path = os.path.join(tmpdir, "main.xlsx")
        with open(main_path, "wb") as f:
            f.write(main_file.getbuffer())

        night_csv_paths = {}
        for f in night_files or []:
            m = re.search(r"(\d{1,2})\s*月", f.name)
            if not m:
                st.warning(f"「{f.name}」は月が判別できないためスキップしました（ファイル名に「7月」等を含めてください）")
                continue
            month = int(m.group(1))
            p = os.path.join(tmpdir, f"night_{month}.csv")
            with open(p, "wb") as out:
                out.write(f.getbuffer())
            night_csv_paths[month] = p

        try:
            with st.spinner("チェック中..."):
                combined, summary = mc.run_all_checks(main_path, night_csv_paths)
        except Exception as e:
            st.error(f"チェック中にエラーが発生しました: {e}")
            st.exception(e)
            st.stop()

        st.success(f"対象月: {', '.join(str(m) + '月' for m in summary['months'])} / 検査日数: {summary['n_days']}日")

        cols = st.columns(min(len(summary["by_rule"]) + 1, 8))
        cols[0].metric("総違反候補", summary["total"])
        for i, (no, cnt) in enumerate(summary["by_rule"].items()):
            cols[(i + 1) % len(cols)].metric(f"No.{no}", cnt)

        if summary["warnings"]:
            with st.expander("注記（データ不足等でスキップした項目）", expanded=True):
                for w in summary["warnings"]:
                    st.markdown(f"- {w}")

        if len(combined):
            rows = combined.copy()
            rows["sev_rank"] = rows["重要度"].apply(severity_rank)
            rows = rows.sort_values(["sev_rank", "No", "日付"]).drop(columns=["sev_rank"])
        else:
            rows = combined

        st.dataframe(rows, use_container_width=True, height=560)

        # Excelレポートをダウンロード用に生成
        report_path = os.path.join(tmpdir, "menu_check_result.xlsx")
        try:
            mc.write_report(combined, summary["n_days"], report_path)
        except Exception:
            combined.to_excel(report_path, index=False)
        with open(report_path, "rb") as f:
            st.download_button(
                "Excelレポートをダウンロード",
                data=f.read(),
                file_name="menu_check_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
elif main_file is None:
    st.info("メニューワークブック(.xlsx)をアップロードすると「チェックを実行」が押せるようになります。")
