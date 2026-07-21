# やどかり弁当 メニューチェック（Streamlit版）

37項目のメニュー構成ルールに照らして、月次メニューを自動チェックするアプリです。
[Streamlit Community Cloud](https://streamlit.io/cloud) での公開を想定しています。

## 公開までの手順

### 1. このフォルダの中身をGitHubリポジトリにpushする

既存の `new-bento-checker` などのリポジトリを使うか、新しいリポジトリを作成してください。

```bash
cd menu_check_streamlit
git init
git add .
git commit -m "menu check streamlit app"
git branch -M main
git remote add origin https://github.com/unikame/（リポジトリ名）.git
git push -u origin main
```

フォルダ構成（このまま丸ごとpushしてください）:

```
menu_check_streamlit/
├── streamlit_app.py       ← Streamlit Community Cloudが自動検出するエントリーポイント
├── menu_checker.py        ← チェックロジック本体
├── requirements.txt
├── scripts/
│   └── check_menu.py      ← 判定エンジン（bento-menu-checkスキル由来）
└── assets/
    └── master_ingredients.csv  ← 食材グループのマスタ
```

### 2. Streamlit Community Cloudでデプロイする

1. [https://share.streamlit.io/](https://share.streamlit.io/) にアクセスし、GitHubアカウントでログイン
2. 「New app」をクリック
3. Repository: 上でpushしたリポジトリを選択
4. Branch: `main`
5. Main file path: `streamlit_app.py`
6. 「Deploy」をクリック

数分でビルドが完了し、`https://（アプリ名）.streamlit.app/` のようなURLが発行されます
（`condition-check.streamlit.app` と同じ仕組みです）。認証は特に設定していないため、
URLを知っている人は誰でもアクセスできます。

### 3. 使い方

1. メニューワークブック(.xlsx)をアップロード
   - 月ごとに `{月}月使用食材` シートと `{月}月昼夕...` シートを含めてください
     （例: `7月使用食材`, `7月昼夕比較`, `8月使用食材`, `8月昼夕`）
2. 夜（夕）の食材データがワークブック内に無い月がある場合は、CSVを追加でアップロード
   （ファイル名に「7月」のように月を含めること）
3. 「チェックを実行」を押すと、違反候補が画面に表示され、Excelでもダウンロードできます

## 対象外のルール

- **No.8・17・19・20**: 特定月固有の「調理法/調味料タグ」列に依存するため、この汎用版では判定対象外
- **No.9・10・13・16・18・33・35**: 見た目類似・盛付工数・重量・ソース分類などの追加マスタデータが無いため未実装
- **No.27**: 魚弁当専用5商材のみ判定（かぼちゃ煮物・豆乳等はFD月次メニューの色分けが不安定なため対象外）

## ローカルで動作確認したい場合

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

`http://localhost:8501/` で確認できます。

## 更新方法

コードを直したら、GitHubにpushするだけで自動的にStreamlit Community Cloud側にも反映されます
（再デプロイ操作は不要です）。
