# プロジェクト概要
- 名前：台湾市場向け 事前リサーチダッシュボード
- 目的：JR東日本 台湾事業開発チームが日本自治体向け提案を行うための事前調査ツール。台湾人旅行者の関心・トレンド・ペルソナを自動収集・AI分析し、PowerPoint提案資料作成を支援する。
- 主な技術：Python, Streamlit, Google Trends (pytrends), BeautifulSoup, jieba, WordCloud, matplotlib, openpyxl, feedparser
- バージョン：v3.0（2026-03）

---

# Claude への基本ルール
- コードを変更する前に必ず理由を説明すること
- 大きな変更は一度に行わず段階的に進めること
- 日本語でコメント・説明を書くこと
- 不明点があれば推測せず確認すること

---

# 自動記録ルール（最重要・必ず守ること）

## エラー解決時（必須）
エラーを解決したら、必ず「トラブル記録」セクションに以下の形式で追記すること：
- 日付
- エラー内容（エラーメッセージ含む）
- 原因
- 解決策
- 再発防止のポイント

## 技術的な判断をしたとき（必須）
ライブラリの選定、設計方針、実装アプローチを決めたら「設計判断ログ」に追記すること：
- 日付
- 何を決めたか
- 選んだ理由
- 却下した代替案

## やってはいけないことが判明したとき（必須）
失敗したアプローチや使ってはいけない方法が判明したら「禁止事項」に追記すること

## セッション終了時（必須）
作業終了時に「作業ログ」に今回やったことの要約を追記すること

---

# 禁止事項

- **Dcard API の使用**：403エラーで完全ブロック。台湾フォーラムデータ取得には使えない。
- **背包客棧（backpackers.com.tw）のスクレイピング**：検索結果がJavaScriptレンダリングのため、requestsでは取得不可。ナビゲーションリンクのみ返ってくる。
- **`@st.cache_data` でGemini APIレスポンスをキャッシュ**：429エラー文字列もキャッシュされ、レート制限解除後もエラーが1時間返り続ける。成功結果のみ session_state に保存すること。
- **`@st.cache_resource` でGeminiモデルをプローブ**：429でNoneがキャッシュされ、アプリ再起動まで「モデルが見つからない」エラーが続く。
- **Gemini モデルをハードコード**：モデル名は変更・廃止されることがある。`/v1beta/models` APIで自動検出すること。
- **Tab内でGemini呼び出しを自動実行（ボタンなし）**：Streamlitは全タブを毎回レンダリングするため、どのボタンを押しても全タブのGemini呼び出しが発火する。必ずボタンとsession_stateで制御すること。
- **Facebook/Instagram APIでの生の声取得**：Meta APIは認証なしでのアクセスが不可。代替として Google News RSS（台湾向け）+ YouTube Data API v3 を使うこと。
- **Google News RSS の直接 ET.fromstring() パース**：Googleが返すRSSにBOMや不正XMLが含まれる場合がある。feedparser ライブラリ経由でパースする方が安全。

---

# トラブル記録

## 2025-02 Dcard 403エラー
- エラー：Dcard API が 403 を返す
- 原因：Dcard が外部APIアクセスをブロック
- 解決策：背包客棧に切り替え → さらにJSレンダリング問題で断念 → 外部リンクボタンパネルに変更
- 再発防止：台湾フォーラムは直接スクレイピング不可。外部リンク誘導で対応。

## 2025-02 背包客棧「no_posts」エラー
- エラー：scrape_backpacker が空リストを返す
- 原因：backpackers.com.tw の検索結果はJavaScriptレンダリング。requestsではHTMLスケルトン（ナビゲーションリンクのみ）しか取得できない
- 解決策：スクレイピングを完全廃止。背包客棧・Mobile01・PTT・Yahoo台湾への外部リンクボタンパネルに置き換え
- 再発防止：JSレンダリングサイトにはrequests+BeautifulSoupは使えない

## 2025-02 Gemini 429レート制限エラーが繰り返す
- エラー：429 TooManyRequests が連続発生
- 原因①：`@st.cache_data(ttl=3600)` が429エラー文字列をキャッシュ → レート制限解除後も1時間エラーが返り続ける
- 原因②：Tab1のGemini呼び出しがボタンなしで自動実行 → 他タブのボタン押下でも発火
- 原因③：2モデル×2リトライ = 1ボタンで最大4リクエスト発生
- 解決策：① cache_data を削除し session_state で成功結果のみキャッシュ ② 全Gemini呼び出しをボタントリガーに変更 ③ 1モデル・リトライなしに簡略化
- 再発防止：Gemini呼び出しは必ずボタン経由、エラーはsession_stateに保存しない

## 2025-02 Gemini 404モデルが見つからない
- エラー：gemini-2.0-flash および gemini-1.5-flash が 404
- 原因：APIキーで利用可能なモデル名がハードコードと異なる
- 解決策：`/v1beta/models` APIでモデル一覧を自動取得し、利用可能なflashモデルを動的に選択。結果はsession_stateにキャッシュ
- 再発防止：モデル名はハードコードせず自動検出する

---

# 設計判断ログ

## 2025-02 フォーラムデータ取得方式
- 決定：外部リンクボタンパネル（背包客棧・Mobile01・PTT・Yahoo台湾）
- 理由：JSレンダリングサイトへのスクレイピングは技術的に困難。ユーザーが直接検索する方が確実
- 却下した代替案：requests+BeautifulSoup（JSレンダリング不可）、Selenium/Playwright（環境依存・複雑）

## 2025-02 Gemini API ライブラリ選定
- 決定：google-generativeai ライブラリを使わず REST API 直接呼び出し
- 理由：ライブラリのバージョン依存・モデル名変更への追従が困難。requestsで直接呼べば柔軟に対応可能
- 却下した代替案：google-generativeai SDK（モデル探索ロジックが複雑になった）

## 2025-02 Geminiモデル選定方式
- 決定：`/v1beta/models` APIで動的検出、session_stateにキャッシュ
- 理由：モデル名は変更・廃止リスクがある。自動検出により将来的な名前変更に対応
- 却下した代替案：ハードコード（404エラーの原因）、`@st.cache_resource`（Noneキャッシュ問題）

## 2026-03 Gemini完全廃止・テンプレート分析移行（v3.0）
- 決定：Gemini API を完全廃止し、テンプレートベース分析（APIキー不要）に置き換え
- 理由：Gemini無料枠（15回/分）では実用上頻繁にレート制限に当たり、ユーザーが利用できない状態になっていた
- 却下した代替案：Groq API（無料だが別途APIキー取得が必要）、OpenAI API（有料）
- テンプレート分析の仕組み：Google Trendsデータ（ピーク月・関連KW）とPTTデータからルールベースでMarkdownレポートを生成

## 2026-03 JNTOデータ自動取得実装（v3.0）
- 決定：JNTOの統計ページをスクレイピングして最新Excelを自動DL・パースする
- 理由：CSVアップロード方式はユーザーの手間が大きく、データの鮮度も担保できない
- 却下した代替案：固定URLのハードコード（URL内の日付が更新ごとに変わるため）
- 実装のポイント：BeautifulSoupで.xlsxリンクを動的に検出 → requests でDL → pandas+openpyxlで台湾行を探してパース → 24時間キャッシュ

## 2026-03 JNTOパース手法をopenpyxl直接展開に変更（v3.0バグ修正）
- 決定：`pandas.read_excel` ではなく `openpyxl.load_workbook` + 手動 `unmerge_ws()` でExcelを解析
- 理由：JNTOのExcelは年ヘッダーが横方向に結合セルで格納されており、`read_excel` では先頭セルのみ値が入り残りがNaN → `ffill()` でも対応不可（そもそも年が文字列型で条件チェック失敗）
- 却下した代替案：pandas の `ffill()`（文字列型年ヘッダーに対応不可）、`pd.read_excel(header=...)` の調整（結合セルには根本対応にならない）
- 実装：unmerge後の全セル値を2Dリストに変換 → `_extract_year()` / `_extract_month()` で文字列/数値両対応

## 2026-03 生の声取得をGoogle News RSS + YouTubeに変更（v3.0）
- 決定：台湾フォーラムの代わりに Google News RSS（台湾）と YouTube Data API v3 を採用
- 理由：背包客棧などJSレンダリングサイトはスクレイピング不可、FB/InstagramはAPI認証が必要で取得不可
- 却下した代替案：Facebook Graph API（認証・審査が必要で実用的でない）、Mobile01（JSレンダリングか要検証）
- Google News RSS：APIキー不要で即利用可能。feedparserでパース
- YouTube：Data API v3（無料10,000ユニット/日）、サイドバーにAPIキー入力欄を追加

---

# 作業ログ

## 2025-02（セッション1〜複数）
- Streamlit アプリ（research_app.py）の基本構造構築
- タブ構成：① JNTO訪問者データ ② Google Trends ③ PTT/フォーラム分析
- Gemini AI統合、WordCloud、jieba、繁体字変換を実装
- Dcard → 背包客棧 → 外部リンクパネルへの移行
- Gemini API：google-generativeai から REST 直接呼び出しに変更
- Gemini レート制限対策：ボタントリガー化・session_stateキャッシュ・モデル自動検出

## 2026-03 v3.0 大規模改修
- **① JNTO自動取得**: JNTOの統計ページをスクレイピング → .xlsxを自動DL → openpyxlで台湾行を自動抽出・グラフ表示（24時間キャッシュ）。CSVアップロードはフォールバックとして維持。
- **② Trendsスポット候補**: Google Trendsの関連クエリ（Top/Rising）からカテゴリ判定してスポット候補を動的表示。タブ2の取得データをsession_stateでタブ1と共有。
- **③ Gemini廃止 → テンプレート分析**: Gemini API 完全削除。`generate_trends_analysis()` と `generate_persona_analysis()` をルールベースで実装（APIキー不要・無制限）。
- **④ 生の声改修**: Google News RSS（台湾）をfeedparserで自動取得 + YouTube Data API v3（任意キー）を追加。PTTスクレイピングは維持。PTT+ニュース統合のキーワード分析・WordCloud・ペルソナ分析も提供。
- **依存ライブラリ**: openpyxl・feedparser追加、google-generativeai削除

## 2026-03 v3.0 バグ修正 + 実務改善
- **JNTOマルチイヤー修正**: pandas `read_excel` による結合セル取得の根本的問題を解決。openpyxl で直接 `_unmerge_ws()` を実装し全結合セルを展開後に台湾行を解析。年ヘッダーが文字列（"2024年"等）でも `re.search(r'(20\d{2})')` で正確に抽出。パスA（正確モード）/パスB（フォールバック）の2段階解析で複数年に対応。
- **月別季節性チャート追加**: `fetch_jnto_taiwan_data()` の戻り値を `(yearly_df, monthly_df, msg)` の3値タプルに変更。直近3年の月別平均をbar_chartで表示し、訪問ピーク月とプロモーション推奨タイミングを自動算出して表示。
- **KPI サマリーカード**: Tab1にst.metricによる4指標カード（直近年訪問者数・前年比・2019年比・収録年数）を追加。
- **自治体向けクイックガイド**: ページ最上部にエクスパンダー形式でデータソース一覧・推奨ワークフロー・提案説明ポイントを追加。
- **提案書コピペ用サマリー**: Tab3末尾にボタン一発でJNTO・Trends・PTTデータを統合したMarkdownサマリーを生成しダウンロードできる機能を追加。
- **NameError対策**: `top_kws` を `if combined_titles:` ブロックの外側で事前に `[]` で初期化。

## 2026-03 提案書サマリー機能のプロフェッショナル版へ刷新
- `_generate_proposal_text()` 関数を新設（約250行）。7章構成の本格的な市場調査レポートを自動生成。
- **構成**: 0.エグゼクティブサマリー / 1.台湾インバウンド市場現状 / 2.ニーズ分析 / 3.課題・リスク / 4.アクションプラン（短中長期）/ 5.JR連携スキーム / 6.経済効果試算 / 7.付録（データ根拠）
- 全取得データ（JNTO・Trends・PTT・ニュース）を統合して定量分析、業界知識ベースで課題・提言を補完。
- UI改善: Markdownプレビュー + Markdownダウンロード + テキストダウンロード（2形式）の2ボタン出力。

## 2026-03 バグ修正2件
- **strftime ValueError**: `%-m` / `%-d` はWindowsで動作しない（Linux専用書式）。`f"{d.year}年{d.month}月{d.day}日"` のf-string直接フォーマットに変更。今後もWindows環境ではstrftimeの`%-`系書式は使用禁止。
- **グラフ改善**: 年別折れ線（st.line_chart）→ 年別棒グラフ（st.bar_chart）に変更。月別データがある場合は月別棒グラフをメイン表示（高さ320px）、年別をサブ表示に配置。月別データがない場合は年別棒グラフのみ表示。

## 2026-03 提案書品質改善（3件）+ 品質チェックパネル追加
- **不適切KWフィルタ**: `_INAPPROPRIATE_PROMO_KWS` frozensetと `_filter_promo_kws()` を追加。Google Trendsの急上昇KWから地震・台風・事故などの語句を除外し、プロモーション不適切なワードがレポートに入らないようにした。
- **FAMアクション説明改善**: 「FAM → 販売ルート確立」のロジックを3ステップの連鎖効果（商品化→社内稟議→自発発信）として具体的に説明。過去事例の費用対効果（100〜300万円で3〜5社商品化）も追記。
- **モデルコース具体化**: `get_inbound_spots(kw)` で実際のスポットデータを取得し「東京/成田→{kw}→スポット1→スポット2」の実際のルートをアクションプランに埋め込む。掲載交渉先（雄獅旅行・可樂旅遊・EZ Travel・kkday）も具体名を記載。
- **データ品質チェックパネル**: 提案書生成ボタンの前に `st.expander` で品質パネルを追加。100点満点のスコア（JNTO 25点・月別10点・Trends 35点・PTT 15点・ニュース15点）と除外KW一覧を可視化。ユーザーが生成前にデータ品質を確認・改善できる仕組み。
