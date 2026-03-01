"""
🇹🇼 台湾市場向け 事前リサーチダッシュボード v3.0
JR東日本 台湾事業開発 | 日本自治体向け提案 事前調査ツール

変更履歴 v3.0:
- JNTO データを自動取得（CSV アップロード不要）
- Google Trends 関連クエリからトレンドスポット候補を動的抽出
- Gemini AI → テンプレートベース分析（APIキー不要・無制限）
- 生の声：台湾フォーラムスクレイピング → Google News RSS + YouTube 動画検索

実行方法:
    streamlit run research_app.py
"""

import io
import os
import re
import time
from collections import Counter
from urllib.parse import quote

import matplotlib
matplotlib.use("Agg")  # GUI バックエンドを無効化（Streamlit との競合防止）
import matplotlib.pyplot as plt
import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

# ============================================================
# ページ設定（必ず最初に呼ぶ）
# ============================================================
st.set_page_config(
    page_title="台湾市場向け 事前リサーチダッシュボード",
    page_icon="🇹🇼",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# カスタム CSS
# ============================================================
st.markdown(
    """
<style>
    /* インサイトボックス（JR東日本レッド） */
    .insight-box {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border-left: 5px solid #c0392b;
        padding: 1.3rem 1.6rem;
        border-radius: 0 10px 10px 0;
        margin: 1rem 0 1.5rem 0;
        color: #ecf0f1;
        line-height: 2.1;
        font-size: 0.94rem;
    }
    .insight-box strong { color: #f1c40f; }

    /* 分析ボックス（統一デザイン） */
    .analysis-box {
        background: linear-gradient(135deg, #1a2e1a 0%, #162e1e 100%);
        border-left: 5px solid #27ae60;
        padding: 1.3rem 1.6rem;
        border-radius: 0 10px 10px 0;
        margin: 1rem 0 1.5rem 0;
        color: #dce9dc;
        line-height: 1.9;
    }

    /* ニュースカード */
    .news-card {
        background: #1e2430;
        border: 1px solid #2d3748;
        border-radius: 8px;
        padding: 0.8rem 1.0rem;
        margin: 0.4rem 0;
        color: #e2e8f0;
        font-size: 0.88rem;
    }
    .news-card a { color: #63b3ed; text-decoration: none; }
    .news-card .source { color: #a0aec0; font-size: 0.78rem; }

    /* タブラベル */
    .stTabs [data-baseweb="tab"] {
        font-size: 1.0rem;
        font-weight: 700;
        padding: 0.55rem 1.1rem;
    }

    /* サイドバー見出し */
    .sidebar-section {
        font-size: 0.78rem;
        color: #888;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 0.3rem;
    }
</style>
""",
    unsafe_allow_html=True,
)

# ============================================================
# 定数
# ============================================================
PTT_SEARCH_URL   = "https://www.ptt.cc/bbs/Japan_Travel/search"
JNTO_STATS_URL   = "https://www.jnto.go.jp/statistics/data/visitors-statistics/"
JNTO_BASE_URL    = "https://www.jnto.go.jp"
FONT_CACHE_PATH  = "cjk_font_cache.ttf"

# CJK フォント候補（OS 別）
_SYSTEM_FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\msjhbd.ttc",
    r"C:\Windows\Fonts\mingliu.ttc",
    r"C:\Windows\Fonts\kaiu.ttf",
    r"C:\Windows\Fonts\meiryo.ttc",
    r"C:\Windows\Fonts\msgothic.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
]

# ストップワード（PTT / 台湾フォーラム タイトル用）
_STOP_WORDS: frozenset[str] = frozenset(
    {
        "日本", "旅遊", "旅行", "問題", "推薦", "請問", "有人",
        "自由行", "行程", "分享", "求助", "心得", "想問", "怎麼",
        "可以", "沒有", "感謝", "謝謝", "請教", "大家", "一下",
        "之旅", "wiki", "想去", "需要", "關於", "一些",
        "有沒有", "這樣", "那個", "什麼", "如何", "地方", "問一下",
        "附近", "跪求", "跪問", "請益", "心得文", "遊記", "再問",
        "Re", "Fw", "fw",
    }
)

# ============================================================
# 観光プロモーション文脈で使用不適切なキーワード
# Google Trends の Rising KW にはニュース性の高い語（災害・事故など）が
# 混入することがある。これらを観光コンテンツに使用することは
# ブランドイメージ毀損・不謹慎となるためフィルタリングする。
# ============================================================
_INAPPROPRIATE_PROMO_KWS: frozenset[str] = frozenset({
    # 災害・安全系
    "地震", "震", "台風", "津波", "洪水", "噴火", "土砂", "崩壊", "豪雨",
    "大雨", "避難", "警報", "注意報", "災害", "被害", "被災", "復旧",
    "停電", "断水", "交通規制", "通行止め", "緊急",
    # 事故・事件系
    "事故", "事件", "死亡", "死者", "負傷", "けが", "怪我",
    "犯罪", "逮捕", "殺", "傷害",
    # 繁体字対応
    "地震", "台風", "海嘯", "洪水", "火山", "噴發", "土石流",
    "事故", "死亡", "死傷", "受傷", "犯罪", "逃跑", "災害",
    # 政治・外交系（観光PRに不適切）
    "戦争", "戦争", "軍事", "ミサイル", "制裁", "外交",
})


def _filter_promo_kws(kws: list[str]) -> list[str]:
    """
    観光プロモーション文脈で使用不適切なキーワードを除外する。
    地震・事故・災害などのキーワードが含まれるものを排除し、
    純粋な観光需要ワードのみを返す。
    """
    result = []
    for k in kws:
        if not any(ng in k for ng in _INAPPROPRIATE_PROMO_KWS):
            result.append(k)
    return result


# テンプレート分析用カテゴリキーワード
_CAT_KW: dict[str, list[str]] = {
    "温泉・スパ":       ["溫泉", "泡湯", "露天", "湯", "温泉", "hot spring", "溫泉鄉"],
    "グルメ・食":       ["美食", "吃", "料理", "食", "餐", "壽司", "拉麵", "牛排", "海鮮", "甜點", "咖啡", "燒肉", "握壽司"],
    "景観・自然":       ["景色", "風景", "自然", "湖", "山", "海", "瀑布", "花", "楓", "雪", "森", "公園", "高原", "絕景"],
    "ショッピング":      ["購物", "免稅", "藥妝", "唐吉訶德", "outlet", "伴手禮", "手信", "血拼", "逛街"],
    "歴史・文化":       ["神社", "城", "古街", "歷史", "寺", "古蹟", "博物館", "傳統", "藝術", "世界遺產"],
    "体験・アクティビティ": ["體驗", "滑雪", "搭乘", "騎", "參觀", "活動", "DIY", "工藝", "攀登", "健行"],
    "交通・アクセス":    ["新幹線", "電車", "JR", "交通", "巴士", "機場", "票", "交通券", "周遊券"],
    "宿泊":             ["住宿", "旅館", "酒店", "民宿", "飯店", "旅店", "旅宿"],
    "季節・イベント":    ["祭り", "祭典", "花火", "楓葉", "賞楓", "賞花", "跨年", "煙火", "慶典", "節慶"],
    "テーマパーク":      ["迪士尼", "環球", "遊樂園", "樂園", "主題"],
}

# ============================================================
# インバウンド人気スポットデータ（地域別）
# ============================================================
_INBOUND_SPOTS: dict[str, list[dict]] = {
    "北海道": [
        {"カテゴリ": "🏔 景観・自然", "スポット名": "小樽運河",             "台湾人への訴求ポイント": "石造り倉庫群と運河がSNS映え抜群。台湾旅行誌掲載率No.1の定番スポット", "最寄り駅・JR連携": "JR小樽駅 徒歩10分"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "函館山夜景",            "台湾人への訴求ポイント": "世界三大夜景のひとつ。台湾人カップル・新婚旅行に圧倒的人気", "最寄り駅・JR連携": "JR函館駅からロープウェイ乗り場へバス"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "富良野・美瑛のラベンダー畑", "台湾人への訴求ポイント": "7月の紫の絶景。写真撮影目的の訪問者が多く「夢の大地」として台湾で知名度大", "最寄り駅・JR連携": "JR富良野駅・美馬牛駅"},
        {"カテゴリ": "🍜 グルメ",     "スポット名": "二条市場・場外市場（カニ・ウニ）", "台湾人への訴求ポイント": "新鮮な海産物を市場で食べる体験が台湾SNSで拡散中", "最寄り駅・JR連携": "JR札幌駅 徒歩15分"},
        {"カテゴリ": "🎪 体験",       "スポット名": "旭山動物園",            "台湾人への訴求ポイント": "ペンギンのもぐもぐタイムが台湾人親子旅行に大人気", "最寄り駅・JR連携": "JR旭川駅からバス約40分"},
    ],
    "青森": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "弘前城・弘前公園（桜）", "台湾人への訴求ポイント": "日本最大級の桜の名所。4月下旬に台湾人の検索が急上昇する季節性コンテンツ", "最寄り駅・JR連携": "JR弘前駅 徒歩20分"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "奥入瀬渓流",            "台湾人への訴求ポイント": "新緑・紅葉シーズンに台湾誌が特集。「日本の原風景」として台湾富裕層に支持", "最寄り駅・JR連携": "JR八戸駅から奥入瀬渓流バス"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "十和田湖",              "台湾人への訴求ポイント": "神秘的な火山湖。奥入瀬とセットで周遊するコースが台湾旅行ブログで定番化", "最寄り駅・JR連携": "JR青森駅・JR八戸駅からバス"},
        {"カテゴリ": "🎪 体験",       "スポット名": "青森ねぶた祭（8月）",   "台湾人への訴求ポイント": "迫力の灯籠行列。祭り体験ニーズが高い台湾人に刺さるコンテンツ", "最寄り駅・JR連携": "JR青森駅 徒歩圏内"},
    ],
    "岩手": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "平泉・中尊寺金色堂",    "台湾人への訴求ポイント": "世界文化遺産。「黄金の世界」として中国語圏に強い訴求力", "最寄り駅・JR連携": "JR平泉駅 徒歩20分"},
        {"カテゴリ": "🍜 グルメ",     "スポット名": "盛岡冷麺・わんこそば",  "台湾人への訴求ポイント": "食体験として台湾のグルメ系インフルエンサーが多数紹介", "最寄り駅・JR連携": "JR盛岡駅 市内徒歩圏"},
    ],
    "宮城": [
        {"カテゴリ": "🏔 景観・自然", "スポット名": "松島（日本三景）",       "台湾人への訴求ポイント": "日本三景のひとつ。台湾人の「日本の絶景」コンテンツで常に上位にランク", "最寄り駅・JR連携": "JR松島海岸駅 徒歩すぐ"},
        {"カテゴリ": "🍜 グルメ",     "スポット名": "仙台牛タン通り（仙台駅前）", "台湾人への訴求ポイント": "「牛舌料理」として台湾SNSで話題。仙台訪問の最大動機のひとつ", "最寄り駅・JR連携": "JR仙台駅 直結"},
        {"カテゴリ": "♨ 温泉",       "スポット名": "秋保温泉・作並温泉",     "台湾人への訴求ポイント": "仙台から近い温泉として台湾人の日帰り・1泊プランに最適", "最寄り駅・JR連携": "JR仙台駅からバス約40分"},
    ],
    "秋田": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "角館武家屋敷・桜並木",   "台湾人への訴求ポイント": "「日本の小京都」として台湾人に人気。桜シーズンの混雑が台湾ブログで話題", "最寄り駅・JR連携": "JR角館駅 徒歩15分"},
        {"カテゴリ": "🎪 体験",       "スポット名": "秋田竿燈まつり（8月）",  "台湾人への訴求ポイント": "秋田のダイナミックな竿燈パフォーマンスが台湾のSNSで驚きを持って紹介", "最寄り駅・JR連携": "JR秋田駅 徒歩10分"},
    ],
    "山形": [
        {"カテゴリ": "♨ 温泉",       "スポット名": "銀山温泉",               "台湾人への訴求ポイント": "大正ロマン建築の街並みが台湾女性旅行者に絶大な人気。SNS投稿数が東北No.1級", "最寄り駅・JR連携": "JR大石田駅からバス30分"},
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "山寺（立石寺）",         "台湾人への訴求ポイント": "1000段以上の石段と山岳寺院の絶景。台湾の旅行誌で「東北必見スポット」に選出", "最寄り駅・JR連携": "JR山寺駅 徒歩すぐ"},
        {"カテゴリ": "🎪 体験",       "スポット名": "蔵王温泉スキー場・樹氷",  "台湾人への訴求ポイント": "スキー初心者対応のスクールが充実。台湾からの「初スキー旅行」先として人気上昇", "最寄り駅・JR連携": "JR山形駅からバス40分"},
    ],
    "福島": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "大内宿（会津）",         "台湾人への訴求ポイント": "茅葺き屋根の宿場町が「日本昔話の世界」として台湾Instagramで拡散", "最寄り駅・JR連携": "JR湯野上温泉駅 バス"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "五色沼・磐梯山",         "台湾人への訴求ポイント": "エメラルドグリーンの湖沼群。「神秘の色」として台湾の写真愛好家に人気", "最寄り駅・JR連携": "JR猪苗代駅 バス20分"},
    ],
    "栃木": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "日光東照宮",             "台湾人への訴求ポイント": "世界文化遺産。「三猿・眠り猫」が台湾人に人気のフォトスポット", "最寄り駅・JR連携": "JR日光駅・東武日光駅 バス"},
        {"カテゴリ": "♨ 温泉",       "スポット名": "鬼怒川温泉",             "台湾人への訴求ポイント": "東京からのアクセスが良い温泉地として台湾1泊旅行の定番", "最寄り駅・JR連携": "東武鬼怒川温泉駅 直近"},
    ],
    "群馬": [
        {"カテゴリ": "♨ 温泉",       "スポット名": "草津温泉",               "台湾人への訴求ポイント": "日本三名泉のひとつ。湯畑の写真が台湾SNSで多数投稿。免税店も充実", "最寄り駅・JR連携": "JR長野原草津口駅 バス25分"},
        {"カテゴリ": "♨ 温泉",       "スポット名": "伊香保温泉（石段街）",   "台湾人への訴求ポイント": "石段と温泉街の風情が「昭和レトロ」として台湾若者層に人気上昇", "最寄り駅・JR連携": "JR渋川駅 バス25分"},
    ],
    "千葉": [
        {"カテゴリ": "🎪 体験",       "スポット名": "東京ディズニーリゾート（浦安）", "台湾人への訴求ポイント": "台湾人の「東京旅行」で最優先されるスポット。家族旅行の絶対的目的地", "最寄り駅・JR連携": "JR舞浜駅 直近"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "国営ひたち海浜公園（ネモフィラ）", "台湾人への訴求ポイント": "4〜5月の青いネモフィラが台湾インスタグラマーに大人気。検索急上昇中", "最寄り駅・JR連携": "JR勝田駅 バス15分"},
    ],
    "東京": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "浅草・雷門・仲見世通り",  "台湾人への訴求ポイント": "東京観光の定番No.1。台湾人の訪問率が最も高い", "最寄り駅・JR連携": "TX浅草駅・東京メトロ浅草駅"},
        {"カテゴリ": "🎪 体験",       "スポット名": "渋谷スクランブル交差点", "台湾人への訴求ポイント": "「都市の象徴」として台湾人の撮影必須スポット", "最寄り駅・JR連携": "JR渋谷駅 直結"},
        {"カテゴリ": "🛍 ショッピング", "スポット名": "秋葉原（電気街・アニメ）", "台湾人への訴求ポイント": "アニメ・ゲーム・ガジェット目的の台湾20代に絶大な人気", "最寄り駅・JR連携": "JR秋葉原駅 直近"},
    ],
    "神奈川": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "鎌倉大仏・鶴岡八幡宮",  "台湾人への訴求ポイント": "「ちむどんどん聖地」鎌倉が台湾ドラマファンにも人気", "最寄り駅・JR連携": "JR鎌倉駅 徒歩・バス"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "箱根（芦ノ湖・富士山ビュー）", "台湾人への訴求ポイント": "富士山をバックにした芦ノ湖の写真が台湾SNSで多数投稿", "最寄り駅・JR連携": "JR小田原駅 箱根登山鉄道乗換"},
        {"カテゴリ": "🍜 グルメ",     "スポット名": "横浜中華街",             "台湾人への訴求ポイント": "台湾人にとって親しみやすい中華フード体験。食べ歩きが人気", "最寄り駅・JR連携": "JR石川町駅 徒歩5分"},
    ],
    "新潟": [
        {"カテゴリ": "🎪 体験",       "スポット名": "越後湯沢スキー場",       "台湾人への訴求ポイント": "東京から新幹線1時間強でアクセスできるスキーリゾートとして台湾人に注目", "最寄り駅・JR連携": "JR越後湯沢駅 直近"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "清津峡渓谷トンネル",     "台湾人への訴求ポイント": "鏡面の床と峡谷が作る絶景が台湾インスタで爆発的に拡散。SNSスポット筆頭", "最寄り駅・JR連携": "JR越後湯沢駅 バス"},
    ],
    "長野": [
        {"カテゴリ": "🎪 体験",       "スポット名": "白馬スキー場",           "台湾人への訴求ポイント": "ニセコに次ぐ「粉雪の聖地」として台湾スキーヤーに人気急上昇中", "最寄り駅・JR連携": "JR白馬駅 直近"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "上高地",                 "台湾人への訴求ポイント": "北アルプスの秘境。台湾人ハイカーに「一生に一度は行くべき場所」と認識", "最寄り駅・JR連携": "JR松本駅 バス乗換"},
        {"カテゴリ": "♨ 温泉",       "スポット名": "地獄谷野猿公苑（スノーモンキー）", "台湾人への訴求ポイント": "温泉に入るニホンザルが世界的に有名。台湾人の「冬の長野」最大の目的", "最寄り駅・JR連携": "JR湯田中駅 バス・徒歩"},
    ],
    "山梨": [
        {"カテゴリ": "🏔 景観・自然", "スポット名": "富士山（五合目・須走口）", "台湾人への訴求ポイント": "台湾人の「日本旅行で行きたい場所No.1」。7〜9月の登山シーズンに集中", "最寄り駅・JR連携": "JR富士駅・JR吉田駅 バス"},
        {"カテゴリ": "🏔 景観・自然", "スポット名": "河口湖・忍野八海",       "台湾人への訴求ポイント": "富士山リフレクションの撮影スポット。台湾の写真家・インフルエンサーの聖地", "最寄り駅・JR連携": "JR大月駅 富士急行乗換"},
    ],
    "富山": [
        {"カテゴリ": "🏔 景観・自然", "スポット名": "立山黒部アルペンルート",  "台湾人への訴求ポイント": "雪の大谷（4〜6月）が台湾で「奇跡の絶景」として大人気。予約が殺到", "最寄り駅・JR連携": "JR富山駅 富山地鉄・立山駅乗換"},
    ],
    "石川": [
        {"カテゴリ": "🏔 景観・自然", "スポット名": "兼六園",                 "台湾人への訴求ポイント": "日本三名園のひとつ。雪吊りの冬景色が台湾人に「京都以外の和」として人気", "最寄り駅・JR連携": "JR金沢駅 バス20分"},
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "ひがし茶屋街",           "台湾人への訴求ポイント": "着物レンタル体験が台湾人女性に絶大人気。Instagramで金沢の象徴的スポット", "最寄り駅・JR連携": "JR金沢駅 バス"},
        {"カテゴリ": "🍜 グルメ",     "スポット名": "近江町市場（海鮮丼・カニ）", "台湾人への訴求ポイント": "「金沢の台所」で海鮮食べ歩き。カニ目的の台湾人が冬に急増", "最寄り駅・JR連携": "JR金沢駅 徒歩15分"},
    ],
    "福井": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "永平寺",                 "台湾人への訴求ポイント": "禅の体験ができる格式ある寺院。精進料理体験が台湾の「本物の日本」需要に合致", "最寄り駅・JR連携": "JR福井駅 バス30分"},
        {"カテゴリ": "🎪 体験",       "スポット名": "福井県立恐竜博物館",     "台湾人への訴求ポイント": "世界三大恐竜博物館のひとつ。台湾人家族旅行の目的地として注目度が急上昇", "最寄り駅・JR連携": "JR福井駅 バス・勝山駅"},
    ],
    "岐阜": [
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "白川郷合掌造り集落",     "台湾人への訴求ポイント": "世界文化遺産。冬の雪景色ライトアップが台湾で「幻の絶景」として大人気", "最寄り駅・JR連携": "JR高山駅 バス50分"},
        {"カテゴリ": "🏯 歴史・文化", "スポット名": "飛騨高山古い町並み（三町）", "台湾人への訴求ポイント": "江戸時代の町並みと地酒・飛騨牛の組み合わせ。台湾の文化旅行者に大人気", "最寄り駅・JR連携": "JR高山駅 徒歩10分"},
    ],
}


def get_inbound_spots(keyword: str) -> list[dict] | None:
    """キーワードに対応するインバウンド人気スポットリストを返す。"""
    for region_key, spots in _INBOUND_SPOTS.items():
        if region_key in keyword or keyword in region_key:
            return spots
    return None


# ============================================================
# 繁体字変換
# ============================================================

_WORD_TC_MAP: dict[str, str] = {
    "温泉": "溫泉",
    "白川郷": "白川鄉",
    "軽井沢": "輕井澤",
    "観光": "觀光",
    "万座": "萬座",
    "国立": "國立",
    "北海道": "北海道",
}

_CHAR_TC_MAP: dict[str, str] = {
    "県": "縣", "区": "區", "郷": "鄉", "万": "萬",
    "温": "溫", "覧": "覽", "観": "觀", "転": "轉",
    "経": "經", "応": "應", "発": "發", "様": "樣",
    "総": "總", "関": "關", "団": "團", "変": "變",
    "実": "實", "国": "國", "気": "氣", "体": "體", "処": "處",
}


def to_tc(keyword: str) -> str:
    """日本語キーワードを台湾繁体字に変換する。"""
    result = keyword
    for jp_word, tc_word in _WORD_TC_MAP.items():
        result = result.replace(jp_word, tc_word)
    for jp_char, tc_char in _CHAR_TC_MAP.items():
        result = result.replace(jp_char, tc_char)
    return result


# ============================================================
# CJK フォント取得（WordCloud 用）
# ============================================================


@st.cache_resource(show_spinner=False)
def get_cjk_font_path() -> str | None:
    """CJK 対応フォントのパスを返す。"""
    for path in _SYSTEM_FONT_CANDIDATES:
        if os.path.exists(path):
            return path

    if os.path.exists(FONT_CACHE_PATH):
        return FONT_CACHE_PATH

    _DL_URL = (
        "https://github.com/googlefonts/noto-fonts/raw/main/"
        "hinted/ttf/NotoSansSC/NotoSansSC-Regular.ttf"
    )
    try:
        resp = requests.get(_DL_URL, timeout=40, stream=True)
        resp.raise_for_status()
        with open(FONT_CACHE_PATH, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        return FONT_CACHE_PATH
    except Exception:
        return None


# ============================================================
# ① JNTO 自動データ取得
# ============================================================


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_jnto_excel_url() -> str | None:
    """
    JNTOの統計ページをスクレイピングして最新の Excel ファイルの完全URLを返す。
    1日キャッシュ（ttl=86400）。
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        )
    }
    try:
        resp = requests.get(JNTO_STATS_URL, headers=headers, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        # href に .xlsx を含むリンクを最初に見つける
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if ".xlsx" in href.lower():
                if href.startswith("http"):
                    return href
                return JNTO_BASE_URL + href
    except Exception:
        pass
    return None


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_jnto_taiwan_data() -> tuple[pd.DataFrame | None, pd.DataFrame | None, str]:
    """
    JNTO Excel（1シート1年形式、2003〜現在）をダウンロードし、
    台湾の年別・月別訪問者数を返す。

    ■ 実際のExcel構造（診断済み）:
    - シート名 = 年（"2003", "2004", ..., "2026"）、計24シート
    - 各シート: 行0=タイトル, 行3=月ヘッダー, 行8=台湾データ
    - 列構造（新形式 2024以降）:
        [国名, None, 1月, 1月前比%, 2月, 2月前比%, ..., 12月, 12月前比%, 計, 計前比%]
    - 列構造（旧形式 2023以前）:
        [国名, 1月, 1月前比%, 2月, 2月前比%, ..., 12月, 12月前比%, 計, 計前比%]
    - 判定方法: taiwan_row[1] が None → 新形式（first_col=2）、数値 → 旧形式（first_col=1）
    - 月m列インデックス: first_col + (m-1) * 2
    - 年計列インデックス: first_col + 24

    Returns:
        (yearly_df または None, monthly_df または None, メッセージ)
        yearly_df 列: 年, 台湾人訪問者数（人）
        monthly_df 列: 年, 月, 訪問者数
    """
    import openpyxl

    excel_url = fetch_jnto_excel_url()
    if not excel_url:
        return None, None, "JNTOページからExcelファイルのURLが取得できませんでした"

    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        resp = requests.get(excel_url, headers=headers, timeout=60)
        resp.raise_for_status()
    except requests.exceptions.Timeout:
        return None, None, "JNTOサーバーへの接続がタイムアウトしました"
    except requests.exceptions.ConnectionError:
        return None, None, "JNTOサーバーへの接続に失敗しました"
    except Exception as exc:
        return None, None, f"ダウンロードエラー: {exc}"

    bio = io.BytesIO(resp.content)
    try:
        wb = openpyxl.load_workbook(bio, data_only=True)
    except Exception as exc:
        return None, None, f"Excelファイルの読み込みに失敗しました: {exc}"

    monthly_records: list[dict] = []
    yearly_records: list[dict] = []

    for ws in wb.worksheets:
        # シート名から年を取得（シート名 = "2003" 〜 "2026"）
        try:
            sheet_year = int(str(ws.title).strip())
            if not (2000 <= sheet_year <= 2030):
                continue
        except (ValueError, AttributeError):
            continue

        # Excelの全行を2Dリストに変換
        grid = [[cell.value for cell in row] for row in ws.iter_rows()]

        if len(grid) < 9:
            continue  # 行数が足りない場合はスキップ

        # 台湾行を探す（通常は行8 = index 8）
        taiwan_row: list | None = None
        for ri in range(min(len(grid), 20)):
            for ci in range(min(5, len(grid[ri]))):
                val = grid[ri][ci]
                if isinstance(val, str) and ("台湾" in val or "臺灣" in val):
                    taiwan_row = grid[ri]
                    break
            if taiwan_row is not None:
                break

        if taiwan_row is None:
            continue

        # 列フォーマット判定:
        #   新形式（2024以降）: col1 = None → データはcol2から
        #   旧形式（2023以前）: col1 = 1月値（数値）→ データはcol1から
        col1_val = taiwan_row[1] if len(taiwan_row) > 1 else None
        if col1_val is None:
            first_col = 2  # 新形式
        elif isinstance(col1_val, (int, float)) and not isinstance(col1_val, bool):
            first_col = 1  # 旧形式
        else:
            first_col = 2  # デフォルトは新形式

        # 月別データを抽出
        # 月 m のデータ列: first_col + (m-1) * 2
        for month in range(1, 13):
            col_idx = first_col + (month - 1) * 2
            if col_idx < len(taiwan_row):
                val = taiwan_row[col_idx]
                if isinstance(val, (int, float)) and not isinstance(val, bool) and val > 0:
                    monthly_records.append({
                        "年": sheet_year,
                        "月": month,
                        "訪問者数": int(round(val)),
                    })

        # 年計データを抽出（計列 = first_col + 12*2 = first_col + 24）
        annual_col = first_col + 24
        if annual_col < len(taiwan_row):
            annual_val = taiwan_row[annual_col]
            if isinstance(annual_val, (int, float)) and not isinstance(annual_val, bool) and annual_val > 0:
                yearly_records.append({
                    "年": sheet_year,
                    "台湾人訪問者数（人）": int(round(annual_val)),
                })
            else:
                # 年計列がない場合は月別合計から計算
                month_sum = sum(
                    r["訪問者数"] for r in monthly_records if r["年"] == sheet_year
                )
                if month_sum > 0:
                    yearly_records.append({
                        "年": sheet_year,
                        "台湾人訪問者数（人）": month_sum,
                    })

    if not yearly_records:
        return None, None, "Excelファイル内に台湾データが見つかりませんでした"

    # 年昇順にソート
    yearly_df = pd.DataFrame(yearly_records).sort_values("年").reset_index(drop=True)
    monthly_df: pd.DataFrame | None = None
    if monthly_records:
        monthly_df = (
            pd.DataFrame(monthly_records)
            .sort_values(["年", "月"])
            .reset_index(drop=True)
        )

    return yearly_df, monthly_df, "OK"


# ============================================================
# ② Google Trends 関連クエリからスポット候補を抽出
# ============================================================


def extract_trending_spots(
    top_df: pd.DataFrame | None,
    rising_df: pd.DataFrame | None,
    base_keyword: str,
) -> list[dict]:
    """
    Google Trends の関連クエリからスポット・テーマ候補を抽出する。
    ストップワードを除き、2文字以上の語をリストアップする。
    """
    candidates = []
    seen: set[str] = set()

    for df, label in [(rising_df, "🚀 急上昇"), (top_df, "🏆 定番人気")]:
        if df is None or df.empty:
            continue
        for _, row in df.iterrows():
            q = str(row.get("query", "")).strip()
            if not q or q in seen:
                continue
            if base_keyword in q or q in base_keyword:
                seen.add(q)
                continue
            if any(sw in q for sw in _STOP_WORDS):
                continue
            if len(q) < 2:
                continue

            # カテゴリ判定
            detected_cat = "その他"
            for cat, kws in _CAT_KW.items():
                if any(kw.lower() in q.lower() for kw in kws):
                    detected_cat = cat
                    break

            seen.add(q)
            val = row.get("value", "")
            candidates.append({
                "トレンド種別": label,
                "キーワード": q,
                "カテゴリ": detected_cat,
                "相対値": str(val),
            })

    return candidates


# ============================================================
# ③ テンプレートベース分析（Gemini 代替・API不要）
# ============================================================

# 季節定義
_SEASON_MAP = {
    "冬": [12, 1, 2],
    "春": [3, 4, 5],
    "夏": [6, 7, 8],
    "秋": [9, 10, 11],
}

_SEASON_CONTENT = {
    "冬": ("雪景色・温泉・年末年始",
           "冬季（12〜2月）がピーク。温泉・雪景色・年越し旅行プランの訴求が有効。\n"
           "12月上旬〜1月上旬の年末年始パッケージは台湾人の予約が集中します。"),
    "春": ("桜・新緑・GW",
           "春（3〜5月）が最盛期。桜シーズン（3月下旬〜4月下旬）に合わせた\n"
           "2〜3月のSNSキャンペーン・旅行博への出展が最も効果的です。"),
    "夏": ("夏祭り・海水浴・アウトドア",
           "夏（6〜8月）にピーク。台湾の夏休み（7〜8月）と重なるため\n"
           "家族旅行・夏祭り体験コンテンツへの関心が特に高い季節です。"),
    "秋": ("紅葉・秋グルメ・ハロウィン",
           "秋（9〜11月）がピーク。紅葉狩り・食の秋への需要が最大化します。\n"
           "10〜11月を軸にした紅葉スポット訴求・グルメ旅行企画が有効です。"),
}


def _detect_categories(queries: list[str]) -> dict[str, list[str]]:
    """クエリリストをカテゴリ別に分類する。"""
    result: dict[str, list[str]] = {}
    for q in queries:
        for cat, kws in _CAT_KW.items():
            if any(kw.lower() in q.lower() for kw in kws):
                result.setdefault(cat, []).append(q)
    return result


def generate_trends_analysis(
    keyword: str,
    tc_keyword: str,
    interest_df,
    top_df,
    rising_df,
) -> str:
    """
    Google Trends データからテンプレートベースで詳細分析レポートを生成する（Markdown）。
    外部API不要・即時実行・無制限。
    """
    parts: list[str] = [
        f"## 📊 「{keyword}」台湾向けトレンド分析レポート",
        f"*調査キーワード（繁体字）: {tc_keyword}　|　データソース: Google Trends（台湾・過去12ヶ月）*",
        "",
        "---",
        "",
    ]

    # ── Section 1: 検索ボリューム詳細分析 ──
    ts = None
    peak_month = None
    season = ""
    season_kw = ""

    if interest_df is not None and not interest_df.empty:
        cols = [c for c in interest_df.columns if c != "isPartial"]
        if cols:
            ts = interest_df[cols[0]]
            peak_idx   = ts.idxmax()
            trough_idx = ts.idxmin()
            peak_month = peak_idx.month
            avg_val    = int(ts.mean())
            max_val    = int(ts.max())
            min_val    = int(ts.min())
            volatility = max_val - min_val

            # 季節判定
            season = next(
                (s for s, months in _SEASON_MAP.items() if peak_month in months), "不明"
            )
            season_kw, season_advice = _SEASON_CONTENT.get(
                season, ("", "シーズン情報が不明です")
            )

            # 安定性評価
            if volatility > 60:
                stability_label = "非常に強い季節性（オフシーズン対策が重要）"
            elif volatility > 30:
                stability_label = "中程度の季節性（閑散期プロモーションで平準化可能）"
            else:
                stability_label = "通年安定した需要（年間を通じた施策が有効）"

            # 直近3ヶ月の傾向
            recent_3m = ts.iloc[-3:]
            recent_avg = int(recent_3m.mean())
            trend_vs_avg = recent_avg - avg_val
            trend_direction = "📈 上昇傾向" if trend_vs_avg > 5 else (
                "📉 下降傾向" if trend_vs_avg < -5 else "➡️ 横ばい"
            )

            # 季節ごとの平均を計算
            month_avgs: dict[str, int] = {}
            for s, months in _SEASON_MAP.items():
                vals = [ts.iloc[i] for i in range(len(ts)) if ts.index[i].month in months]
                if vals:
                    month_avgs[s] = int(sum(vals) / len(vals))

            season_table_rows = []
            for s in ["春", "夏", "秋", "冬"]:
                avg = month_avgs.get(s, 0)
                bar = "█" * (avg // 10) + "░" * (10 - avg // 10)
                marker = " ← **ピーク**" if s == season else ""
                season_table_rows.append(f"| {s} | {bar} | {avg}/100 |{marker}")

            parts.append(f"""### 1️⃣ 検索ボリューム詳細分析

#### 基本指標
| 指標 | 値 | 評価 |
|---|---|---|
| **検索ピーク月** | {peak_idx.strftime('%Y年%m月')}（{season}） | 台湾人が最も旅行計画を立てる時期 |
| **最低月** | {trough_idx.strftime('%Y年%m月')} | 閑散期・オフシーズン |
| **12ヶ月平均** | {avg_val} / 100 | 市場の基礎的な関心度 |
| **変動幅** | {volatility} pt | {stability_label} |
| **直近3ヶ月傾向** | {recent_avg}/100 | {trend_direction}（平均比 {trend_vs_avg:+d}pt） |

#### 季節別平均検索ボリューム
| 季節 | 相対ボリューム | 平均値 |
|---|---|---|
{chr(10).join(season_table_rows)}

**✅ 旬のコンテンツ**: {season}（{season_kw}）
{season_advice}

> **💡 プロモーション最適タイミング**: ピーク月（{peak_idx.strftime('%m月')}）の **2〜3ヶ月前**（{peak_idx.strftime('%m月') and f"{(peak_idx.month - 3) % 12 + 1}〜{(peak_idx.month - 2) % 12 + 1}月"}）に
> 旅行博への出展・SNSキャンペーン・メディアプレスを集中させることで、
> 台湾人旅行者の計画フェーズに刺さる施策が打てます。
""")

    # ── Section 2: 台湾人ニーズ深層分析 ──
    all_queries: list[str] = []
    top_queries: list[str] = []
    rising_queries: list[str] = []

    if top_df is not None and not top_df.empty:
        top_queries = top_df["query"].tolist()
        all_queries += top_queries
    if rising_df is not None and not rising_df.empty:
        rising_queries = rising_df["query"].tolist()
        all_queries += rising_queries

    if all_queries:
        detected_cats = _detect_categories(all_queries)
        top_cats   = _detect_categories(top_queries)
        rising_cats = _detect_categories(rising_queries)

        # 全カテゴリの重要度スコア（TOP=2点、RISING=3点）
        cat_scores: dict[str, int] = {}
        for cat, qs in top_cats.items():
            cat_scores[cat] = cat_scores.get(cat, 0) + len(qs) * 2
        for cat, qs in rising_cats.items():
            cat_scores[cat] = cat_scores.get(cat, 0) + len(qs) * 3

        sorted_cats = sorted(cat_scores.items(), key=lambda x: x[1], reverse=True)

        # 定番 vs 急上昇で異なるカテゴリを抽出
        only_top     = set(top_cats) - set(rising_cats)
        only_rising  = set(rising_cats) - set(top_cats)
        both         = set(top_cats) & set(rising_cats)

        # キーワード翻訳テーブル
        kw_translations = {
            "溫泉": "温泉（日本式露天風呂）",
            "泡湯": "温泉入浴体験",
            "滑雪": "スキー・スノボ",
            "美食": "グルメ・おいしい食事",
            "住宿": "宿泊先",
            "行程": "旅行スケジュール・ルート",
            "自由行": "個人旅行・FIT",
            "賞楓": "紅葉狩り",
            "賞花": "花見・フラワーツーリズム",
            "免稅": "免税ショッピング",
            "購物": "ショッピング",
            "新幹線": "新幹線（台湾人に人気の体験）",
            "交通": "交通手段・アクセス",
            "景色": "絶景・風景撮影",
            "神社": "神社参拝",
            "體驗": "体験型アクティビティ",
        }

        trans_rows = []
        seen_kws: set[str] = set()
        for q in all_queries[:20]:
            if q in seen_kws:
                continue
            seen_kws.add(q)
            jp = kw_translations.get(q, "—")
            source = "🚀急上昇" if q in rising_queries else "🏆定番"
            trans_rows.append(f"| `{q}` | {jp} | {source} |")

        trend_comparison_rows = []
        if only_top:
            trend_comparison_rows.append(f"- **定番のみ（安定需要）**: {', '.join(f'`{c}`' for c in only_top)}")
        if only_rising:
            trend_comparison_rows.append(f"- **急上昇のみ（新トレンド）**: {', '.join(f'`{c}`' for c in only_rising)}")
        if both:
            trend_comparison_rows.append(f"- **両方に登場（最重要ニーズ）**: {', '.join(f'`{c}`' for c in both)}")

        parts.append(f"""### 2️⃣ 台湾人ニーズの深層分析

#### ニーズカテゴリ 重要度ランキング
（Top KW=2点・Rising KW=3点で加重スコア算出）

| 順位 | カテゴリ | スコア | 代表キーワード |
|---|---|---|---|
{chr(10).join(f"| {i+1} | **{cat}** | {score}pt | {'、'.join((detected_cats.get(cat, [])[:3]))} |" for i, (cat, score) in enumerate(sorted_cats[:6]))}

#### 定番需要 vs 新興トレンドの比較

{chr(10).join(trend_comparison_rows) if trend_comparison_rows else "- データ不足のため比較できません"}

> **💡 施策示唆**:
> 「両方に登場」カテゴリ = 台湾人が最も強く求めているニーズ → **最優先でコンテンツ化**
> 「急上昇のみ」カテゴリ = まだ競合が少ない先行者優位エリア → **先手コンテンツとして早期発信**
> 「定番のみ」カテゴリ = 既に競合が多い成熟需要 → **差別化要素を加えて訴求**

#### 関連キーワード 翻訳・解説
| 台湾語KW | 日本語訳 | 種別 |
|---|---|---|
{chr(10).join(trans_rows) if trans_rows else "| — | データなし | — |"}
""")

    # ── Section 3: プロモーションカレンダー ──
    if ts is not None and season:
        monthly_vals = []
        for i in range(len(ts)):
            monthly_vals.append((ts.index[i].month, int(ts.iloc[i])))

        cal_rows = []
        month_names = ["1月", "2月", "3月", "4月", "5月", "6月",
                       "7月", "8月", "9月", "10月", "11月", "12月"]
        for m, val in monthly_vals:
            intensity = "🔴 高" if val >= 70 else ("🟡 中" if val >= 40 else "🔵 低")
            action = ""
            if val >= 70:
                action = "プロモーション実施・旅行博出展"
            elif val >= 50:
                action = "SNS発信強化・メディア投稿"
            elif val >= 30:
                action = "コンテンツ準備・素材制作"
            else:
                action = "オフシーズン体験コンテンツ（冬の温泉など）訴求"
            cal_rows.append(f"| {month_names[m-1]} | {val}/100 | {intensity} | {action} |")

        parts.append(f"""### 3️⃣ 月別プロモーションカレンダー

| 月 | 検索量 | 優先度 | 推奨アクション |
|---|---|---|---|
{chr(10).join(cal_rows)}

> **読み方**: 🔴高優先月＝台湾人が最も予約・比較検討する時期。このタイミングに広告・露出を集中させる。
> 🔵低優先月はオフシーズン旅行の訴求（温泉・イルミネーション・祭りなど）でリピーターを狙う。
""")

    # ── Section 4: JR東日本 具体施策提言 ──
    recs: list[str] = []
    counter = 1

    # 季節に基づく施策
    if season and ts is not None:
        cols = [c for c in interest_df.columns if c != "isPartial"]
        if cols:
            ts2 = interest_df[cols[0]]
            pre_season = {
                "春": "1〜2月", "夏": "4〜5月", "秋": "7〜8月", "冬": "9〜10月"
            }.get(season, "2〜3ヶ月前")
            recs.append(
                f"{counter}. **プロモーション集中時期**: {season}のピーク月前（{pre_season}）に"
                f"旅行博・台湾メディアへのプレス発表・SNSキャンペーンを集中実施。"
                f"自治体との共同ブースで「JRパス × {keyword}体験」をアピール。"
            )
            counter += 1

    # カテゴリ別施策
    if all_queries:
        detected_cats2 = _detect_categories(all_queries)
        if "温泉・スパ" in detected_cats2:
            recs.append(
                f"{counter}. **温泉 × JRパス連携**: 沿線温泉地を「JRパスで行く台湾人の温泉旅行」として"
                "パッケージ提案。泡湯体験を推したハッシュタグキャンペーンでSNS拡散を狙う。"
            )
            counter += 1
        if "グルメ・食" in detected_cats2:
            recs.append(
                f"{counter}. **グルメツーリズム訴求**: 地域の食ブランド（和牛・海鮮・地酒等）を"
                "「産地直送グルメ」として前面に出したモデルコースを自治体提案に組み込む。"
                "台湾グルメメディアとのタイアップが効果的。"
            )
            counter += 1
        if "景観・自然" in detected_cats2:
            recs.append(
                f"{counter}. **フォトスポットSNS戦略**: Instagram・小紅書（RED）向けに絶景フォトスポット情報を"
                "繁体字で発信。台湾インフルエンサーを招聘してのファムトリップ実施を検討。"
            )
            counter += 1
        if "交通・アクセス" in detected_cats2:
            recs.append(
                f"{counter}. **JRパス利便性の強調**: 台湾人が「交通」を検索していることは"
                "アクセス不安の裏返し。JRパスで乗り換えなし・乗り放題の利便性をビジュアルで分かりやすく伝える。"
            )
            counter += 1
        if "宿泊" in detected_cats2:
            recs.append(
                f"{counter}. **宿泊施設との連携パッケージ**: 旅館・ホテルとJRパスのセット商品を"
                "台湾旅行会社（EZ Travel・Lion Travel等）に提案。予約ハードルを下げる。"
            )
            counter += 1

    if not recs:
        recs.append("トレンドデータが不足しているため、より一般的なキーワードでお試しください。")

    parts.append(f"""### 4️⃣ JR東日本 具体施策提言

{chr(10).join(recs)}

---

> 📌 **本レポートについて**
> Google Trends（台湾向け・過去12ヶ月）の実データとルールベース分析エンジンにより自動生成。
> キーワード翻訳・カテゴリ分類・季節判定・施策提言はデータに基づくテンプレート分析です。
> 提案書への引用時は「Google Trendsデータ分析（{keyword}、台湾向け）」と出典を明記してください。
""")

    return "\n".join(parts)


def generate_persona_analysis(
    keyword: str,
    tc_keyword: str,
    titles: list[str],
    top_kws: list[tuple],
) -> str:
    """
    PTT投稿タイトルと頻出KWからテンプレートベースで詳細ペルソナを生成する（Markdown）。
    外部API不要・即時実行・無制限。
    """
    total = len(titles)
    if total == 0:
        return "⚠️ 分析に必要な投稿データがありません。キーワードを変えてPTTから投稿を取得してください。"

    # ── 投稿パターン分析 ──
    _Q_WORDS  = ["請問", "想問", "怎麼", "如何", "哪裡", "哪個", "需要", "有人", "可以嗎", "有無"]
    _S_WORDS  = ["心得", "分享", "遊記", "記錄", "回報", "報告", "紀錄", "旅遊日記"]
    _P_WORDS  = ["行程", "計劃", "路線", "攻略", "規劃", "安排", "幾天", "幾日"]
    _G_WORDS  = ["家人", "朋友", "情侶", "親子", "全家", "媽媽", "爸爸", "小孩", "帶娃", "伴侶"]
    _FIT_WORDS = ["自由行", "自駕", "一個人", "獨旅", "背包"]
    _COST_WORDS = ["便宜", "省錢", "預算", "划算", "cost", "費用", "價格"]
    _LUXURY_WORDS = ["高級", "奢華", "精品", "商務", "頭等", "豪華"]
    _REPEAT_WORDS = ["再訪", "第二次", "第三次", "又來", "回訪", "常去"]
    _FIRST_WORDS  = ["第一次", "初次", "首次", "新手", "第一回"]

    q_count      = sum(1 for t in titles if any(w in t for w in _Q_WORDS))
    s_count      = sum(1 for t in titles if any(w in t for w in _S_WORDS))
    p_count      = sum(1 for t in titles if any(w in t for w in _P_WORDS))
    g_count      = sum(1 for t in titles if any(w in t for w in _G_WORDS))
    fit_count    = sum(1 for t in titles if any(w in t for w in _FIT_WORDS))
    cost_count   = sum(1 for t in titles if any(w in t for w in _COST_WORDS))
    luxury_count = sum(1 for t in titles if any(w in t for w in _LUXURY_WORDS))
    repeat_count = sum(1 for t in titles if any(w in t for w in _REPEAT_WORDS))
    first_count  = sum(1 for t in titles if any(w in t for w in _FIRST_WORDS))

    # ── 旅行スタイル判定 ──
    if q_count > total * 0.45:
        style_label = "情報収集型（初心者・計画段階）"
        style_desc  = "初めて訪問を検討している旅行者が多く、基本的な観光情報・アクセス情報・宿泊ガイドへのニーズが高い"
        style_action = "わかりやすい「入門コンテンツ」（アクセス図・モデル1泊2日プランなど）が最も効果的"
    elif s_count > total * 0.35:
        style_label = "体験共有型（経験者・リピーター）"
        style_desc  = "訪問経験があるリピーターが多く、穴場スポット・深堀り体験・新しい発見を求めている"
        style_action = "穴場スポット・季節限定コンテンツ・地元民おすすめの「玄人向け情報」が刺さる"
    elif p_count > total * 0.35:
        style_label = "計画立案型（アクティブプランナー）"
        style_desc  = "事前に詳細な行程を組む傾向が強く、効率的なルート・交通手段・滞在日数の最適化に関心が高い"
        style_action = "具体的なモデルコース（日数別・テーマ別）とJRパス活用ガイドが最も響く"
    else:
        style_label = "複合型（幅広い層が混在）"
        style_desc  = "初心者から経験者まで幅広い台湾人旅行者が関心を持っており、多様なコンテンツ需要がある"
        style_action = "入門・上級者両方に対応したコンテンツ階層化（入門ガイド + 深堀り情報）が有効"

    # ── セグメント分析 ──
    solo_pct    = f"{fit_count/total*100:.0f}%"  if total else "—"
    group_pct   = f"{g_count/total*100:.0f}%"    if total else "—"
    budget_pct  = f"{cost_count/total*100:.0f}%" if total else "—"
    luxury_pct  = f"{luxury_count/total*100:.0f}%" if total else "—"
    repeat_pct  = f"{repeat_count/total*100:.0f}%" if total else "—"
    first_pct   = f"{first_count/total*100:.0f}%"  if total else "—"

    # ── 主要カテゴリニーズ ──
    kw_text = [k for k, _ in top_kws]
    all_kws = kw_text + [t for t in titles[:30]]
    detected_cats = _detect_categories(all_kws)

    cat_need_rows = []
    cat_strategy_map = {
        "温泉・スパ":        "温泉体験パッケージをJRパスと組み合わせてセット販売。写真映えする露天風呂を前面に出す",
        "グルメ・食":        "地域の食ブランドを「産地直送」として訴求。食べ歩きMAP・グルメモデルルートを提供",
        "景観・自然":        "季節の絶景フォトスポット情報を繁体字SNSで発信。インフルエンサー招聘ファムトリップ実施",
        "ショッピング":      "沿線の免税対応店舗リストとJRパスを連携したショッピングガイドを台湾旅行会社に配布",
        "歴史・文化":        "着物・甲冑体験・伝統工芸ワークショップを組み込んだ文化体験パッケージの提案",
        "体験・アクティビティ": "スキー・農業・漁業体験など「台湾では体験できないこと」を差別化軸に据えたコンテンツ化",
        "交通・アクセス":    "JRパスの利便性を地図・動画で分かりやすく説明。台湾語での乗り換え案内コンテンツを整備",
        "宿泊":             "温泉旅館・ゲストハウス・ビジネスホテルの選択肢を台湾人好みの条件でフィルタリング提示",
        "季節・イベント":    "祭り・花火・紅葉など台湾人が熱狂するイベントを軸にした「このシーズンだけ」コンテンツ",
    }

    for cat, qs in detected_cats.items():
        strategy = cat_strategy_map.get(cat, "地域の独自コンテンツとして差別化訴求")
        cat_need_rows.append(
            f"| **{cat}** | `{'` `'.join(qs[:4])}` | {strategy} |"
        )

    # ── 旅行者ジャーニーマップ ──
    journey_data = [
        ("🔍 情報収集", "PTT・Dcard・Instagramで口コミ検索\nGoogle・Youtubeで旅行動画確認",
         "PTT投稿に返信する・ファムトリップ動画の拡散を促す"),
        ("📅 計画立案", "旅行サイト（EZ Travel・kkday等）で行程検討\n友人・家族とLINEで相談",
         "モデルルート・費用シミュレーターを台湾旅行サイトに掲載"),
        ("🎫 予約",     "Agoda・Booking.comで宿泊予約\nJRパス・航空券をオンラインで購入",
         "台湾旅行会社経由のJRパス販売チャネルを強化"),
        ("🚅 現地体験", "JRで移動・観光スポット巡り\n食事・温泉・ショッピングを楽しむ",
         "現地での多言語案内（繁体字）・QRコード案内を整備"),
        ("📸 情報発信", "Instagram・小紅書・PTTに旅行記を投稿\nUGC（ユーザー生成コンテンツ）として拡散",
         "ハッシュタグ設計・フォトスポット提案で自発的UGC投稿を促進"),
    ]

    journey_rows = []
    for phase, behavior, opportunity in journey_data:
        journey_rows.append(f"| {phase} | {behavior.replace(chr(10), '<br>')} | {opportunity} |")

    # ── 潜在的懸念・ニーズ ──
    concern_signals = []
    for t in titles:
        if "語言" in t or "語言不通" in t or "不懂日文" in t:
            concern_signals.append("言語バリア（日本語ができない不安）")
        if "安全" in t or "治安" in t:
            concern_signals.append("安全・治安への懸念")
        if "費用" in t or "預算" in t or "貴" in t or "貴不貴" in t:
            concern_signals.append("費用・コスト感への不安")
        if "交通" in t or "怎麼去" in t or "怎麼搭" in t:
            concern_signals.append("アクセス・交通手段の分からなさ")
        if "住宿" in t or "哪裡住" in t:
            concern_signals.append("宿泊先の選び方に迷い")

    concern_summary = list(dict.fromkeys(concern_signals))[:5]
    if not concern_summary:
        concern_summary = ["投稿数が少ないため潜在懸念の特定に十分なデータがありません"]

    concern_lines = [f"- {c}" for c in concern_summary]

    parts = [
        f"## 🎯 「{keyword}」ターゲットペルソナ 詳細分析レポート",
        f"*分析対象: PTT Japan_Travel板 投稿 {total} 件　|　分析日: 最新データ*",
        "",
        "---",
        "",
        "### 1️⃣ ターゲットペルソナ プロファイル",
        "",
        f"**旅行スタイル分類**: 🏷️ **{style_label}**",
        f"> {style_desc}",
        "",
        "#### 旅行者セグメント分析",
        "",
        "| セグメント | 比率 | 示唆 |",
        "|---|---|---|",
        f"| 個人・FIT旅行者 | {solo_pct} | 自分でプランを組む層。詳細情報・自由度の高い選択肢が必要 |",
        f"| グループ・家族旅行 | {group_pct} | 家族・カップル向けプランのニーズ。安全性・利便性重視 |",
        f"| 初訪問者 | {first_pct} | 基本情報への需要が高い。入門コンテンツで掴む |",
        f"| リピーター | {repeat_pct} | 新しい体験・穴場を求めている。上級者向け情報が響く |",
        f"| 節約志向 | {budget_pct} | 費用対効果を重視。JRパスの「お得感」訴求が有効 |",
        f"| 高品質志向 | {luxury_pct} | 体験の質を重視。ラグジュアリー旅館・グルメ訴求が刺さる |",
        "",
        f"**📌 コンテンツ戦略**: {style_action}",
        "",
        "---",
        "",
        "### 2️⃣ 関心カテゴリ × マーケティング戦略",
        "",
        "| 関心カテゴリ | 関連キーワード | 推奨施策 |",
        "|---|---|---|",
        *(cat_need_rows if cat_need_rows else ["| — | データ不足 | — |"]),
        "",
        "---",
        "",
        "### 3️⃣ 潜在的な懸念・不安（投稿タイトルから抽出）",
        "",
        "台湾人旅行者が{keyword}エリア旅行において持ちやすい懸念事項:",
        "",
        *concern_lines,
        "",
        "> **💡 対策**: 懸念事項に対する「答え」をコンテンツとして先回り提供する。",
        "> 例：アクセス不安 → 繁体字の乗り換え案内動画を作成 / 言語バリア → 多言語対応施設リストを整備",
        "",
        "---",
        "",
        "### 4️⃣ 台湾人旅行者 カスタマージャーニーマップ",
        "",
        "| フェーズ | 典型的な行動 | JR東日本/自治体の介入機会 |",
        "|---|---|---|",
        *journey_rows,
        "",
        "---",
        "",
        "### 5️⃣ 投稿データ詳細サマリー",
        "",
        f"| 指標 | 件数 | 比率 |",
        f"|---|---|---|",
        f"| 分析対象投稿数 | {total} 件 | 100% |",
        f"| 質問・情報収集系 | {q_count} 件 | {q_count/total*100:.0f}% |",
        f"| 体験共有・心得系 | {s_count} 件 | {s_count/total*100:.0f}% |",
        f"| 行程計画・攻略系 | {p_count} 件 | {p_count/total*100:.0f}% |",
        f"| グループ・家族旅行 | {g_count} 件 | {g_count/total*100:.0f}% |",
        "",
        f"**頻出キーワード（Top {len(top_kws)}語）**: "
        + "、".join(f"`{k}`（{c}回）" for k, c in top_kws),
        "",
        "---",
        "",
        "### 6️⃣ JR東日本 提案への落とし込み",
        "",
        f"1. **ターゲット定義**: {keyword}に興味を持つ台湾人は「{style_label}」が主軸。"
        "この層に刺さる訴求軸を提案書の冒頭に明示する",
        "",
        "2. **コンテンツ優先順位**: 上記カテゴリ分析の上位ニーズからコンテンツを優先的に整備する",
        "",
        "3. **懸念先回り**: 旅行者の不安要因（アクセス・言語・費用）を取り除くコンテンツが"
        "「最後の一押し」として予約転換率を高める",
        "",
        "4. **UGC活用**: PTTの投稿を「台湾人旅行者の生の声（VOC）」として提案資料に引用。"
        "「これだけの台湾人が話題にしている」という定量的な根拠として活用できる",
        "",
        "> 📌 本レポートはPTT投稿タイトルのパターン分析に基づくテンプレート生成です。"
        "提案書への引用時は「PTT Japan_Travel板 投稿分析」と出典を明記してください。",
    ]

    return "\n".join(parts)


# ============================================================
# Google Trends
# ============================================================


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_google_trends(keyword: str) -> tuple:
    """
    Google Trends（台湾向け、過去12ヶ月）でデータを取得する。

    Returns:
        (interest_df, related_queries_dict, error_type)
        error_type: None | "rate_limit" | "other:<message>"
    """
    try:
        from pytrends.request import TrendReq  # type: ignore

        pytrends = TrendReq(hl="ja-JP", tz=-540)
        pytrends.build_payload(
            kw_list=[keyword],
            geo="TW",
            timeframe="today 12-m",
        )
        interest_df = pytrends.interest_over_time()
        related = pytrends.related_queries()
        return interest_df, related, None
    except Exception as exc:
        msg = str(exc)
        if "429" in msg or "too many requests" in msg.lower() or "response code" in msg.lower():
            return None, None, "rate_limit"
        return None, None, f"other:{msg}"


# ============================================================
# PTT スクレイピング
# ============================================================


@st.cache_data(ttl=1800, show_spinner=False)
def scrape_ptt(keyword: str) -> tuple[pd.DataFrame, str | None]:
    """
    PTT Japan_Travel 板からキーワード関連の投稿を最大15件取得する。
    """
    url = f"{PTT_SEARCH_URL}?q={quote(keyword)}"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-TW,zh;q=0.9,ja;q=0.8,en;q=0.7",
        "Referer": "https://www.ptt.cc/bbs/Japan_Travel/index.html",
    }
    cookies = {"over18": "1"}

    try:
        resp = requests.get(url, headers=headers, cookies=cookies, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        posts: list[dict] = []

        for entry in soup.select(".r-ent")[:15]:
            title_a = entry.select_one(".title a")
            date_el  = entry.select_one(".date")
            push_el  = entry.select_one(".nrec span")
            if title_a:
                posts.append(
                    {
                        "ソース":   "PTT",
                        "タイトル": title_a.get_text(strip=True),
                        "URL":      "https://www.ptt.cc" + title_a.get("href", ""),
                        "日付":     date_el.get_text(strip=True) if date_el else "—",
                        "推薦":     push_el.get_text(strip=True) if push_el else "0",
                    }
                )

        if not posts:
            return pd.DataFrame(), "no_posts"
        return pd.DataFrame(posts), None

    except requests.exceptions.ConnectionError:
        return pd.DataFrame(), "connection_error"
    except requests.exceptions.Timeout:
        return pd.DataFrame(), "timeout"
    except Exception as exc:
        return pd.DataFrame(), f"other:{exc}"


# ============================================================
# ④ Google News RSS（台湾向け）
# ============================================================


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_google_news_taiwan(keyword: str) -> tuple[pd.DataFrame, str | None]:
    """
    Google News RSS（台湾向け）からキーワード関連記事を取得する。
    台湾人旅行者の「生の声」として活用可能なブログ・ニュース記事を自動収集する。
    """
    try:
        import feedparser  # type: ignore
    except ImportError:
        return pd.DataFrame(), "feedparser_not_installed"

    query = quote(f"{keyword} 日本旅遊")
    url = (
        f"https://news.google.com/rss/search"
        f"?q={query}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"
    )
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
    except requests.exceptions.ConnectionError:
        return pd.DataFrame(), "connection_error"
    except requests.exceptions.Timeout:
        return pd.DataFrame(), "timeout"
    except Exception as exc:
        return pd.DataFrame(), f"fetch_error:{exc}"

    try:
        feed = feedparser.parse(resp.text)
    except Exception as exc:
        return pd.DataFrame(), f"parse_error:{exc}"

    articles: list[dict] = []
    for entry in feed.entries[:20]:
        title = entry.get("title", "")
        link  = entry.get("link", "")
        pub   = entry.get("published", "")
        src   = entry.get("source", {}).get("title", "") if "source" in entry else ""

        # "タイトル - メディア名" 形式を分離
        if " - " in title and not src:
            parts = title.rsplit(" - ", 1)
            title = parts[0].strip()
            src   = parts[1].strip()

        articles.append({
            "タイトル": title,
            "ソース":   src or "Google News",
            "日付":     pub[:16] if pub else "—",
            "URL":      link,
        })

    if not articles:
        return pd.DataFrame(), "no_articles"
    return pd.DataFrame(articles), None


# ============================================================
# ④ YouTube 動画検索（YouTube Data API v3 使用、キー任意）
# ============================================================


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_youtube_taiwan(keyword: str, api_key: str) -> tuple[pd.DataFrame, str | None]:
    """
    YouTube Data API v3 で台湾人向け旅行動画を検索する。
    台湾人旅行者が投稿した動画タイトル・概要からニーズを読み取る。
    """
    query = quote(f"{keyword} 日本旅遊")
    url = (
        "https://www.googleapis.com/youtube/v3/search"
        f"?part=snippet&q={query}&regionCode=TW&relevanceLanguage=zh"
        f"&type=video&maxResults=10&key={api_key}"
    )

    try:
        resp = requests.get(url, timeout=15)
    except requests.exceptions.ConnectionError:
        return pd.DataFrame(), "connection_error"
    except requests.exceptions.Timeout:
        return pd.DataFrame(), "timeout"

    if resp.status_code == 400:
        return pd.DataFrame(), "invalid_key"
    if resp.status_code == 403:
        body = resp.json() if resp.content else {}
        reason = body.get("error", {}).get("errors", [{}])[0].get("reason", "")
        if reason == "quotaExceeded":
            return pd.DataFrame(), "quota_exceeded"
        return pd.DataFrame(), "forbidden"
    if resp.status_code != 200:
        return pd.DataFrame(), f"api_error:{resp.status_code}"

    items = resp.json().get("items", [])
    if not items:
        return pd.DataFrame(), "no_results"

    videos: list[dict] = []
    for item in items:
        snip = item.get("snippet", {})
        vid_id = item.get("id", {}).get("videoId", "")
        videos.append({
            "タイトル":     snip.get("title", ""),
            "チャンネル":   snip.get("channelTitle", ""),
            "投稿日":       snip.get("publishedAt", "")[:10],
            "概要":         snip.get("description", "")[:80],
            "URL":          f"https://www.youtube.com/watch?v={vid_id}",
        })

    return pd.DataFrame(videos), None


# ============================================================
# キーワード頻度分析 + WordCloud
# ============================================================


def _tokenize_jieba(text: str) -> list[str]:
    """jieba で単語分割。ImportError 時は regex にフォールバック。"""
    try:
        import jieba  # type: ignore
        return list(jieba.cut(text, cut_all=False))
    except ImportError:
        cjk   = re.findall(r"[\u4e00-\u9fff\u3040-\u30ff]{2,}", text)
        latin = re.findall(r"[a-zA-Z]{3,}", text)
        return cjk + [w.lower() for w in latin]


def analyze_top_keywords(
    titles: list[str], exclude_word: str, top_n: int = 5
) -> list[tuple[str, int]]:
    """タイトルリストから頻出キーワードを抽出して上位 top_n 件を返す。"""
    text = " ".join(titles)
    words = _tokenize_jieba(text)
    stop = _STOP_WORDS | {exclude_word} | set(exclude_word)
    filtered = [w.strip() for w in words if w.strip() and w.strip() not in stop and len(w.strip()) >= 2]
    return Counter(filtered).most_common(top_n)


def generate_wordcloud_fig(
    titles: list[str], font_path: str | None, exclude_word: str
) -> plt.Figure | None:
    """WordCloud を生成して matplotlib Figure を返す。"""
    try:
        from wordcloud import WordCloud  # type: ignore

        text     = " ".join(titles)
        words    = _tokenize_jieba(text)
        stop     = _STOP_WORDS | {exclude_word} | set(exclude_word)
        filtered = [w.strip() for w in words if w.strip() and w.strip() not in stop and len(w.strip()) >= 2]

        if not filtered:
            return None

        freq = Counter(filtered)
        wc_kwargs: dict = {
            "width": 900, "height": 400,
            "background_color": "white", "max_words": 100,
            "colormap": "tab10", "collocations": False,
            "prefer_horizontal": 0.85,
        }
        if font_path:
            wc_kwargs["font_path"] = font_path

        wc = WordCloud(**wc_kwargs)
        wc.generate_from_frequencies(dict(freq))

        fig, ax = plt.subplots(figsize=(11, 4.5))
        ax.imshow(wc, interpolation="bilinear")
        ax.axis("off")
        plt.tight_layout(pad=0)
        return fig
    except Exception:
        return None


# ============================================================
# UI ヘルパー関数
# ============================================================


def _insight_box(html: str) -> None:
    st.markdown(f'<div class="insight-box">{html}</div>', unsafe_allow_html=True)


def _analysis_box(result: str) -> None:
    """テンプレート分析結果を描画する。"""
    if result.startswith("⚠️") or result.startswith("❌"):
        st.error(result)
    else:
        st.markdown(
            '<div class="analysis-box">📊 <strong>データ分析レポート（テンプレートベース・API不要）</strong></div>',
            unsafe_allow_html=True,
        )
        st.markdown(result)


def _source_badge(source: str) -> str:
    if source == "PTT":
        return "🟠 PTT"
    return f"📌 {source}"


def _show_ptt_error(err: str) -> None:
    if err == "connection_error":
        st.warning("PTT への接続に失敗しました（ConnectionError）")
    elif err == "timeout":
        st.warning("PTT への接続がタイムアウトしました")
    elif err == "no_posts":
        st.info("該当する投稿が見つかりませんでした")
    else:
        st.warning(f"PTT エラー: {err}")


def _show_news_error(err: str | None) -> None:
    if err == "feedparser_not_installed":
        st.warning("feedparser がインストールされていません。`pip install feedparser` を実行してください。")
    elif err == "connection_error":
        st.warning("Google News への接続に失敗しました")
    elif err == "timeout":
        st.warning("Google News がタイムアウトしました。少し待ってから再試行してください。")
    elif err == "no_articles":
        st.info("該当する記事が見つかりませんでした。キーワードを変えてお試しください。")
    elif err:
        st.warning(f"Google News 取得エラー: {err}")


def _show_youtube_error(err: str | None) -> None:
    if err == "invalid_key":
        st.error("YouTube APIキーが無効です。Google Cloud Console で確認してください。")
    elif err == "quota_exceeded":
        st.warning("YouTube API の無料枠（10,000ユニット/日）を使い切りました。明日以降に再試行してください。")
    elif err == "forbidden":
        st.error("YouTube APIへのアクセスが拒否されました。APIキーの権限を確認してください。")
    elif err == "connection_error":
        st.warning("YouTube APIへの接続に失敗しました")
    elif err == "no_results":
        st.info("動画が見つかりませんでした。キーワードを変えてお試しください。")
    elif err:
        st.warning(f"YouTube APIエラー: {err}")


# ============================================================
# 提案書サマリー生成（プロフェッショナル版）
# ============================================================


def _generate_proposal_text(
    kw: str,
    tc_kw: str,
    jnto_df,           # yearly DataFrame or None
    jnto_monthly_df,   # monthly DataFrame or None
    top_df_t,          # Trends top DataFrame or None
    rising_df_t,       # Trends rising DataFrame or None
    combined_titles: list[str],
    top_kws: list[tuple],
    ptt_count: int,
    news_count: int,
) -> str:
    """
    全取得データを統合し、JR東日本が自治体向け提案書で実際に使える
    プロフェッショナルな市場調査レポート（Markdown形式）を生成する。

    構成：
      0. エグゼクティブサマリー（1枚で全体把握）
      1. 台湾インバウンド市場の現状と回復動向
      2. 台湾人旅行者のニーズ分析
      3. 現状の課題・リスク分析
      4. 提言（短期・中期・長期アクション）
      5. JR東日本との連携スキーム案
      6. 期待効果の試算
      7. 付録（データ根拠・調査概要）
    """
    import datetime
    # strftime の %-m / %-d はWindowsで動作しないため f-string で直接フォーマット
    try:
        d = datetime.date.today()
        today = f"{d.year}年{d.month}月{d.day}日"
    except Exception:
        today = "2026年3月"

    parts = []

    # ══════════════════════════════════════════════════════════
    # ヘッダー
    # ══════════════════════════════════════════════════════════
    parts.append(f"""# {kw} 台湾インバウンド市場調査レポート
**作成**: JR東日本 台湾事業開発チーム　｜　**調査日**: {today}
**調査対象**: 台湾人旅行者の「{kw}」への関心・トレンド・ニーズ
**目的**: {kw}への台湾人誘客に向けた提案書の事実基盤整備

---
""")

    # ══════════════════════════════════════════════════════════
    # 0. エグゼクティブサマリー
    # ══════════════════════════════════════════════════════════
    # データから主要数値を抽出
    latest_yr, latest_val, prev_yr, prev_val, recovery = None, None, None, None, None
    peak_month, low_month = None, None
    yoy_rate = None

    if jnto_df is not None and not jnto_df.empty:
        latest = jnto_df.iloc[-1]
        latest_yr  = int(latest["年"])
        latest_val = int(latest["台湾人訪問者数（人）"])
        prev_row = jnto_df[jnto_df["年"] == latest_yr - 1]
        if not prev_row.empty:
            prev_val = int(prev_row.iloc[0]["台湾人訪問者数（人）"])
            if prev_val > 0:
                yoy_rate = (latest_val - prev_val) / prev_val * 100
        pre19 = jnto_df[jnto_df["年"] == 2019]
        if not pre19.empty:
            base = int(pre19.iloc[0]["台湾人訪問者数（人）"])
            if base > 0:
                recovery = latest_val / base * 100

    if jnto_monthly_df is not None and not jnto_monthly_df.empty:
        mo_avg = jnto_monthly_df.groupby("月")["訪問者数"].mean()
        peak_month = int(mo_avg.idxmax())
        low_month  = int(mo_avg.idxmin())

    # 主要ニーズカテゴリ抽出
    kw_text_list = [k for k, _ in top_kws]
    detected_cats_all = _detect_categories(kw_text_list + combined_titles[:50])
    top_cats = sorted(detected_cats_all.items(), key=lambda x: len(x[1]), reverse=True)[:3]
    top_cat_str = " / ".join(c for c, _ in top_cats) if top_cats else "データ収集中"

    # 急上昇KW（不適切なキーワードを除外してから使用）
    rising_kws_raw: list[str] = []
    if rising_df_t is not None and not rising_df_t.empty:
        rising_kws_raw = [str(q) for q in rising_df_t["query"].tolist()[:20]]
    rising_kws_list = _filter_promo_kws(rising_kws_raw)[:6]  # フィルタ後に上位6件
    filtered_count = len(rising_kws_raw) - len(_filter_promo_kws(rising_kws_raw))  # 除外件数

    # 訪問者数の概況文
    if latest_yr and latest_val:
        visitor_summary = f"台湾人訪問者数 **{latest_val:,}人**（{latest_yr}年）"
        if yoy_rate is not None:
            trend_word = "増加中" if yoy_rate > 0 else "減少傾向"
            visitor_summary += f"、前年比 **{yoy_rate:+.1f}%** で{trend_word}"
        if recovery:
            visitor_summary += f"、コロナ前比 **{recovery:.0f}%** まで回復"
    else:
        visitor_summary = "JNTOデータを取得すると具体的な訪問者数が表示されます"

    promo_month = (peak_month - 3) if peak_month and peak_month > 3 else (
        (peak_month + 9) if peak_month else None
    )

    exec_summary_lines = [
        "## 0. エグゼクティブサマリー",
        "",
        "> **本レポートの結論**: 台湾人旅行者の「{kw}」への潜在需要は確実に存在しており、"
        "的確なプロモーション施策と受入体制整備により、インバウンド消費の大幅な拡大が見込まれる。"
        "JR東日本との連携によりアクセス改善とコンテンツ発信を同時に実現することで、"
        "他県との差別化を図れる好機にある。".format(kw=kw),
        "",
        "### ■ 主要ファクト（本調査で判明した事実）",
        "",
        f"| 指標 | 数値・状況 |",
        f"|---|---|",
        f"| 台湾人訪問者数 | {visitor_summary} |",
    ]

    if peak_month:
        exec_summary_lines.append(f"| 訪日ピーク月 | {peak_month}月（プロモーション最適期: {promo_month}月頃） |")
    if top_cat_str:
        exec_summary_lines.append(f"| 主要関心テーマ | {top_cat_str} |")
    if rising_kws_list:
        # フィルタ済みの観光関連KWのみ（災害・事故ワードは除外済み）
        exec_summary_lines.append(f"| 急上昇キーワード（観光関連） | {' / '.join(rising_kws_list[:4])} |")
    elif rising_kws_raw:
        exec_summary_lines.append("| 急上昇キーワード | 観光関連KWなし（取得されたKWは非観光関連のため除外） |")
    exec_summary_lines.append(f"| 分析データ総量 | PTT投稿・ニュース記事 計{len(combined_titles)}件、Google Trends 12ヶ月分 |")

    # ── スポットデータから代表スポット・JR連携を取得（アクション3の具体化用）──
    spots_for_kw = get_inbound_spots(kw)
    top_spot_name = spots_for_kw[0]["スポット名"] if spots_for_kw else f"{kw}の代表観光地"
    top_spot_jr   = spots_for_kw[0].get("最寄り駅・JR連携", "JR沿線") if spots_for_kw else "JR沿線"
    course_spots  = [s["スポット名"] for s in (spots_for_kw or [])[:4]]
    course_str    = " → ".join(course_spots) if len(course_spots) >= 2 else f"{kw}観光スポット巡り"

    action1_timing = f"{promo_month}月" if promo_month else "旅行ピークの3ヶ月前"
    action1_peak   = f"{peak_month}月" if peak_month else "訪日ピーク月"

    # アクション2: 観光関連KWがある場合のみ具体的KWを使用
    if rising_kws_list:
        target_kw = rising_kws_list[0]
        action2_lines = [
            f"2. **繁体字SNSコンテンツの先行整備（テーマ: 「{target_kw}」）**",
            f"   - *なぜ今か*: 「{target_kw}」はGoogle Trendsで検索急増中にもかかわらず、台湾語の旅行コンテンツがまだ少ない。",
            f"     今コンテンツを整備すれば低コストで検索上位を獲得できる（先行者優位）。",
            f"   - *具体的な実施内容*: 「{target_kw} × {kw}旅行」を軸にしたInstagram投稿・ブログ記事を",
            f"     繁体字で5〜10本作成し、台湾旅行系インフルエンサーに配布または自社SNSから発信する。",
        ]
    else:
        action2_lines = [
            f"2. **{kw}の定番見どころを繁体字SNSコンテンツで整備**",
            f"   - *なぜ今か*: 台湾語の{kw}旅行コンテンツは絶対数が少なく、今から整備すれば検索・SNSの上位を取れる。",
            f"   - *具体的な内容*: {top_spot_name}など台湾人に刺さるスポットを",
            f"     繁体字でInstagram・YouTube・旅行ブログに発信。季節ごとの絶景写真を中心に据える。",
        ]

    action3_lines = [
        f"3. **JRパス × {kw} モデルコースを台湾語で作成・主要旅行サイトに掲載**",
        f"   - *なぜ今か*: 台湾人FIT旅行者の最大障壁は「JRでどう行くかわからない」こと。",
        f"     具体的な行程が示されることで、計画から予約への転換率が大幅に向上する。",
        f"   - *コース例（2〜3泊）*: 東京/成田 →（新幹線・特急）→ {kw} → {course_str}",
        f"     （起点: {top_spot_jr}）",
        f"   - *掲載交渉先*: 雄獅旅行・可樂旅遊の{kw}特集ページ、EZ Travel・kkday台湾版",
        f"     に「{kw} {action1_peak}旅行完全攻略」として掲載を依頼する。",
    ]

    exec_summary_lines += [
        "",
        "### ■ 今すぐ着手すべき最優先アクション（本提案の核心）",
        "",
        f"1. **{action1_timing}までに台湾旅行会社向けFAM（現地視察）を実施**",
        f"   - *なぜ今か*: {action1_peak}の訪日ピークから逆算すると、旅行会社がパッケージ商品を",
        f"     組成・カタログ掲載するには最低3ヶ月のリードタイムが必要。今実施しなければ今シーズンを逃す。",
        f"   - *メカニズム*: FAMに参加した商品担当者（各社2〜3名、計10〜15名規模）が{kw}を実体験することで、",
        f"     ①自分の言葉でパッケージを商品化できる、②社内稟議で現地の魅力を具体的に説明できる、",
        f"     ③帰国後に旅行会社のSNS・ニュースレターで自発的に発信が生まれる——という3つの効果が連鎖する。",
        f"     過去事例では1回のFAM（招聘費用: 100〜300万円）で3〜5社への商品化、",
        f"     年間500〜2,000人規模の誘客に繋がることが多い。",
        "",
        *action2_lines,
        "",
        *action3_lines,
        "",
        "---",
        "",
    ]
    parts.append("\n".join(exec_summary_lines))

    # ══════════════════════════════════════════════════════════
    # 1. 台湾インバウンド市場の現状
    # ══════════════════════════════════════════════════════════
    market_lines = [
        "## 1. 台湾インバウンド市場の現状と{kw}の位置づけ".format(kw=kw),
        "",
        "### 1-1. 日本全体の台湾市場規模",
        "",
        "| 指標 | 数値 | 出典 |",
        "|---|---|---|",
        "| 訪日台湾人旅行者数（2024年） | 約 **589万人**（年間） | JNTO |",
        "| 訪日外国人全体に占める台湾の割合 | 約 **15%**（韓国・中国に次ぐ3位） | JNTO |",
        "| 台湾人1人当たり旅行消費単価 | 約 **148,000円** | 観光庁 訪日外客消費動向調査 2023 |",
        "| 訪日台湾人の総消費額（推計） | 約 **8,700億円** | 上記2指標から試算 |",
        "",
        "> **解釈**: 台湾は「一人当たり単価の高さ」「リピーター比率の高さ」「個人旅行（FIT）比率の高さ」の3点で、"
        "インバウンドの質的観点からも最優先すべき市場である。",
        "",
        "### 1-2. {kw}への台湾人訪問者数の推移".format(kw=kw),
        "",
    ]

    if jnto_df is not None and not jnto_df.empty:
        market_lines.append("| 年 | 台湾人訪問者数 | 前年比 |")
        market_lines.append("|---|---|---|")
        prev_v = None
        for _, row in jnto_df.tail(6).iterrows():
            yr_r = int(row["年"])
            vr   = int(row["台湾人訪問者数（人）"])
            if prev_v and prev_v > 0:
                yoy = (vr - prev_v) / prev_v * 100
                market_lines.append(f"| {yr_r}年 | {vr:,}人 | {yoy:+.1f}% |")
            else:
                market_lines.append(f"| {yr_r}年 | {vr:,}人 | — |")
            prev_v = vr

        if recovery is not None:
            recovery_comment = (
                "コロナ前を大幅に上回っており、市場の成熟と新たな需要層の拡大が示唆される" if recovery > 110
                else "コロナ前水準をほぼ回復しており、今後は横ばいか微増の局面" if recovery > 95
                else f"コロナ前の{recovery:.0f}%水準まで回復。完全回復には旅行会社との連携強化が必要"
            )
            market_lines.append(f"\n> **コロナ前（2019年）比 {recovery:.0f}%**: {recovery_comment}。")
    else:
        market_lines.append("*（タブ1でJNTOデータを取得すると年別推移表が自動挿入されます）*")

    if peak_month:
        month_names_jp = {1:"1月", 2:"2月", 3:"3月", 4:"4月", 5:"5月", 6:"6月",
                          7:"7月", 8:"8月", 9:"9月", 10:"10月", 11:"11月", 12:"12月"}
        season_map_local = {12:"冬", 1:"冬", 2:"冬", 3:"春", 4:"春", 5:"春",
                            6:"夏", 7:"夏", 8:"夏", 9:"秋", 10:"秋", 11:"秋"}
        peak_season = season_map_local.get(peak_month, "")
        low_season  = season_map_local.get(low_month, "") if low_month else ""

        market_lines += [
            "",
            "### 1-3. 月別訪問パターン（季節性分析）",
            "",
            f"- **ピーク月**: {month_names_jp.get(peak_month, str(peak_month)+'月')}（{peak_season}）"
            f" → 台湾の連休・長期休暇と一致。**プロモーション投下は{promo_month}月頃が最適**",
            f"- **閑散月**: {month_names_jp.get(low_month, str(low_month)+'月')}（{low_season}）"
            " → 閑散期対策として「台湾では体験できない{low_season}コンテンツ」を訴求することで平準化が可能".format(low_season=low_season),
            f"- **プロモーションカレンダー提言**: 繁忙期の**2〜3ヶ月前**（約{promo_month}月）に"
            "旅行博・旅行会社向け商談・SNSキャンペーンを集中実施",
        ]

    market_lines += ["", "---", ""]
    parts.append("\n".join(market_lines))

    # ══════════════════════════════════════════════════════════
    # 2. 台湾人旅行者のニーズ分析
    # ══════════════════════════════════════════════════════════
    needs_lines = [
        f"## 2. 台湾人旅行者のニーズ分析（{kw}に対する関心）",
        "",
        "### 2-1. 主要関心カテゴリ（Google Trends + PTT/ニュース統合分析）",
        "",
    ]

    # カテゴリ分析
    cat_strategy_detail = {
        "温泉・スパ":        ("💰 高単価・高満足度", "台湾人が最も支出を惜しまないカテゴリ。「日本の温泉」は台湾人の憧れ体験であり、高級旅館の宿泊は客単価向上に直結する"),
        "グルメ・食":        ("🍽 訪日動機 Top3", "「食べるために日本へ行く」台湾人は多く、地域の食ブランドは最強の集客コンテンツ。産地直送・職人体験を組み込むと差別化できる"),
        "景観・自然":        ("📸 SNS拡散力 最大", "台湾人は写真・動画撮影目的での旅行が多い。絶景スポットの1枚がSNSで数万人に拡散され、観光地の認知が爆発的に広がる"),
        "ショッピング":      ("🛍 消費単価向上", "ドラッグストア・酒造・工芸品等の地域特産品購入。インバウンド消費の中でも購買転換率が高い"),
        "歴史・文化":        ("🎭 体験型 差別化", "着物・甲冑・伝統工芸など「台湾では体験できないこと」への需要。体験型コンテンツは客単価が高く、リピート動機にもなる"),
        "体験・アクティビティ": ("❄ 季節限定 希少性", "スキー・農業・漁業体験など季節限定の体験は「今しか行けない」感を生み出し、旅行計画の優先度を上げる"),
        "交通・アクセス":    ("🚅 JR連携 最重要", "「どうやって行くか」の情報不足がそのまま旅行機会の損失になる。JRパスを使った台湾語アクセスガイドは必須"),
        "宿泊":             ("🏨 滞在日数に直結", "宿泊施設の多様性（旅館・ゲストハウス・ホテル）と台湾人好みの選択肢提示が、滞在日数延長に直結する"),
        "季節・イベント":    ("🎆 旬の訴求力", "祭り・花火・紅葉など季節限定イベントは「これのために行く」動機を形成。最も強力な旅行動機づけコンテンツ"),
        "テーマパーク":      ("👨‍👩‍👧 家族旅行 大型需要", "家族連れ台湾人の訪日動機として非常に強力。近隣テーマパークとのセットプランが有効"),
    }

    if detected_cats_all:
        needs_lines += [
            "| 関心カテゴリ | 特性 | 戦略インサイト | 関連キーワード |",
            "|---|---|---|---|",
        ]
        for cat, kws_in_cat in sorted(detected_cats_all.items(), key=lambda x: len(x[1]), reverse=True):
            char_label, insight = cat_strategy_detail.get(cat, ("—", "地域固有の訴求ポイントを整理して発信する"))
            needs_lines.append(
                f"| **{cat}** | {char_label} | {insight} | {' / '.join(kws_in_cat[:3])} |"
            )
    else:
        needs_lines.append("*（タブ2・3でデータを取得するとカテゴリ分析が表示されます）*")

    needs_lines += [
        "",
        "### 2-2. 台湾人旅行者の旅行スタイル傾向",
        "",
        "| 旅行スタイル | 傾向 | {kw}への示唆 |".format(kw=kw),
        "|---|---|---|",
        "| **FIT（個人旅行）** | 台湾人訪日の約70%がFIT。自分でルートを組む | JRパスの使い方・モデルルートを台湾語で発信することが最重要 |",
        "| **グループ・家族旅行** | 連休（春節・清明節・国慶節）に集中 | 子ども向けアクティビティ・移動の容易さをアピール |",
        "| **リピーター** | 訪日10回超の超リピーターも存在。「まだ行っていない地方」を探している | 「東京ではない日本」として地方の個性を前面に押し出す |",
        "| **富裕層・高品質志向** | 旅館の個室露天風呂・料亭・伝統体験に高い支出意欲 | 高付加価値プランの造成と旅行会社への販路確立が収益に直結 |",
        "",
        "### 2-3. 情報収集・予約行動の特徴",
        "",
        "台湾人旅行者は**複数の情報源を横断的に参照**してから予約する傾向がある:",
        "",
        "| フェーズ | 主要メディア | {kw}としての対応策 |".format(kw=kw),
        "|---|---|---|",
        "| 発見・関心 | PTT・Instagram・YouTube・小紅書 | インフルエンサー招聘でUGCを生成、繁体字ハッシュタグ設計 |",
        "| 情報収集 | Google検索・旅行ブログ・Yahoo奇摩旅遊 | SEO対応の繁体字コンテンツをWebに整備 |",
        "| 比較・検討 | Dcard・友人LINE | 台湾旅行会社へのパンフレット提供・旅行博への出展 |",
        "| 予約 | Agoda・Booking.com・JRパス公式 | 台湾旅行会社との販売代理契約・JRパス連携プランの登録 |",
        "| 帰国後発信 | PTT・Instagram・小紅書 | フォトスポット整備・ハッシュタグ設計でUGC拡散を促進 |",
        "",
        "---",
        "",
    ]
    parts.append("\n".join(needs_lines))

    # ══════════════════════════════════════════════════════════
    # 3. 現状の課題・リスク分析
    # ══════════════════════════════════════════════════════════
    issue_lines = [
        f"## 3. 現状の課題・リスク分析",
        "",
        "### 3-1. {kw}が台湾市場で直面している構造的課題".format(kw=kw),
        "",
        "| # | 課題 | 深刻度 | 背景・詳細 |",
        "|---|---|---|---|",
        f"| 1 | **認知度の絶対的不足** | 🔴 高 | 台湾人の旅行先として「{kw}」が名前で語られることが少ない。「東北」「北陸」という広域概念では認識されていても、{kw}単独のブランドが確立していない |",
        f"| 2 | **アクセス情報の台湾語コンテンツ不足** | 🔴 高 | JRパスで{kw}にどう行くか、乗り換え情報・所要時間の台湾語案内がほぼ存在しない。アクセス不安が旅行計画断念の最大要因 |",
        f"| 3 | **SNS上での情報量の少なさ** | 🟠 中 | InstagramやPTTで「{kw}」を検索しても投稿数が少なく、旅行先として参考にする情報が得にくい |",
        f"| 4 | **旅行会社への商品化が不十分** | 🟠 中 | 台湾の旅行会社（雄獅旅行・可樂旅遊等）のパンフレットに{kw}を含む商品が少ない |",
        f"| 5 | **受入体制（多言語化）** | 🟡 低〜中 | 飲食店・観光施設での繁体字メニュー・案内不足が、訪問後の満足度を下げリピート阻害要因となっている |",
        "",
        "### 3-2. 機会とリスクのマトリクス",
        "",
        "| | 内部（{kw}側） | 外部（市場・競合） |".format(kw=kw),
        "|---|---|---|",
        "| **機会** | 未開拓エリアゆえのブルーオーシャン、地域固有の体験コンテンツ、JR線沿いの交通利便性 | 円安継続による訪日コスト低下、台湾人の地方旅行志向の高まり、LCC就航による地方空港への直行便増加 |",
        "| **脅威** | コンテンツ開発・受入整備への予算・人員不足、繁体字コンテンツ制作ノウハウ不足 | 他県・他地域との誘客競争激化、インフルエンサー依存によるトレンド変動リスク、オーバーツーリズムによる地域住民との摩擦 |",
        "",
    ]

    # 急上昇KWからリスクを読む
    if rising_kws_list:
        issue_lines += [
            "### 3-3. トレンドから読む「先手を打つべき領域」",
            "",
            "以下は台湾でのGoogle検索で急上昇しているキーワードであり、"
            "まだ情報供給が需要に追いついていない「先手コンテンツ領域」を示す:",
            "",
        ]
        for i, rk in enumerate(rising_kws_list[:5], 1):
            issue_lines.append(f"{i}. **{rk}**: 検索急増中 → 繁体字コンテンツを早期に整備することで検索上位を獲得できる機会")

    issue_lines += ["", "---", ""]
    parts.append("\n".join(issue_lines))

    # ══════════════════════════════════════════════════════════
    # 4. 提言（短期・中期・長期）
    # ══════════════════════════════════════════════════════════
    action_lines = [
        "## 4. 提言：具体的アクションプラン",
        "",
        "> **基本方針**: 「認知 → 関心 → 計画 → 予約 → 訪問 → 発信」の台湾人旅行者ジャーニー全体に介入し、"
        "各フェーズで{kw}が正しく選ばれる仕組みを構築する。".format(kw=kw),
        "",
        "### 短期アクション（0〜3ヶ月以内）",
        "",
        "| Priority | アクション | 担当 | 効果 |",
        "|---|---|---|---|",
        f"| ⭐⭐⭐ | **台湾語版「{kw}旅行ガイド」ページをWebに公開** | {kw}観光協会 + JR東日本 | FIT旅行者の情報収集ニーズに直接対応、Google検索流入を獲得 |",
        f"| ⭐⭐⭐ | **JRパスを使った{kw}アクセスガイドを繁体字で動画化（YouTube/IG Reels）** | JR東日本 | 最大の訪問障壁「アクセス不安」を払拭 |",
        f"| ⭐⭐ | **台湾人インフルエンサー（フォロワー5〜10万人規模）を{kw}に招聘** | {kw}観光協会 | PTT・Instagram・YouTubeで一次UGCを生成し認知を急速に拡大 |",
        f"| ⭐⭐ | **Google検索急上昇キーワードに対応した繁体字ブログ記事を5〜10本作成・公開** | JR東日本コンテンツチーム | 検索需要に応じたSEOコンテンツでオーガニック流入獲得 |",
        "",
        "### 中期アクション（3〜12ヶ月）",
        "",
        "| Priority | アクション | 担当 | 効果 |",
        "|---|---|---|---|",
        "| ⭐⭐⭐ | **台湾旅行博（台北国際旅遊展）への出展（例年10月開催）** | JR東日本 + 自治体 | 台湾の旅行会社・消費者へ直接リーチ、商談機会の獲得 |",
        f"| ⭐⭐⭐ | **台湾大手旅行会社（雄獅・可樂・易飛網等）へのFAM（視察招聘）と商品化交渉** | JR東日本営業 | {kw}を含むパッケージ商品をカタログ掲載、流通販売チャネルを確立 |",
        f"| ⭐⭐ | **JRパス × {kw}を組み合わせたモデルコース（2〜4泊）を3〜5コース造成・旅行サイト掲載** | JR東日本 | 旅行計画段階での{kw}選択肢化 |",
        f"| ⭐⭐ | **{kw}内の主要観光施設・飲食店での繁体字メニュー・案内整備支援** | {kw}観光協会 | 訪問後満足度を向上させリピーター化・口コミUGC増加に直結 |",
        "| ⭐ | **小紅書（中国版Instagram、台湾でも広く使われる）への公式アカウント開設** | JR東日本または自治体 | 20〜30代台湾女性旅行者への直接発信チャネルを構築 |",
        "",
        "### 長期アクション（1〜3年）",
        "",
        "| Priority | アクション | 担当 | 効果 |",
        "|---|---|---|---|",
        "| ⭐⭐⭐ | **台湾人特化の体験型コンテンツ（温泉・食・農業・工芸）の商品化と年間販売スキーム構築** | 自治体 + JR東日本 | 「観光地」から「体験目的地」へのブランドシフト |",
        "| ⭐⭐ | **台湾の旅行会社との継続的なパートナーシップ協定（覚書・定例商談）の締結** | JR東日本・自治体 | 中長期の安定的な誘客ルートを確保 |",
        "| ⭐⭐ | **台湾直行便路線・LCC誘致への働きかけ（地方空港活用）** | 自治体・県・国土交通省 | アクセス利便性の抜本的改善により訪問コストを大幅低減 |",
        "| ⭐ | **訪日台湾人の満足度調査・KPI定点観測（年2回）の制度化** | 観光協会 | PDCAサイクルによる施策効果の継続的改善 |",
        "",
        "---",
        "",
    ]
    parts.append("\n".join(action_lines))

    # ══════════════════════════════════════════════════════════
    # 5. JR東日本との連携スキーム
    # ══════════════════════════════════════════════════════════
    jr_lines = [
        "## 5. JR東日本との連携スキーム（提案）",
        "",
        "### 5-1. JRパスを活用した{kw}誘客モデル".format(kw=kw),
        "",
        "JR東日本は「台湾からの訪日客」にとってアクセスの実質的な担い手であり、"
        "以下のスキームで自治体と連携することで相互にWin-Winの関係を構築できる:",
        "",
        "```",
        "【台湾人旅行者の目線で見た連携モデル】",
        "",
        "  台湾 ─→ 成田/羽田/仙台/新千歳など ─→ JR東日本新幹線・特急 ─→ {kw}".format(kw=kw),
        "  　　　　　　↑　　　　　　　　　　　　　 ↑　　　　　　　　　　　　↑",
        "  　台湾旅行会社が販売　　　　JRパスを旅行会社が代理販売　　　観光協会がガイド提供",
        "```",
        "",
        "### 5-2. JR東日本が自治体に提供できる価値",
        "",
        "| 提供価値 | 具体的内容 |",
        "|---|---|",
        "| **流通チャネル** | 台湾旅行会社へのJRパス販売ルート（既存チャネル）を活用し、{kw}を含む商品を流通させる |".format(kw=kw),
        "| **コンテンツ発信** | JR東日本の台湾向け公式SNS・Webサイトに{kw}の観光情報を掲載 |".format(kw=kw),
        "| **プロモーション機会** | 台湾旅行博・旅行会社向け商談会への合同出展 |",
        "| **データ提供** | 本ダッシュボードのようなGoogle Trends・JNTOデータ分析の共有 |",
        "| **造成支援** | JRパスを使ったモデルコース設計・旅行会社への商品化提案 |",
        "",
        "### 5-3. 自治体が準備すべき前提条件",
        "",
        "JR東日本との連携を最大化するために、自治体は以下を先行して整備することが求められる:",
        "",
        f"1. **受入体制の最低限の整備**: 主要観光施設での繁体字案内・Wi-Fi整備・交通系ICカード利用可否の確認",
        f"2. **コンテンツの撮影素材確保**: 繁体字SNS用の高解像度写真・短尺動画（季節別・テーマ別）の制作",
        f"3. **地域内の「推せるコンテンツ」の選定と磨き込み**: 台湾人が喜ぶ体験・食・景観のうち、実際に提供できる「自信を持って案内できる3〜5コンテンツ」の確定",
        "",
        "---",
        "",
    ]
    parts.append("\n".join(jr_lines))

    # ══════════════════════════════════════════════════════════
    # 6. 期待効果の試算
    # ══════════════════════════════════════════════════════════
    if latest_val and latest_yr:
        # 試算のベースライン
        target_5pct  = int(latest_val * 1.05)
        target_10pct = int(latest_val * 1.10)
        target_20pct = int(latest_val * 1.20)
        spend_per_person = 148000  # 観光庁 訪日外客消費動向調査 2023

        eco_5pct  = int(target_5pct  * spend_per_person / 100000000)  # 億円
        eco_10pct = int(target_10pct * spend_per_person / 100000000)
        eco_20pct = int(target_20pct * spend_per_person / 100000000)

        econ_lines = [
            "## 6. 期待効果の試算",
            "",
            "> **試算の前提**: 台湾人1人当たり旅行消費単価 148,000円（観光庁 訪日外客消費動向調査 2023年）を使用。",
            f"> **ベースライン**: {latest_yr}年の台湾人訪問者数 {latest_val:,}人",
            "",
            "| シナリオ | 目標訪問者数 | 増加率 | 期待消費額（試算）| 主な施策 |",
            "|---|---|---|---|---|",
            f"| **現状維持** | {latest_val:,}人 | ±0% | 約{int(latest_val * spend_per_person / 100000000)}億円 | 現行施策の継続 |",
            f"| **保守的成長** | {target_5pct:,}人 | +5% | 約{eco_5pct}億円 | SNS整備・インフルエンサー招聘のみ |",
            f"| **現実的目標** | {target_10pct:,}人 | +10% | 約{eco_10pct}億円 | 短期〜中期アクション実施 |",
            f"| **積極的拡大** | {target_20pct:,}人 | +20% | 約{eco_20pct}億円 | 全アクションプラン実施 + LCC活用 |",
            "",
            "> 📌 **留意事項**: 上記は試算であり、実際の消費額は滞在日数・旅行スタイル・施策効果によって変動する。"
            "また消費額全体がそのまま当該地域に落ちるわけではなく、地域内消費比率（通常50〜70%）を掛けた額が実質的な地域経済効果となる。",
            "",
            "---",
            "",
        ]
    else:
        econ_lines = [
            "## 6. 期待効果の試算",
            "",
            "*（タブ1でJNTOデータを取得すると経済効果試算が自動生成されます）*",
            "",
            "---",
            "",
        ]
    parts.append("\n".join(econ_lines))

    # ══════════════════════════════════════════════════════════
    # 7. 付録（データ根拠）
    # ══════════════════════════════════════════════════════════
    appendix_lines = [
        "## 7. 付録：本レポートのデータ根拠",
        "",
        "| データ | ソース | 取得方法 | 信頼性 |",
        "|---|---|---|---|",
        f"| 台湾人訪問者数（年別・月別） | JNTO（日本政府観光局）公式統計 | 自動取得（Excel解析） | ◎ 一次公的統計 |",
        f"| 検索トレンド（12ヶ月） | Google Trends（台湾向け） | pytrends API | ○ 相対値（0〜100スコア） |",
        f"| 台湾人の旅行関連口コミ | PTT Japan_Travel板（台湾最大匿名掲示板） | 自動スクレイピング | ○ リアルタイム生声 |",
        f"| 台湾向け旅行ニュース・ブログ | Google News RSS（台湾版・繁体字） | feedparser自動取得 | ○ 直近30日の最新情報 |",
        f"| 消費単価 | 観光庁 訪日外客消費動向調査 2023年 | 公開統計 | ◎ 一次公的統計 |",
        f"| 台湾インバウンド市場概況 | JNTO年次レポート2024 | 公開統計 | ◎ 一次公的統計 |",
        "",
        f"**調査対象データ数**: PTT投稿・ニュース記事 合計**{len(combined_titles)}件** 分析 | Google Trends 12ヶ月分",
        f"**調査実施日**: {today}",
        f"**作成ツール**: {kw}台湾市場向けリサーチダッシュボード v3.0（JR東日本 台湾事業開発）",
        "",
        "---",
        "",
        "> ⚠️ **免責事項**: 本レポートは公開データのみを使用した事前調査であり、最終的な意思決定にあたっては"
        "現地調査・専門家意見・最新統計等と照合のうえご利用ください。",
        "> Google Trendsは相対的な検索人気度（0〜100）を示すものであり、絶対的な検索数を意味しません。",
    ]
    parts.append("\n".join(appendix_lines))

    return "\n".join(parts)


# ============================================================
# メイン UI
# ============================================================


def main() -> None:

    # ── ページヘッダー ──────────────────────────────────────────
    st.title("🇹🇼 台湾市場向け 事前リサーチダッシュボード")
    st.caption("JR東日本 台湾事業開発 ｜ 日本自治体向け提案 事前調査ツール　v3.0")

    # ── 自治体向けクイックガイド（初回表示） ──
    with st.expander("📋 このツールの使い方（提案担当者向けガイド）", expanded=False):
        st.markdown("""
#### 本ツールで「何が」わかるか

| タブ | データソース | わかること | 提案書への活用 |
|---|---|---|---|
| 📊 基礎データ | JNTO公式統計 | 台湾人訪問者数の年別・月別推移 | 市場規模・回復率の数値根拠 |
| 🔍 検索トレンド | Google Trends（台湾） | 何が検索されているか・季節変動 | プロモーション時期・訴求テーマの特定 |
| 💬 生の声 | Google News・PTT掲示板 | 台湾人旅行者の実際の関心・悩み | ターゲットペルソナ・課題設定の根拠 |

#### 推奨ワークフロー（30分で提案資料のデータ収集完了）
1. **左サイドバーで調査地域を入力**（例: 青森、長野）
2. **タブ1「JNTOデータを自動取得」** → 訪問者数グラフ・月別季節性を確認
3. **タブ2でトレンド取得** → 「トレンド分析レポートを生成」で即時分析
4. **タブ3で生の声収集** → 「ペルソナ分析を生成」でターゲット像を確立
5. 各セクションのグラフ・レポートをスクリーンショット → 提案書に貼付

> 💡 **自治体担当者への説明ポイント**: 「JNTOの公式データ」「Google Trendsの実データ」「台湾掲示板の実投稿」という3つの独立したデータソースを組み合わせることで、台湾市場ニーズを客観的に示しています。
""")

    st.divider()

    # ══════════════════════════════════════════════════════════
    # サイドバー
    # ══════════════════════════════════════════════════════════
    with st.sidebar:
        st.header("⚙️ 調査設定")
        st.divider()

        # ── キーワード入力 ──
        keyword = st.text_input(
            "🔍 調査したい地域・キーワード",
            value="青森",
            placeholder="例: 青森、長野、新潟",
            help="都道府県名または観光地名を日本語で入力してください",
        )
        kw = keyword.strip()

        tc_kw = to_tc(kw) if kw else ""
        if tc_kw and tc_kw != kw:
            st.caption(f"🔄 繁体字変換: `{kw}` → `{tc_kw}`")
        elif tc_kw:
            st.caption(f"🔄 繁体字変換なし: `{tc_kw}` のまま使用")

        st.divider()

        # ── YouTube API キー（任意）──
        st.markdown('<p class="sidebar-section">🎥 YouTube（任意・無料）</p>', unsafe_allow_html=True)
        yt_key = st.text_input(
            "YouTube Data API v3 キー",
            type="password",
            placeholder="AIzaSy...",
            help=(
                "Google Cloud Console で無料取得できます（10,000ユニット/日）。\n"
                "入力するとタブ3でYouTube旅行動画が表示されます。\n\n"
                "取得先: console.cloud.google.com\n"
                "→「YouTube Data API v3」を有効化 → APIキーを作成"
            ),
        )
        if yt_key.strip():
            st.success("✅ YouTube 動画検索 有効")
        else:
            st.info("💡 YouTubeキーを入力すると\n台湾人旅行動画が表示されます")

        st.divider()

        if st.button("🔄 データを再取得（キャッシュクリア）", use_container_width=True):
            st.cache_data.clear()
            st.toast("キャッシュをクリアしました！", icon="✅")

        st.divider()
        st.markdown(
            """
**📌 使い方**
1. キーワードを入力
2. タブ1: JNTOデータ自動取得
3. タブ2: 検索トレンド分析
4. タブ3: 台湾の生の声収集
5. 結果を提案資料に活用
"""
        )
        st.caption("© JR東日本 台湾事業開発チーム")

    # ══════════════════════════════════════════════════════════
    # メインエリア：3 タブ
    # ══════════════════════════════════════════════════════════
    tab1, tab2, tab3 = st.tabs(
        [
            "📊 基礎データ（JNTO自動取得）",
            "🔍 検索トレンド（Google Trends + 分析）",
            "💬 生の声（ニュース・YouTube・PTT）",
        ]
    )

    # ──────────────────────────────────────────────────────────
    # タブ1：基礎データ（JNTO自動取得）
    # ──────────────────────────────────────────────────────────
    with tab1:
        st.header("📊 台湾人訪問者数の推移（JNTO自動取得）")
        st.info(
            "**JNTO（日本政府観光局）** の公開統計 Excel を自動取得・解析します。  \n"
            "「自動取得」ボタンを押すと最新データを取得します（24時間キャッシュ）。"
        )

        col_btn, col_status = st.columns([1, 3])
        with col_btn:
            run_jnto = st.button("📥 JNTOデータを自動取得", use_container_width=True, type="primary")
        with col_status:
            if "jnto_data" in st.session_state:
                st.success("✅ JNTOデータ取得済み（キャッシュ）")

        if run_jnto:
            with st.spinner("JNTOサイトからExcelを取得・解析中…（10〜30秒）"):
                jnto_df, jnto_monthly_df, jnto_msg = fetch_jnto_taiwan_data()
            if jnto_df is not None:
                st.session_state["jnto_data"]         = jnto_df
                st.session_state["jnto_monthly_data"] = jnto_monthly_df
                st.session_state["jnto_msg"]          = jnto_msg
                st.rerun()
            else:
                st.error(
                    f"❌ JNTOデータの取得に失敗しました。  \n"
                    f"詳細: {jnto_msg}  \n\n"
                    "**対処法**: 下部の「CSVアップロード」から手動で入力できます。  \n"
                    f"データ取得先: {JNTO_STATS_URL}"
                )

        # 取得済みデータの表示
        if "jnto_data" in st.session_state:
            jnto_df         = st.session_state["jnto_data"]
            jnto_monthly_df = st.session_state.get("jnto_monthly_data")
            jnto_msg        = st.session_state.get("jnto_msg", "")

            # ── KPI サマリー（自治体向けエグゼクティブビュー）──
            latest    = jnto_df.iloc[-1]
            latest_yr = int(latest["年"])
            latest_val = int(latest["台湾人訪問者数（人）"])

            # 前年比
            prev_row = jnto_df[jnto_df["年"] == latest_yr - 1]
            yoy_delta = None
            if not prev_row.empty:
                prev_val  = int(prev_row.iloc[0]["台湾人訪問者数（人）"])
                yoy_delta = f"{(latest_val - prev_val) / prev_val * 100:+.1f}%"

            # コロナ前比（2019年）
            pre_covid = jnto_df[jnto_df["年"] == 2019]
            recovery_str = "—"
            if not pre_covid.empty:
                base_val = int(pre_covid.iloc[0]["台湾人訪問者数（人）"])
                if base_val > 0:
                    r = latest_val / base_val * 100
                    recovery_str = f"{r:.1f}%"

            # データ年数
            num_years = len(jnto_df)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric(
                f"台湾人訪問者数（{latest_yr}年）",
                f"{latest_val:,.0f} 人",
                delta=yoy_delta,
            )
            m2.metric("コロナ前（2019年）比", recovery_str)
            m3.metric("データ期間", f"{int(jnto_df.iloc[0]['年'])}〜{latest_yr}年")
            m4.metric("収録年数", f"{num_years} 年分")

            if jnto_msg and jnto_msg != "OK":
                st.caption(f"ℹ️ 解析モード: {jnto_msg}　※年ヘッダーが自動推定されています")

            st.divider()

            # ── メインチャート：月別データ優先 ──
            if jnto_monthly_df is not None and not jnto_monthly_df.empty:
                # ── 月別棒グラフ（メイン表示）──
                st.subheader("📅 月別訪問者数の季節性（直近3年平均）")
                st.caption(
                    "各月の平均訪問者数。**いつプロモーション・旅行博参加を集中させるべきか**が一目でわかります。  \n"
                    "台湾人の検索ピーク（旅行計画時期）は実際の訪日の約2〜3ヶ月前に来ます。"
                )

                # 月別平均を計算（直近3年分のみ使用）
                recent_years = sorted(jnto_monthly_df["年"].unique())[-3:]
                recent_monthly = jnto_monthly_df[jnto_monthly_df["年"].isin(recent_years)]
                monthly_avg = (
                    recent_monthly.groupby("月")["訪問者数"]
                    .mean()
                    .reset_index()
                    .rename(columns={"訪問者数": "月別平均訪問者数（人）"})
                    .sort_values("月")
                )

                # 月名ラベル（季節色分けのためラベルに季節を付与）
                month_labels = {
                    1: "1月 冬", 2: "2月 冬", 3: "3月 春", 4: "4月 春",
                    5: "5月 春", 6: "6月 夏", 7: "7月 夏", 8: "8月 夏",
                    9: "9月 秋", 10: "10月 秋", 11: "11月 秋", 12: "12月 冬",
                }
                monthly_avg["月ラベル"] = monthly_avg["月"].map(month_labels)

                peak_mo = int(monthly_avg.loc[monthly_avg["月別平均訪問者数（人）"].idxmax(), "月"])
                low_mo  = int(monthly_avg.loc[monthly_avg["月別平均訪問者数（人）"].idxmin(), "月"])
                peak_promo_mo = (peak_mo - 3) if peak_mo > 3 else (peak_mo + 9)

                col_mc, col_mt = st.columns([4, 1])
                with col_mc:
                    st.bar_chart(
                        monthly_avg.set_index("月ラベル")["月別平均訪問者数（人）"],
                        use_container_width=True,
                        height=320,
                    )
                with col_mt:
                    st.metric("訪問ピーク月", f"{peak_mo}月")
                    st.metric("プロモーション\n推奨時期", f"{peak_promo_mo}月頃")
                    st.metric("閑散月", f"{low_mo}月")
                    st.caption(f"集計: {min(recent_years)}〜{max(recent_years)}年平均")

                st.info(
                    f"**📢 プロモーション推奨**: **{peak_promo_mo}月頃**（ピーク {peak_mo}月 の2〜3ヶ月前）  \n"
                    f"旅行博・SNSキャンペーン・旅行会社商談をこの時期に集中させることが効果的です。  \n"
                    f"閑散期（{low_mo}月）には「{low_mo}月限定」の特別コンテンツで閑散期需要を掘り起こします。"
                )

                st.divider()

                # ── 年別推移（棒グラフ、サブ表示）──
                st.subheader("📊 台湾人訪問者数 年別推移（JNTOデータ）")
                recent_df = jnto_df.tail(10).copy()
                col_chart, col_table = st.columns([3, 1])
                with col_chart:
                    st.bar_chart(
                        recent_df.set_index("年")["台湾人訪問者数（人）"],
                        use_container_width=True,
                        height=280,
                    )
                with col_table:
                    disp = recent_df.copy()
                    disp["台湾人訪問者数（人）"] = disp["台湾人訪問者数（人）"].map(lambda x: f"{x:,.0f}")
                    st.dataframe(disp, use_container_width=True, hide_index=True)

            else:
                # 月別データなし → 年別棒グラフをメインで表示
                st.subheader("📊 台湾人訪問者数 年別推移（JNTOデータ）")
                st.caption("※月別データが取得できなかったため年別表示です。月別季節性はデータ取得後に表示されます。")
                recent_df = jnto_df.tail(10).copy()
                col_chart, col_table = st.columns([3, 1])
                with col_chart:
                    st.bar_chart(
                        recent_df.set_index("年")["台湾人訪問者数（人）"],
                        use_container_width=True,
                        height=320,
                    )
                with col_table:
                    disp = recent_df.copy()
                    disp["台湾人訪問者数（人）"] = disp["台湾人訪問者数（人）"].map(lambda x: f"{x:,.0f}")
                    st.dataframe(disp, use_container_width=True, hide_index=True)

            _insight_box(
                "<strong>💡 自治体提案 活用ポイント</strong><br>"
                "・ JNTOの公式データをグラフで示すことで「台湾市場の回復曲線」を数値根拠として提示<br>"
                "・ コロナ前（2019年）比の回復率が高いほど、JRパス連携提案の訴求力が増す<br>"
                "・ 月別季節性グラフを使い「いつ仕掛けるか」を具体的なプロモーション計画として提案に落とし込める<br>"
                "・ 最新年のデータを「台湾市場ポテンシャルの指標」として提案書の冒頭に引用"
            )
        else:
            # まだ取得していない場合はCSVアップローダー
            st.info("👆 上の「JNTOデータを自動取得」ボタンを押してください。")

        # ── CSV フォールバック（手動入力オプション）──
        with st.expander("📂 CSVで手動入力（自動取得が失敗した場合）"):
            uploaded = st.file_uploader(
                "JNTO 訪日外客数データ（CSV）をアップロード",
                type=["csv"],
                help="列構成の推奨: 年度 | 台湾人訪問者数（人）",
            )
            if uploaded is not None:
                try:
                    df_raw = pd.read_csv(uploaded, encoding="utf-8-sig")
                    st.success(f"✅ CSV読み込み完了（{len(df_raw)} 行）")
                    num_cols = df_raw.select_dtypes(include="number").columns.tolist()
                    if num_cols and len(df_raw.columns) >= 2:
                        idx_col = df_raw.columns[0]
                        st.line_chart(df_raw.set_index(idx_col)[num_cols], use_container_width=True)
                    st.dataframe(df_raw.head(20), use_container_width=True)
                except Exception as exc:
                    try:
                        uploaded.seek(0)
                        df_raw = pd.read_csv(uploaded, encoding="shift-jis")
                        st.success(f"✅ CSV読み込み完了（Shift-JIS）（{len(df_raw)} 行）")
                        st.dataframe(df_raw.head(20), use_container_width=True)
                    except Exception as e2:
                        st.error(f"CSVの読み込みに失敗しました: {e2}")

        # ── インバウンド人気スポット ──────────────────────────────
        st.divider()
        st.subheader("🗺️ 台湾人インバウンドに人気のスポット・グルメ")
        st.caption(
            "台湾のSNS・旅行誌・口コミで言及の多い観光地・グルメ・体験施設です。"
            "自治体への提案時に「台湾人が実際に関心を持っているコンテンツ」の根拠資料として活用できます。"
        )

        if not kw:
            st.info("👈 左のサイドバーからキーワードを入力すると対応エリアのスポットが表示されます。")
        else:
            spots_list = get_inbound_spots(kw)

            if spots_list:
                spots_df = pd.DataFrame(spots_list)
                st.dataframe(
                    spots_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "カテゴリ":           st.column_config.TextColumn("カテゴリ", width="small"),
                        "スポット名":         st.column_config.TextColumn("スポット名", width="medium"),
                        "台湾人への訴求ポイント": st.column_config.TextColumn("台湾人への訴求ポイント", width="large"),
                        "最寄り駅・JR連携":   st.column_config.TextColumn("最寄り駅・JR連携", width="medium"),
                    },
                )
            else:
                st.info(
                    f"「{kw}」の定番スポットデータは未登録です。  \n"
                    "下の「Trendsデータからスポット候補を取得」でトレンドベースの候補が表示されます。"
                )

        # ── Google Trends トレンドスポット候補 ──────────────────
        st.divider()
        st.subheader("📈 Google Trendsによる注目スポット候補（動的）")
        st.caption(
            "タブ2でトレンドデータを取得すると、Google Trendsの関連クエリから"
            "台湾人が今注目しているスポット・テーマを自動抽出します。"
        )

        trends_key = f"trends_related_{kw}"
        if trends_key in st.session_state:
            top_df_cached, rising_df_cached = st.session_state[trends_key]
            trending_spots = extract_trending_spots(top_df_cached, rising_df_cached, kw)
            if trending_spots:
                trend_spots_df = pd.DataFrame(trending_spots)
                st.dataframe(
                    trend_spots_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "トレンド種別": st.column_config.TextColumn("種別", width="small"),
                        "キーワード":   st.column_config.TextColumn("キーワード", width="medium"),
                        "カテゴリ":     st.column_config.TextColumn("カテゴリ", width="medium"),
                        "相対値":       st.column_config.TextColumn("相対値", width="small"),
                    },
                )
                _insight_box(
                    "<strong>💡 トレンドスポット活用法</strong><br>"
                    "・ 急上昇キーワード（🚀）がまだ定番化していない「先手コンテンツ」の候補<br>"
                    "・ 定番人気（🏆）はすでに需要が確立されており、安定した訴求力を持つ<br>"
                    "・ カテゴリを組み合わせて「温泉×グルメ」「景観×体験」のモデルコースを提案"
                )
            else:
                st.info("Trendsキーワードからスポット候補が見つかりませんでした。")
        else:
            st.info(
                "💡 **タブ2「検索トレンド」でキーワードを検索すると**、  \n"
                "ここに最新のトレンドスポット候補が自動表示されます。"
            )

    # ──────────────────────────────────────────────────────────
    # タブ2：検索トレンド（Google Trends + テンプレート分析）
    # ──────────────────────────────────────────────────────────
    with tab2:
        st.header("🔍 台湾での Google 検索トレンド分析")
        st.info(
            "**Google Trends（台湾向け）** データから、台湾人の検索ボリュームと季節変動を把握します。  \n"
            "「データ分析レポートを生成」ボタンで **APIキー不要のテンプレート分析** が即座に実行されます。"
        )

        if not kw:
            st.warning("👈 左のサイドバーからキーワードを入力してください。")
        else:
            with st.spinner(f"Google Trends からデータ取得中（`{kw}`）…"):
                interest_df, related_queries, err = fetch_google_trends(kw)

            if err == "rate_limit":
                st.error(
                    """
### ⏱️ Google Trends のアクセス制限に引っかかりました

**対処法：**
- **5〜10 分待ってから再試行**してください
- サイドバーの「データを再取得」ボタンでキャッシュをクリアしてから再試行
"""
                )
            elif err is not None:
                detail = err[len("other:"):] if err.startswith("other:") else err
                st.warning(
                    f"⚠️ データ取得エラー。しばらく待ってから再試行してください。\n\n```\n{detail}\n```"
                )
            elif interest_df is None or interest_df.empty:
                st.warning(
                    f"⚠️ **「{kw}」** の台湾での検索データが見つかりませんでした。  \n"
                    "都道府県名など、より一般的なキーワードでお試しください。"
                )
            else:
                # 関連クエリをセッションステートに保存（タブ1のTrendsスポットで使用）
                rq_key   = next(iter(related_queries), None) if related_queries else None
                rq_data  = related_queries.get(rq_key, {}) if rq_key else {}
                top_df   = rq_data.get("top")
                rising_df = rq_data.get("rising")
                st.session_state[f"trends_related_{kw}"] = (top_df, rising_df)

                # ── 検索ボリューム推移 ──
                st.subheader(f"📈 「{kw}」の台湾での検索ボリューム推移（過去 12 ヶ月）")
                plot_cols = [c for c in interest_df.columns if c != "isPartial"]
                if plot_cols:
                    plot_df = interest_df[plot_cols].copy()
                    plot_df.columns = ["検索人気度（0-100）"]
                    st.line_chart(plot_df, use_container_width=True)

                _insight_box(
                    "<strong>💡 活用ポイント：旅行前の情報収集時期の特定</strong><br>"
                    "・ 検索ピーク月 ＝ 台湾人が旅行計画を本格化している時期<br>"
                    "・ ピーク月の 2〜3 ヶ月前に旅行博・SNS キャンペーンを集中実施<br>"
                    "・ 「このタイミングでプロモーション予算を投下すべき」と自治体に具体的に提示可能"
                )

                st.divider()

                # ── 関連クエリ ──
                st.subheader("🔗 関連キーワード分析")
                col_top, col_rising = st.columns(2)
                with col_top:
                    st.markdown("#### 🏆 人気のトピック（Top）")
                    st.caption("検索ボリューム上位のキーワード ＝ 定番の関心事")
                    if top_df is not None and not top_df.empty:
                        st.dataframe(
                            top_df.rename(columns={"query": "キーワード", "value": "人気度（相対値）"}),
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.info("データが取得できませんでした")

                with col_rising:
                    st.markdown("#### 🚀 急上昇（Rising）")
                    st.caption("最近急増しているキーワード ＝ 新たなトレンド")
                    if rising_df is not None and not rising_df.empty:
                        st.dataframe(
                            rising_df.rename(columns={"query": "キーワード", "value": "上昇率（%）"}),
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.info("データが取得できませんでした")

                # ── テンプレート分析（ボタン押下時のみ実行） ──
                st.divider()
                st.subheader("📊 データ分析レポート（APIキー不要・無制限）")
                st.caption("ボタンを押すとGoogle Trendsデータから即座に分析レポートを生成します。")

                trends_ss_key = f"trends_analysis_{kw}"
                if st.button("📊 トレンド分析レポートを生成", key="run_trends_analysis", type="primary"):
                    result = generate_trends_analysis(kw, tc_kw, interest_df, top_df, rising_df)
                    st.session_state[trends_ss_key] = result
                    st.rerun()

                if trends_ss_key in st.session_state:
                    _analysis_box(st.session_state[trends_ss_key])

    # ──────────────────────────────────────────────────────────
    # タブ3：生の声（Google News RSS + YouTube + PTT）
    # ──────────────────────────────────────────────────────────
    with tab3:
        st.header("💬 台湾人旅行者の生の声")
        st.info(
            "**Google News RSS（台湾）** から旅行ブログ・記事を自動収集します。  \n"
            "YouTube APIキーがあれば **台湾人旅行動画** も取得できます。  \n"
            "台湾掲示板 **PTT Japan_Travel 板** の投稿も合わせて分析します。"
        )

        if not kw:
            st.warning("👈 左のサイドバーからキーワードを入力してください。")
        else:
            # ── セクション A: Google News RSS ──────────────────
            st.subheader("🗞️ Google News（台湾）から見る旅行者の関心記事")
            st.caption(
                f"「{kw} 日本旅遊」で台湾向け Google News を自動検索。"
                "台湾人が実際に読んでいる旅行ブログ・記事のタイトルから関心事を把握できます。"
            )

            with st.spinner(f"Google News 台湾版を取得中…"):
                news_df, news_err = fetch_google_news_taiwan(kw)

            if news_err:
                _show_news_error(news_err)
            elif not news_df.empty:
                st.success(f"✅ {len(news_df)} 件の記事を取得")
                for _, row in news_df.iterrows():
                    st.markdown(
                        f'<div class="news-card">'
                        f'<a href="{row["URL"]}" target="_blank">📰 {row["タイトル"]}</a><br>'
                        f'<span class="source">📍 {row["ソース"]} ｜ 🕐 {row["日付"]}</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

                # ニュースからキーワード抽出・可視化
                news_titles = news_df["タイトル"].tolist()
                news_kws = analyze_top_keywords(news_titles, exclude_word=kw, top_n=8)
                if news_kws:
                    st.divider()
                    st.markdown("**🔤 ニュース記事の頻出キーワード（台湾語）**")
                    nkw_df = pd.DataFrame(news_kws, columns=["キーワード", "出現回数"]).set_index("キーワード")
                    st.bar_chart(nkw_df, use_container_width=True)
            else:
                st.info("Google News から記事を取得できませんでした。")

            st.divider()

            # ── セクション B: YouTube 動画 ──────────────────────
            st.subheader("🎥 YouTube 台湾人旅行動画")
            if yt_key.strip():
                st.caption(
                    f"「{kw} 日本旅遊」で台湾向け YouTube を検索。"
                    "台湾人 YouTuber の旅行動画タイトルから関心テーマを把握できます。"
                )
                with st.spinner("YouTube 動画を検索中…"):
                    yt_df, yt_err = fetch_youtube_taiwan(kw, yt_key.strip())

                if yt_err:
                    _show_youtube_error(yt_err)
                elif not yt_df.empty:
                    st.success(f"✅ {len(yt_df)} 件の動画を取得")
                    st.dataframe(
                        yt_df[["タイトル", "チャンネル", "投稿日", "概要"]],
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "タイトル": st.column_config.LinkColumn(
                                "動画タイトル",
                                display_text="▶ %(タイトル)s",
                            ),
                        },
                    )
                    # YouTube タイトルをリンク付きで表示
                    for _, row in yt_df.iterrows():
                        st.markdown(
                            f"🎬 [{row['タイトル']}]({row['URL']})  \n"
                            f"   📺 {row['チャンネル']} ｜ 📅 {row['投稿日']}"
                        )
                else:
                    st.info("YouTube 動画が見つかりませんでした。")
            else:
                st.info(
                    "💡 **YouTube 動画を表示するには：**  \n"
                    "左サイドバーに YouTube Data API v3 キーを入力してください。  \n"
                    "Google Cloud Console で無料取得（1日10,000ユニット）できます。"
                )

            st.divider()

            # ── セクション C: PTT Japan_Travel 板 ───────────────
            st.subheader("📌 PTT Japan_Travel 板（台湾最大掲示板）")
            st.caption("台湾最大の匿名掲示板 PTT の旅行板から生の声を収集します。")

            col_ptt, col_retry = st.columns([5, 1])
            with col_ptt:
                with st.spinner(f"PTT 取得中（`{kw}`）…"):
                    ptt_df, ptt_err = scrape_ptt(kw)

            # 繁体字でリトライ
            if ptt_err == "no_posts" and tc_kw and tc_kw != kw:
                with st.spinner(f"繁体字（`{tc_kw}`）でリトライ中…"):
                    ptt_df, ptt_err = scrape_ptt(tc_kw)

            if ptt_err:
                _show_ptt_error(ptt_err)
            else:
                st.success(f"✅ {len(ptt_df)} 件取得")
                st.dataframe(
                    ptt_df[["タイトル", "日付", "推薦"]],
                    use_container_width=True,
                    hide_index=True,
                )

                # 投稿リンク
                with st.expander(f"🔗 PTT 全投稿リンク（計 {len(ptt_df)} 件）"):
                    for _, row in ptt_df.iterrows():
                        st.markdown(f"🟠 [{row['タイトル']}]({row['URL']})")

            st.divider()

            # ── セクション D: キーワード頻度 + WordCloud + ペルソナ分析 ──
            # PTT データがある場合のみ
            combined_titles: list[str] = []
            top_kws: list[tuple] = []  # Section E でも参照するため先に初期化
            if ptt_df is not None and not ptt_df.empty:
                combined_titles += ptt_df["タイトル"].tolist()
            if news_df is not None and not news_df.empty:
                combined_titles += news_df["タイトル"].tolist()

            if combined_titles:
                st.subheader("🧠 キーワード頻度分析（PTT + ニュース統合）")
                top_kws = analyze_top_keywords(combined_titles, exclude_word=kw)

                if top_kws:
                    kw_col, bar_col = st.columns([1, 2])
                    with kw_col:
                        st.markdown("#### 📊 頻出キーワード Top 5")
                        kw_df = pd.DataFrame(top_kws, columns=["キーワード", "出現回数"])
                        st.dataframe(kw_df, use_container_width=True, hide_index=True)
                    with bar_col:
                        bar_df = (
                            pd.DataFrame(top_kws, columns=["キーワード", "出現回数"])
                            .set_index("キーワード")
                        )
                        st.bar_chart(bar_df, use_container_width=True)

                    _insight_box(
                        "<strong>【分析結果】台湾のユーザーはこれらのキーワードに強い関心を持っています。</strong><br>"
                        "・ 上記キーワードを軸にしたモデルコース・JRパス連携提案が有効<br>"
                        "・ PTTニュースの両方で共通して出現するキーワードが特に重要なニーズを示す"
                    )

                st.divider()

                # WordCloud
                st.subheader("☁️ WordCloud（台湾語テキスト可視化）")
                with st.spinner("CJK フォントを確認中…"):
                    font_path = get_cjk_font_path()

                if font_path is None:
                    st.warning("⚠️ CJK 対応フォントが見つかりません。WordCloud を表示できません。")
                else:
                    with st.spinner("WordCloud を生成中…"):
                        fig = generate_wordcloud_fig(combined_titles, font_path, exclude_word=kw)

                    if fig is not None:
                        st.pyplot(fig, use_container_width=True)
                        plt.close(fig)
                    else:
                        st.info("WordCloud の生成に失敗しました（データ不足か wordcloud ライブラリ未インストール）。")

                st.divider()

                # ── ペルソナ分析（テンプレートベース・即時実行） ──
                st.subheader("🎯 ターゲットペルソナ分析（APIキー不要・即時生成）")
                persona_ss_key = f"persona_{kw}"

                if st.button("🎯 ペルソナ分析レポートを生成", key="run_persona", type="primary"):
                    ptt_titles = ptt_df["タイトル"].tolist() if ptt_df is not None and not ptt_df.empty else []
                    result = generate_persona_analysis(kw, tc_kw, ptt_titles, top_kws if top_kws else [])
                    st.session_state[persona_ss_key] = result
                    st.rerun()

                if persona_ss_key in st.session_state:
                    _analysis_box(st.session_state[persona_ss_key])
                else:
                    st.info("上のボタンを押すとペルソナ分析が即座に生成されます（APIキー不要）。")

            st.divider()

            # ── セクション E: 提案書コピペ用サマリー ──────────────
            st.subheader("📋 提案書コピペ用サマリー（プロフェッショナル版）")
            st.caption(
                "収集した全データを統合し、自治体向け提案書で実際に使える市場調査レポートを自動生成します。  \n"
                "**エグゼクティブサマリー・市場分析・課題整理・アクションプラン・経済効果試算**を含む構成で出力されます。"
            )

            # ── データ品質チェックパネル（専門性確認用） ──────────────
            with st.expander("🔍 データ品質チェック（生成前に確認・推奨）", expanded=False):
                # データソースの完備状況を確認
                _chk_jnto    = st.session_state.get("jnto_data") is not None
                _chk_jnto_mo = st.session_state.get("jnto_monthly_data") is not None
                _chk_trends  = f"trends_related_{kw}" in st.session_state
                _chk_ptt     = ptt_df is not None and not ptt_df.empty
                _chk_news    = news_df is not None and not news_df.empty

                # 品質スコア（100点満点: 各データの重要度で重み付け）
                _score_items = [
                    (_chk_jnto,    25, "JNTO訪問者数",    "タブ1「JNTO自動取得」ボタン"),
                    (_chk_jnto_mo, 10, "JNTO月別データ",  "タブ1「JNTO自動取得」ボタン"),
                    (_chk_trends,  35, "Google Trends",   "タブ2「Google Trendsを取得」ボタン"),
                    (_chk_ptt,     15, "PTT投稿",         "タブ3「PTT投稿を取得」ボタン"),
                    (_chk_news,    15, "ニュース記事",     "タブ3（自動取得）"),
                ]
                _score = sum(pts for ok, pts, _, __ in _score_items if ok)

                qc_col1, qc_col2 = st.columns([1, 2])

                with qc_col1:
                    st.metric("レポート品質スコア", f"{_score} / 100点")
                    st.progress(_score / 100)
                    if _score >= 80:
                        st.success("🎉 高品質なレポートを生成できます")
                    elif _score >= 50:
                        st.warning("⚠️ 一部のデータが不足しています")
                    else:
                        st.error("🔴 データ不足。各タブで取得してください")

                with qc_col2:
                    st.markdown("**📊 データソース別ステータス**")
                    for ok, pts, name, hint in _score_items:
                        if ok:
                            st.write(f"✅ **{name}** +{pts}点")
                        else:
                            st.write(f"⚠️ **{name}** 未取得 → {hint}")

                # 除外キーワードの表示（透明性の確保）
                _tr_data_qc   = st.session_state.get(f"trends_related_{kw}")
                _rising_df_qc = _tr_data_qc[1] if _tr_data_qc else None
                if _rising_df_qc is not None and not _rising_df_qc.empty:
                    _raw_qc      = [str(q) for q in _rising_df_qc["query"].tolist()[:20]]
                    _filtered_qc = _filter_promo_kws(_raw_qc)
                    _excluded_qc = [k for k in _raw_qc if k not in _filtered_qc]

                    st.divider()
                    st.markdown("**🚫 自動フィルタリング（プロモーション品質管理）**")
                    if _excluded_qc:
                        st.warning(
                            f"⚠️ 不適切キーワード **{len(_excluded_qc)}件** をレポートから自動除外しました  \n"
                            f"（観光PRとして不謹慎・ブランドイメージ毀損のリスクがある語句）"
                        )
                        cols_ex = st.columns(min(len(_excluded_qc), 3))
                        for i, ex_kw in enumerate(_excluded_qc):
                            with cols_ex[i % 3]:
                                st.error(f"~~{ex_kw}~~")
                        st.caption("上記は Google Trends の急上昇ワードとして検出されましたが、災害・事故・政治的に敏感な語句のためレポートへの掲載を除外しています。")
                    else:
                        st.success("✅ 急上昇KWに不適切なワードは含まれていません（全件使用可能）")

                    if _filtered_qc:
                        st.markdown(f"**使用する急上昇KW（観光関連）**: {' / '.join(_filtered_qc[:6])}")

                # 品質向上のヒント
                _tips = []
                for ok, pts, name, hint in _score_items:
                    if not ok:
                        _tips.append((name, hint, pts))

                if _tips:
                    st.divider()
                    st.markdown("**💡 品質向上のヒント（不足データを取得するとレポートが充実します）**")
                    for name, hint, pts in _tips:
                        st.markdown(f"- **{name}** が未取得（+{pts}点） → {hint}")

            # ── 生成ボタン ───────────────────────────────────────
            col_btn_e, col_note_e = st.columns([2, 3])
            with col_btn_e:
                gen_btn = st.button("📋 提案書レポートを生成", key="gen_summary", type="primary", use_container_width=True)
            with col_note_e:
                missing = []
                if "jnto_data" not in st.session_state:
                    missing.append("JNTO数値（タブ1）")
                if f"trends_related_{kw}" not in st.session_state:
                    missing.append("Trendsデータ（タブ2）")
                if missing:
                    st.info(f"💡 未取得データがあると一部が「データ収集中」と表示されます: {' / '.join(missing)}")
                else:
                    st.success("✅ 全データ取得済み。最も詳細なレポートが生成されます")

            if gen_btn:
                # session_state から各データを取得
                _jnto_df    = st.session_state.get("jnto_data")
                _jnto_mon   = st.session_state.get("jnto_monthly_data")
                _tr_data    = st.session_state.get(f"trends_related_{kw}")
                _top_df_t   = _tr_data[0] if _tr_data else None
                _rising_df_t = _tr_data[1] if _tr_data else None

                summary_text = _generate_proposal_text(
                    kw=kw,
                    tc_kw=tc_kw,
                    jnto_df=_jnto_df,
                    jnto_monthly_df=_jnto_mon,
                    top_df_t=_top_df_t,
                    rising_df_t=_rising_df_t,
                    combined_titles=combined_titles,
                    top_kws=top_kws,
                    ptt_count=len(ptt_df) if ptt_df is not None and not ptt_df.empty else 0,
                    news_count=len(news_df) if news_df is not None and not news_df.empty else 0,
                )
                st.session_state["proposal_summary"] = summary_text
                st.rerun()

            if "proposal_summary" in st.session_state:
                # プレビュー（Markdownレンダリング）
                with st.expander("📄 レポートプレビュー（クリックで展開）", expanded=True):
                    st.markdown(st.session_state["proposal_summary"])

                # ダウンロードボタン
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        "💾 Markdownとしてダウンロード（PowerPoint/Word貼付用）",
                        data=st.session_state["proposal_summary"],
                        file_name=f"{kw}_台湾インバウンド市場調査レポート.md",
                        mime="text/markdown",
                        use_container_width=True,
                    )
                with col_dl2:
                    # テキスト形式（メール・チャット貼付用）
                    plain_text = st.session_state["proposal_summary"].replace("**", "").replace("*", "").replace("#", "").replace("|", "│").replace("---", "───")
                    st.download_button(
                        "📝 プレーンテキストとしてダウンロード（メール・チャット用）",
                        data=plain_text,
                        file_name=f"{kw}_台湾インバウンド市場調査レポート.txt",
                        mime="text/plain",
                        use_container_width=True,
                    )


# ============================================================
# エントリーポイント
# ============================================================
if __name__ == "__main__":
    main()
