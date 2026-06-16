# -*- coding: utf-8 -*-
"""
生わらび餅 営業リスト自動作成ツール（Webアプリ版 / Streamlit）

URLを貼るだけで、屋号・場所・業態・問い合わせ先・営業メール（件名＋本文）を
自動で作り、画面表示＋Excel/CSVでダウンロードできる。

必要な secrets（Streamlit Cloud の Settings → Secrets に設定）:
    ANTHROPIC_API_KEY = "sk-ant-..."
    APP_PASSWORD      = "合言葉"
"""

import io
import re
import time

import requests
import pandas as pd
import streamlit as st
from anthropic import Anthropic

DEFAULT_MODEL = "claude-haiku-4-5-20251001"
PAGE_TEXT_LIMIT = 6000

# ===== 商品情報の初期値（京都利休の生わらび餅）=====
DEFAULT_PRODUCT = """【ブランド】京都利休の生わらび餅（京都の和菓子職人の技術 × 千利休の"おもてなし"の精神）

【商品の特徴・売り文句】
1. とろける食感：口に入れた瞬間にほどける滑らかさ。"飲むわらび餅"としてメディア・SNSでも話題
2. 最高級の品質：国産Aランク本わらび粉＋無添加素材。京都ブランドの最高級イメージ
3. 手間いらず：解凍後そのまま盛り付けるだけ
4. アレンジ自在：ラテ・パフェ・デザートプレートなどに展開しやすい
5. 粗利率40%超：製造を効率化し、導入しやすい価格帯を実現
6. 長い消費期限：冷凍8か月／解凍後5日（従来の生わらび餅は1〜2日。廃棄リスクを大幅に低減）
7. 小ロット対応：個包装「60g×10ピース」から発注可能

【メニュー活用例・上代】
定番：和三盆＋京きな粉 600〜1400円／宇治抹茶＋抹茶きな粉 700〜1500円／沖縄黒糖＋京きな粉 700〜1500円／和三盆＋あずき 700〜1500円
ラテ：いちご・抹茶・黒蜜・ミルクティー 各500〜800円
小売・お土産：和三盆＋きな粉 450〜900円／宇治抹茶＋抹茶きな粉 450〜900円／黒糖＋抹茶きな粉 450〜900円／利休三彩餅 1500〜1800円

【導入メリット】
・幅広い客層に訴求（特にファミリー層に人気）
・SNS・メディアでの話題性（ふわとろ食感／飲むわらび餅）
・京都ブランドの特別感・非日常の演出"""

DEFAULT_SIG = {
    "会社名": "京都利休の生わらび餅 南関東（株式会社BBスクウェア）",
    "担当者名": "",
    "住所": "〒214-0038 神奈川県川崎市多摩区生田6-26-7",
    "電話": "",
    "メール": "",
    "HP": "",
}

COLUMNS = ["URL", "屋号", "場所（市）", "業態", "問い合わせ有無", "問い合わせ先",
           "営業メール件名", "営業メール本文", "ステータス", "備考"]

# ポータル運営者（楽天・Yahoo!・食べログ等）自身の問い合わせを除外するための判定
PORTAL_HOST = re.compile(
    r"(rakuten\.co\.jp|rms\.rakuten|faq\.rakuten|help\.rakuten|ichiba|yahoo\.co\.jp|"
    r"shopping\.yahoo|store\.shopping\.yahoo|tabelog\.com|gnavi\.co\.jp|hotpepper\.jp|"
    r"recruit\.co\.jp|amazon\.co\.jp|askul\.co\.jp|minne\.com|creema\.jp|ozmall\.co\.jp)",
    re.I,
)

SNS_PATTERNS = [
    ("Instagram", re.compile(r"instagram\.com/([^/?#]+)", re.I)),
    ("Facebook", re.compile(r"(?:facebook|fb)\.com/([^/?#]+)", re.I)),
    ("X(Twitter)", re.compile(r"(?:twitter|x)\.com/([^/?#]+)", re.I)),
    ("TikTok", re.compile(r"tiktok\.com/@?([^/?#]+)", re.I)),
    ("Threads", re.compile(r"threads\.(?:net|com)/@?([^/?#]+)", re.I)),
    ("LINE", re.compile(r"(?:lin\.ee|page\.line\.me)/([^/?#]+)", re.I)),
]
SNS_SKIP = re.compile(r"^(p|reel|reels|explore|accounts|stories|tv|s|hashtag|pages|profile\.php)$", re.I)


# ===== URL抽出 =====
def extract_urls(text):
    """貼り付けた文章から http(s) URL を全部拾う。区切りは何でもOK・重複は除く。"""
    if not text:
        return []
    raw = re.findall(r"https?://[^\s,、　\"'<>）)]+", text)
    seen, urls = set(), []
    for u in raw:
        u = u.rstrip(".,;　)")  # 末尾の句読点・括弧を除去
        if u and u not in seen:
            seen.add(u)
            urls.append(u)
    return urls


# ===== サイト取得・解析 =====
def detect_sns(url):
    for name, pat in SNS_PATTERNS:
        m = pat.search(url)
        if m:
            handle = (m.group(1) or "").lstrip("@")
            if SNS_SKIP.match(handle):
                handle = ""
            return {"platform": name, "handle": handle}
    return None


def fetch_page(url):
    resp = requests.get(
        url, timeout=20, allow_redirects=True,
        headers={"User-Agent": "Mozilla/5.0 (compatible; WarabiSalesBot/1.0)"},
    )
    resp.raise_for_status()
    if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
        resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def html_to_text(html):
    t = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
    t = re.sub(r"<style[\s\S]*?</style>", " ", t, flags=re.I)
    t = re.sub(r"<noscript[\s\S]*?</noscript>", " ", t, flags=re.I)

    def first(pattern):
        m = re.search(pattern, t, re.I)
        return (m.group(1) if m else "").strip()

    title = first(r"<title[^>]*>([\s\S]*?)</title>")
    desc = first(r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']*)')
    og_title = first(r'<meta[^>]+property=["\']og:title["\'][^>]+content=["\']([^"\']*)')

    body = re.sub(r"<[^>]+>", " ", t)
    body = (body.replace("&nbsp;", " ").replace("&amp;", "&")
            .replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"'))
    body = re.sub(r"\s+", " ", body).strip()

    out = f"タイトル: {title or og_title}\n説明: {desc}\n本文: {body}".strip()
    return out[:PAGE_TEXT_LIMIT]


def resolve_url(href, base_url):
    href = (href or "").strip()
    if not href or href.startswith("#") or href.lower().startswith("javascript:"):
        return ""
    if re.match(r"^https?://", href, re.I):
        return href
    m = re.match(r"^(https?://[^/]+)", base_url, re.I)
    origin = m.group(1) if m else ""
    if not origin:
        return href
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("/"):
        return origin + href
    return re.sub(r"[^/]*$", "", base_url) + href


def extract_contacts(html, base_url):
    m = re.match(r"^https?://([^/]+)", base_url, re.I)
    host = m.group(1) if m else ""
    on_portal = bool(PORTAL_HOST.search(host))

    emails = []
    for mm in re.finditer(r"mailto:([^\"'?\s>]+)", html, re.I):
        addr = mm.group(1)
        if re.match(r"^[^@]+@(example|sample)\.", addr, re.I):
            continue
        if addr not in emails:
            emails.append(addr)

    forms = []
    for mm in re.finditer(r"href\s*=\s*[\"']([^\"']+)[\"']", html, re.I):
        href = mm.group(1)
        if not re.search(r"contact|inquiry|toiawase|otoiawase|mail|form|お問い合わせ|問い合わせ|問合せ", href, re.I):
            continue
        absu = resolve_url(href, base_url)
        if not absu:
            continue
        lm = re.match(r"^https?://([^/]+)", absu, re.I)
        link_host = lm.group(1) if lm else ""
        if on_portal and PORTAL_HOST.search(link_host):
            continue
        if absu not in forms:
            forms.append(absu)

    return {"emails": emails, "forms": forms[:5], "on_portal": on_portal, "host": host}


# ===== プロンプト構築・AI呼び出し =====
def build_messages(page_text, contacts, sns, product_text):
    sns_hint = ""
    if sns:
        handle = sns.get("handle")
        sns_hint = (
            f"※このURLは {sns['platform']} のプロフィールページです"
            + (f"（アカウント名: @{handle}）" if handle else "")
            + "。\nSNSはページ本文を取得できないことが多いため、web_search ツールで「"
            + (f"@{handle}」「{handle} 店舗" if handle else "このアカウント")
            + "」などを検索し、店名・地域・業態・公式サイト・問い合わせ先を特定してください。\n"
        )

    contact_hint = (
        f"対象サイトのドメイン: {contacts.get('host') or '不明'}\n"
        + sns_hint
        + ("※このURLは楽天・Yahoo!・食べログ等のポータル/モール上の出店ページです。"
           "ポータル運営者宛ての問い合わせは無視し、ショップ自身の問い合わせ先のみを採用してください。\n"
           if contacts.get("on_portal") else "")
        + "検出したメール: " + (", ".join(contacts["emails"]) if contacts["emails"] else "（なし）") + "\n"
        + "検出した問い合わせページ: " + (", ".join(contacts["forms"]) if contacts["forms"] else "（なし）")
    )

    system_text = (
        "あなたは生わらび餅の卸営業を支援するアシスタントです。"
        "与えられたWebサイトの内容から会社情報を読み取り、相手に合わせた営業メールを書きます。"
        "誇張や事実と異なる表現は絶対に書きません。返答は指定された見出し形式のみで出力します。"
    )

    user_text = (
        f"【自社（生わらび餅）の商品・差出人情報】\n{product_text}\n\n"
        f"【対象サイトから検出した連絡先候補】\n{contact_hint}\n\n"
        f"【対象サイトのテキスト】\n{page_text}\n\n"
        "上記の会社・店舗について、次の見出し形式だけで出力してください（JSONやコードブロックは使わない）。\n"
        "各見出しは必ず行頭に置き、本文は指定の囲みタグの中に書くこと。\n\n"
        "屋号: <店舗・会社の名前>\n"
        "場所: <市区町村まで。例: 東京都渋谷区 / 大阪市。不明なら空>\n"
        "業態: <何屋か。例: たい焼き店, ネイルサロン, カフェ>\n"
        "問い合わせ有無: <あり または なし>\n"
        "問い合わせ先: <そのショップ自身のメール または 問い合わせフォームURL。無ければ空>\n"
        "件名: <営業メールの件名>\n"
        "本文:\n"
        "<<<本文ここから>>>\n"
        "（ここに宛名から署名までの営業メール本文。改行はそのまま書いてよい）\n"
        "<<<本文ここまで>>>\n\n"
        "問い合わせ先の判定ルール：\n"
        "- 採用するのは「そのショップ・店舗自身」の問い合わせ先だけ\n"
        "- 楽天・Yahoo!ショッピング・食べログ・ぐるなび・ホットペッパー等、ポータル/モール運営者宛ての問い合わせは採用しない\n"
        "- 与えられたページがポータル/モールの出店ページだったり、ショップ自身の問い合わせ先が見つからない場合は、"
        "web_search ツールで「屋号＋場所」や「屋号＋公式サイト」などを検索し、そのショップの公式サイト・問い合わせページ・メールアドレスを調べ直すこと\n"
        "- 検索で見つけたサイトが、同じ屋号・同じ地域の本人のものか必ず確認してから採用する（別店舗・同名他店を混同しない）\n"
        "- 本当に見つからない場合のみ「問い合わせ有無=なし」とする\n\n"
        "営業メール本文の方針：\n"
        "- 1行目に宛名を入れる。屋号がわかれば「（屋号）　ご担当者様」、わからなければ「ご担当者様」（氏名は無理に書かない）\n"
        "- 宛名の次に、初めての連絡を想定した挨拶を置く。「お世話になります。」「突然のご連絡失礼いたします。」といった一言を入れる\n"
        "- 続けて、自分が何者かを名乗る（差出人の会社名・ブランド名を使い、生わらび餅を卸している旨を簡潔に自己紹介する）\n"
        "- そのうえで、相手の業態・特徴に具体的に触れてから、生わらび餅の導入メリットを訴求する\n"
        "- 飲食店ならメニュー追加・粗利率・盛り付けの手軽さ・長い消費期限を、サロンなら来店時のおもてなしや手土産・物販を切り口にする\n"
        "- 上代やメニュー例は商品情報の範囲だけを使い、数字を創作しない\n"
        "- 落ち着いた丁寧なトーン。宛名・署名を除いた本文は300〜450字程度\n"
        "- 問い合わせ先が無い場合でもメール本文は作成する\n\n"
        "署名のルール（ビジネスメールとして整える）：\n"
        "- 本文の最後に、区切り線「----------------------------------------」を1行入れてから署名を置く\n"
        "- 署名は次の順で、1項目1行ずつ縦に並べる：会社名／担当者名／住所／TEL／Mail／HP\n"
        "- 差出人情報のうち未記入（空欄）の項目はその行ごと省略する\n"
        "- TEL・Mail・HP のように項目名を付けて見やすくする"
    )
    return system_text, user_text


def parse_result(text):
    if not text or not text.strip():
        raise ValueError("AIの返答が空でした")

    body = ""
    bm = re.search(r"<<<本文ここから>>>\s*([\s\S]*?)\s*<<<本文ここまで>>>", text)
    if bm:
        body = bm.group(1).strip()
    else:
        fb = re.search(r"(?:^|\n)\s*本文\s*[:：]?\s*\n?([\s\S]*)$", text)
        if fb:
            body = re.sub(r"<<<[^>]*>>>", "", fb.group(1)).strip()

    idx = text.find("<<<本文ここから>>>")
    if idx >= 0:
        head = text[:idx]
    elif text.find("本文") >= 0:
        head = text[:text.find("本文")]
    else:
        head = text

    def get(label):
        # コロンの後は改行を含まない空白のみ許可（値が空のとき次行を拾わないため）
        m = re.search(r"(?:^|\n)[ \t　]*" + re.escape(label) + r"[ \t　]*[:：][ \t　]*(.*)", head)
        return m.group(1).strip() if m else ""

    result = {
        "屋号": get("屋号"),
        "場所": get("場所"),
        "業態": get("業態"),
        "問い合わせ有無": get("問い合わせ有無"),
        "問い合わせ先": get("問い合わせ先"),
        "メール件名": get("件名"),
        "メール本文": body,
    }
    if not result["メール本文"] and not result["屋号"]:
        raise ValueError("AIの返答を解析できませんでした（先頭: " + text[:120] + "）")
    return result


def call_claude(client, model, max_search, page_text, contacts, sns, product_text):
    system_text, user_text = build_messages(page_text, contacts, sns, product_text)
    tools = [{"type": "web_search_20250305", "name": "web_search", "max_uses": max_search}] if max_search > 0 else []
    msg = client.messages.create(
        model=model,
        max_tokens=2000,
        system=system_text,
        tools=tools,
        messages=[{"role": "user", "content": user_text}],
    )
    text = "".join(getattr(b, "text", "") for b in msg.content if getattr(b, "type", "") == "text")
    return parse_result(text)


def process_url(client, model, max_search, url, product_text):
    url = url.strip()
    if not re.match(r"^https?://", url, re.I):
        raise ValueError("URLが http:// または https:// で始まっていません")
    sns = detect_sns(url)
    page_text = ""
    contacts = {"emails": [], "forms": [], "on_portal": False, "host": ""}
    try:
        html = fetch_page(url)
        page_text = html_to_text(html)
        contacts = extract_contacts(html, url)
    except Exception as e:
        if not sns:
            raise
        page_text = f"（{sns['platform']}のページ本文は取得できませんでした）"
    return call_claude(client, model, max_search, page_text, contacts, sns, product_text)


def build_product_text(product_copy, sig):
    lines = [product_copy.strip(), "", "【差出人情報】"]
    labels = [("会社名", "会社名"), ("担当者名", "担当者名"), ("住所", "住所"),
              ("電話", "電話"), ("メール", "メール"), ("HP", "HP")]
    for key, label in labels:
        v = (sig.get(key) or "").strip()
        if v:
            lines.append(f"{label}：{v}")
    return "\n".join(lines)


# ===== Streamlit 画面 =====
def read_secret_key():
    """Secretsから APIキーを読む。名前のゆらぎ・前後空白に強くする。"""
    try:
        for k in ("ANTHROPIC_API_KEY", "anthropic_api_key", "ANTHROPIC_APIKEY", "API_KEY"):
            v = st.secrets.get(k, "")
            if v:
                return str(v).strip()
    except Exception:
        pass
    return ""


def check_password():
    # APP_PASSWORD が未設定（空）なら合言葉なしで誰でも使える
    expected = st.secrets.get("APP_PASSWORD", "")
    if not expected:
        return True
    if st.session_state.get("auth_ok"):
        return True
    st.title("🍡 生わらび餅 営業リスト作成ツール")
    pw = st.text_input("合言葉を入力してください", type="password")
    if pw:
        if pw == expected:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("合言葉が違います。")
    return False


def main():
    st.set_page_config(page_title="生わらび餅 営業リスト作成ツール", page_icon="🍡", layout="wide")

    if not check_password():
        st.stop()

    secret_key = read_secret_key()  # Secretsから読めればそれを使う

    st.title("🍡 生わらび餅 営業リスト作成ツール")
    st.caption("お客様のサイトURLを貼るだけで、屋号・場所・業態・問い合わせ先・営業メールを自動作成します。")

    # サイドバー：APIキー・商品情報・署名・詳細設定
    with st.sidebar:
        st.header("⚙ 設定")
        if secret_key:
            api_key = secret_key
            st.success("APIキーは設定済みです（Secrets）")
        else:
            api_key = st.text_input("Anthropic APIキー", type="password",
                                    help="sk-ant- で始まるキー。Secretsに設定済みなら不要").strip()
        with st.expander("商品情報・売り文句", expanded=False):
            product_copy = st.text_area("内容", DEFAULT_PRODUCT, height=300)
        with st.expander("差出人（署名）", expanded=False):
            sig = {
                "会社名": st.text_input("会社名", DEFAULT_SIG["会社名"]),
                "担当者名": st.text_input("担当者名", DEFAULT_SIG["担当者名"]),
                "住所": st.text_input("住所", DEFAULT_SIG["住所"]),
                "電話": st.text_input("電話", DEFAULT_SIG["電話"]),
                "メール": st.text_input("メール", DEFAULT_SIG["メール"]),
                "HP": st.text_input("HP・問い合わせURL", DEFAULT_SIG["HP"]),
            }
        with st.expander("詳細設定", expanded=False):
            model = st.selectbox("使用モデル", [DEFAULT_MODEL, "claude-sonnet-4-6"], index=0,
                                 help="精度を上げたいときは Sonnet（やや高コスト）")
            max_search = st.slider("1社あたりのWeb検索 上限", 0, 5, 3,
                                   help="0にすると検索しません（コスト最小・SNSは弱くなる）")

    if not api_key:
        st.warning("左の「⚙ 設定」で Anthropic APIキーを入力してください（または管理者がSecretsに設定）。")
        st.stop()
    client = Anthropic(api_key=api_key)

    product_text = build_product_text(product_copy, sig)

    st.subheader("① URLを貼る（何件でもOK）")
    urls_text = st.text_area(
        "お客様のサイトURL（改行・スペース・カンマ区切り、まとめて貼ってOK）",
        height=240,
        placeholder="https://example-restaurant.com\nhttps://www.instagram.com/example_shop/\n…どんどん貼ってください",
    )
    found = extract_urls(urls_text)
    if found:
        st.caption(f"📌 URLを {len(found)} 件 認識しました")

    run = st.button("② 営業リストを作成する", type="primary")

    if run:
        urls = extract_urls(urls_text)
        if not urls:
            st.warning("URLを1件以上入力してください。")
            st.stop()

        rows = []
        prog = st.progress(0.0)
        status = st.empty()
        for i, url in enumerate(urls, start=1):
            status.write(f"処理中… {i}/{len(urls)}　{url}")
            try:
                d = process_url(client, model, max_search, url, product_text)
                rows.append({
                    "URL": url, "屋号": d["屋号"], "場所（市）": d["場所"], "業態": d["業態"],
                    "問い合わせ有無": d["問い合わせ有無"], "問い合わせ先": d["問い合わせ先"],
                    "営業メール件名": d["メール件名"], "営業メール本文": d["メール本文"],
                    "ステータス": "完了", "備考": "",
                })
            except Exception as e:
                rows.append({
                    "URL": url, "屋号": "", "場所（市）": "", "業態": "",
                    "問い合わせ有無": "", "問い合わせ先": "", "営業メール件名": "", "営業メール本文": "",
                    "ステータス": "エラー", "備考": str(e)[:300],
                })
            prog.progress(i / len(urls))
            time.sleep(0.3)
        status.write("完了しました。")

        df = pd.DataFrame(rows, columns=COLUMNS)
        st.session_state["result_df"] = df

    if "result_df" in st.session_state:
        df = st.session_state["result_df"]
        st.subheader("③ 結果")
        ok = (df["ステータス"] == "完了").sum()
        ng = (df["ステータス"] == "エラー").sum()
        st.write(f"成功 {ok} 件 / エラー {ng} 件")
        st.dataframe(df, use_container_width=True, height=360)

        # ダウンロード
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="営業リスト")
        c1, c2 = st.columns(2)
        c1.download_button("📊 Excelをダウンロード", buf.getvalue(), "営業リスト.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c2.download_button("📄 CSVをダウンロード", df.to_csv(index=False).encode("utf-8-sig"),
                           "営業リスト.csv", mime="text/csv")

        # 営業メールをコピペしやすく表示（各ブロック右上のコピーアイコンで1クリックコピー）
        st.subheader("④ 営業メール（コピー用）")
        st.caption("各ブロックの右上にカーソルを合わせると出る📋アイコンでコピーできます。")
        for _, r in df.iterrows():
            if r["ステータス"] != "完了":
                continue
            title = f"{r['屋号'] or r['URL']}　／　{r['業態']}　{r['場所（市）']}".strip()
            with st.expander(title):
                if r["問い合わせ先"]:
                    st.write(f"**問い合わせ先：** {r['問い合わせ先']}")
                st.markdown("**件名**")
                st.code(r["営業メール件名"], language=None)
                st.markdown("**本文**")
                st.code(r["営業メール本文"], language=None)

        st.info("営業メールは自動生成です。送信前に内容を必ず確認してください。")


if __name__ == "__main__":
    main()
