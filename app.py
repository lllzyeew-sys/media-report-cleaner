import io
import re
import csv
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="PLAUD Media Report Cleaner", layout="wide")

APP_TITLE = "PLAUD Media Report Cleaner"

TITLE_FIELDS = ["title", "headline", "article title", "name"]
HIT_FIELDS = ["hit sentence", "hit text", "snippet", "opening text", "summary", "description"]
CONTENT_FIELDS = ["content", "article content", "body", "full text", "text", "article text"]
SOURCE_FIELDS = ["source", "publication", "outlet", "media name", "domain", "site name"]
URL_FIELDS = ["url", "link", "article url", "source url"]
COUNTRY_FIELDS = ["country", "region", "market", "location", "country/region"]

LISTICLE_PATTERNS = [
    r"\btop\b", r"\bbest\b", r"\blist\b", r"\brecommend", r"\bround\s*up\b",
    r"\btools?\b", r"\bgadgets?\b", r"\bproductivity\b", r"\bwork companion\b",
    r"\bai note takers?\b", r"\btranscription tools?\b", r"\bmeeting tools?\b",
    r"\bfor work\b", r"\bfor meetings\b", r"\bfor productivity\b", r"\bcompare", r"\breview",
]

SUBSTANTIVE_PATTERNS = [
    r"plaud\s+note", r"plaud\s+note\s*pin", r"plaud\s+app", r"ai\s+recorder",
    r"voice\s+recorder", r"transcri", r"summari", r"record", r"meeting", r"note[- ]?taking",
    r"launch", r"announce", r"release", r"funding", r"review", r"hands[- ]?on", r"feature",
]

COMPARISON_PATTERNS = [
    r"\bvs\.?\b", r"\bversus\b", r"\bcompared?\s+(with|to)\b", r"\balternative\s+to\b",
    r"\blike\s+plaud\b", r"\bsimilar\s+to\s+plaud\b", r"\bcompetes?\s+with\s+plaud\b",
]

BOILERPLATE_PATTERNS = [
    r"sponsored", r"affiliate", r"advertisement", r"partner\s+content", r"promo\s+code",
    r"recommended\s+articles", r"related\s+articles", r"tagged", r"footer", r"newsletter",
]

EXECUTIVE_NAME_PATTERNS = [
    r"\bnathan\b", r"\bnathan\s+hsu\b", r"\bnathan\s+xu\b",
]

EXECUTIVE_ROLE_PATTERNS = [
    r"\bceo\b", r"chief\s+executive", r"founder", r"co[- ]?founder",
    r"executive", r"president", r"spokesperson",
]

INTERVIEW_PATTERNS = [
    r"interview", r"q\s*&\s*a", r"q&a", r"podcast", r"conversation\s+with",
    r"spoke\s+with", r"talked\s+with", r"sit[- ]?down", r"profile",
    r"said\s+nathan", r"according\s+to\s+nathan", r"nathan\s+(said|says|told|explained|shared)",
    r"独家", r"专访", r"访谈", r"采访", r"对话",
]

AMERICAS = {
    "united states", "usa", "us", "u.s.", "u.s.a.", "america", "canada", "ca", "mexico",
    "brazil", "argentina", "chile", "colombia", "peru", "uruguay", "paraguay", "bolivia",
    "ecuador", "venezuela", "costa rica", "panama", "guatemala", "honduras", "el salvador",
    "nicaragua", "dominican republic", "puerto rico", "latin america", "latam",
}
JP = {"japan", "jp", "日本"}
EU = {
    "uk", "united kingdom", "england", "scotland", "wales", "ireland", "germany", "france", "italy", "spain",
    "netherlands", "sweden", "switzerland", "austria", "belgium", "denmark", "finland", "norway",
    "poland", "portugal", "greece", "czech republic", "czechia", "hungary", "romania", "bulgaria",
    "croatia", "slovakia", "slovenia", "estonia", "latvia", "lithuania", "luxembourg", "europe", "eu",
}
APAC_EX_JP = {
    "australia", "new zealand", "singapore", "malaysia", "thailand", "indonesia", "vietnam", "philippines",
    "india", "south korea", "korea", "taiwan", "hong kong", "hk", "china", "mainland china", "macau",
    "pakistan", "bangladesh", "sri lanka", "nepal", "cambodia", "laos", "myanmar", "apac", "asia pacific",
}


def normalize_col(col: str) -> str:
    return re.sub(r"\s+", " ", str(col).strip().lower())


def find_column(df: pd.DataFrame, candidates: List[str]) -> str | None:
    norm_map = {normalize_col(c): c for c in df.columns}
    for cand in candidates:
        if cand in norm_map:
            return norm_map[cand]
    for norm, original in norm_map.items():
        for cand in candidates:
            if cand in norm:
                return original
    return None


def value(row: pd.Series, col: str | None) -> str:
    if not col or pd.isna(row.get(col, "")):
        return ""
    return str(row.get(col, "")).strip()


def contains_any(text: str, patterns: List[str]) -> bool:
    return any(re.search(p, text, flags=re.IGNORECASE) for p in patterns)


def source_is_medium_or_substack(source_text: str) -> bool:
    s = source_text.lower()
    return (
        "medium.com" in s
        or "substack.com" in s
        or re.search(r"(^|\W)medium($|\W)", s) is not None
        or "substack" in s
    )


def is_listicle(title: str) -> bool:
    t = title.lower()
    return contains_any(t, LISTICLE_PATTERNS) and ("ai" in t or "tool" in t or "gadget" in t or "productivity" in t or "work" in t or "meeting" in t or "note" in t or "transcription" in t)


def is_executive_interview_or_quote(text: str) -> bool:
    """Return True when the row looks like a PLAUD executive interview, quote, profile, or podcast.

    We require more than the generic word "CEO" alone to avoid keeping unrelated CEO articles.
    Strong signals include Nathan by name, PLAUD CEO, or an interview/quote context tied to PLAUD.
    """
    t = text.lower()
    has_plaud = "plaud" in t
    has_nathan = contains_any(t, EXECUTIVE_NAME_PATTERNS)
    has_role = contains_any(t, EXECUTIVE_ROLE_PATTERNS)
    has_interview = contains_any(t, INTERVIEW_PATTERNS)

    if re.search(r"plaud\s+(ceo|founder|co[- ]?founder|executive)", t):
        return True
    if re.search(r"(ceo|founder|co[- ]?founder|executive)\s+(of\s+)?plaud", t):
        return True
    if has_nathan and (has_interview or has_role or has_plaud):
        return True
    if has_plaud and has_role and has_interview:
        return True
    return False


def classify_row(row: pd.Series, cols: Dict[str, str | None]) -> Tuple[str, str, str]:
    title = value(row, cols["title"])
    hit = value(row, cols["hit"])
    content = value(row, cols.get("content"))
    source = " ".join([value(row, cols["source"]), value(row, cols["url"])])
    title_l = title.lower()
    hit_l = hit.lower()
    content_l = content.lower()
    combined_l = f"{title_l} {hit_l} {content_l}"

    if source_is_medium_or_substack(source):
        return "Remove", "来源为 Medium 或 Substack，按硬性规则剔除", "R1"

    title_has_plaud = "plaud" in title_l
    hit_has_plaud = "plaud" in hit_l or "plaud" in content_l
    listicle_title = is_listicle(title)

    if is_executive_interview_or_quote(combined_l):
        return "Keep", "公司高管采访/专访/播客/观点引用类内容，命中 CEO / Nathan / interview 等信号", "K3"

    if not title_has_plaud and not listicle_title:
        return "Remove", "标题不含 Plaud，且不是榜单/推荐/测评集合类文章", "R2"

    if "plaud" in combined_l and contains_any(combined_l, BOILERPLATE_PATTERNS):
        return "Remove", "Plaud 仅出现在广告、赞助、推荐或页面附属信息中", "R5"

    if hit_has_plaud and contains_any(hit_l, COMPARISON_PATTERNS) and not title_has_plaud:
        return "Remove", "Hit Sentence 主要是在用 Plaud 作对比或参照", "R4"

    if title_has_plaud:
        if not hit_has_plaud and len(hit_l) < 25:
            return "Review", "标题含 Plaud，但命中句信息不足，建议人工复核", "V2"
        if contains_any(combined_l, SUBSTANTIVE_PATTERNS) or hit_has_plaud:
            return "Keep", "标题明确包含 Plaud，且内容与 PLAUD 品牌/产品相关", "K1"
        return "Review", "标题含 Plaud，但无法确认文章主体是否为 PLAUD", "V1"

    if listicle_title:
        if hit_has_plaud and contains_any(combined_l, SUBSTANTIVE_PATTERNS):
            return "Keep", "榜单/推荐/测评集合类文章，且 Plaud 被实质介绍", "K2"
        if hit_has_plaud:
            return "Review", "榜单类标题且提到 Plaud，但实质介绍程度需人工确认", "V1"
        return "Remove", "榜单类标题但 Hit Sentence 未体现 Plaud 被推荐或介绍", "R3"

    if hit_has_plaud:
        return "Review", "标题不含 Plaud，但 Hit Sentence 提到 Plaud，需确认文章主体", "V1"

    return "Remove", "未匹配到有效 PLAUD 相关内容", "R2"


def classify_region(country_text: str) -> str:
    raw = str(country_text or "").strip()
    if not raw or raw.lower() == "nan":
        return "other_unassigned"
    parts = re.split(r"[,;/|]+", raw.lower())
    normalized = {p.strip() for p in parts if p.strip()}
    full = raw.lower().strip()
    candidates = normalized | {full}

    if candidates & JP:
        return "jp"
    if candidates & AMERICAS:
        return "americas"
    if candidates & EU:
        return "eu"
    if candidates & APAC_EX_JP:
        return "apac_exclude_jp"
    return "other_unassigned"


def read_csv(uploaded_file) -> pd.DataFrame:
    raw = uploaded_file.getvalue()
    encodings = ["utf-16", "utf-8-sig", "utf-8", "latin1"]
    last_error = None
    for enc in encodings:
        try:
            text = raw.decode(enc)
            sample = text[:4096]
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
                sep = dialect.delimiter
            except Exception:
                sep = "\t" if "\t" in sample else ","
            return pd.read_csv(io.StringIO(text), sep=sep, dtype=str, keep_default_na=False)
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"无法读取 CSV，请检查编码或分隔符。最后错误：{last_error}")


def build_outputs(df: pd.DataFrame, cols: Dict[str, str | None]) -> pd.DataFrame:
    results = df.apply(lambda row: classify_row(row, cols), axis=1, result_type="expand")
    out = df.copy()
    out["Decision"] = results[0]
    out["Reason"] = results[1]
    out["Rule Matched"] = results[2]
    country_col = cols["country"]
    out["Region Group"] = out[country_col].apply(classify_region) if country_col else "other_unassigned"
    return out


def summary_df(classified: pd.DataFrame) -> pd.DataFrame:
    decision_counts = classified["Decision"].value_counts().to_dict()
    keep = classified[classified["Decision"] == "Keep"]
    region_counts = keep["Region Group"].value_counts().to_dict()
    rows = [
        ["Total rows", len(classified)],
        ["Keep", decision_counts.get("Keep", 0)],
        ["Remove", decision_counts.get("Remove", 0)],
        ["Review", decision_counts.get("Review", 0)],
        ["Keep - americas", region_counts.get("americas", 0)],
        ["Keep - jp", region_counts.get("jp", 0)],
        ["Keep - eu", region_counts.get("eu", 0)],
        ["Keep - apac_exclude_jp", region_counts.get("apac_exclude_jp", 0)],
        ["Keep - other_unassigned", region_counts.get("other_unassigned", 0)],
        ["Rule note", "Medium/Substack sources are removed before title/hit-sentence checks"],
        ["Rule note", "Executive interviews/quotes mentioning PLAUD CEO, Nathan, interview, Q&A, or podcast are kept as K3"],
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def to_excel_bytes(classified: pd.DataFrame) -> bytes:
    keep = classified[classified["Decision"] == "Keep"]
    removed = classified[classified["Decision"] == "Remove"]
    review = classified[classified["Decision"] == "Review"]
    sheets = {
        "Summary": summary_df(classified),
        "Keep": keep,
        "Removed": removed,
        "Review": review,
        "All Classified": classified,
        "Americas": keep[keep["Region Group"] == "americas"],
        "JP": keep[keep["Region Group"] == "jp"],
        "EU": keep[keep["Region Group"] == "eu"],
        "APAC exclude JP": keep[keep["Region Group"] == "apac_exclude_jp"],
        "Other Unassigned": keep[keep["Region Group"] == "other_unassigned"],
    }
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, index=False, sheet_name=name[:31])
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for col_idx, column_cells in enumerate(ws.columns, start=1):
                max_len = 10
                for cell in column_cells[:80]:
                    max_len = max(max_len, min(len(str(cell.value or "")), 45))
                ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2
    bio.seek(0)
    return bio.read()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


st.title(APP_TITLE)
st.caption("上传 Meltwater / 媒体监测 CSV 后，自动完成清理、Review 标注、地区拆分，并剔除 Medium / Substack 来源。")

uploaded = st.file_uploader("上传 CSV 文件", type=["csv"])

with st.expander("当前内置规则", expanded=False):
    st.markdown(
        """
- 来源为 Medium 或 Substack：直接 Remove。
- 标题含 Plaud 且内容相关：Keep。
- 高管采访 / 专访 / 播客 / 观点引用类内容，命中 CEO / Nathan / interview / Q&A 等信号：Keep。
- 标题不含 Plaud，但属于 Top / Best / 推荐 / 测评集合类，且 Plaud 被实质介绍：Keep。
- 仅简单提及、作为对比、广告/赞助/推荐链接、无关同名词：Remove。
- 信息不足或边界案例：Review。
- Keep 文章会继续拆分为 americas、jp、eu、apac_exclude_jp、other_unassigned。
        """
    )

if uploaded:
    try:
        df = read_csv(uploaded)
        st.success(f"已读取 {len(df)} 条记录，{len(df.columns)} 个字段。")

        cols = {
            "title": find_column(df, TITLE_FIELDS),
            "hit": find_column(df, HIT_FIELDS),
            "content": find_column(df, CONTENT_FIELDS),
            "source": find_column(df, SOURCE_FIELDS),
            "url": find_column(df, URL_FIELDS),
            "country": find_column(df, COUNTRY_FIELDS),
        }
        st.write("字段识别：", cols)

        if not cols["title"] or not cols["hit"]:
            st.error("未能识别 Title 或 Hit Sentence 字段。请检查表头，或把列名改为 Title / Hit Sentence 后重新上传。")
        else:
            classified = build_outputs(df, cols)
            summary = summary_df(classified)

            st.subheader("处理结果")
            st.dataframe(summary, use_container_width=True, hide_index=True)

            st.subheader("预览：All Classified")
            st.dataframe(classified.head(50), use_container_width=True)

            excel_bytes = to_excel_bytes(classified)
            base = uploaded.name.rsplit(".", 1)[0].replace(" ", "_")
            st.download_button(
                "下载完整 Excel 报告",
                data=excel_bytes,
                file_name=f"{base}_cleaned_media_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            c1, c2, c3, c4 = st.columns(4)
            keep = classified[classified["Decision"] == "Keep"]
            removed = classified[classified["Decision"] == "Remove"]
            review = classified[classified["Decision"] == "Review"]
            with c1:
                st.download_button("下载 Keep CSV", to_csv_bytes(keep), f"{base}_keep.csv", "text/csv")
            with c2:
                st.download_button("下载 Removed CSV", to_csv_bytes(removed), f"{base}_removed.csv", "text/csv")
            with c3:
                st.download_button("下载 Review CSV", to_csv_bytes(review), f"{base}_review.csv", "text/csv")
            with c4:
                st.download_button("下载全量标注 CSV", to_csv_bytes(classified), f"{base}_classified.csv", "text/csv")

    except Exception as exc:
        st.exception(exc)
else:
    st.info("请先上传 CSV。")
