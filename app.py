import io
import re
import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Eeun Work Automator", page_icon="🛠️", layout="wide")


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_instagram_id(value):
    s = normalize_text(value).lower()
    s = re.sub(r"^https?://(www\.)?instagram\.com/", "", s)
    s = s.split("?")[0].strip("/")
    s = s.lstrip("@").strip()
    return s


def normalize_phone(value, output_format="digits"):
    if pd.isna(value):
        return ""
    s = re.sub(r"\D", "", str(value))
    if not s:
        return ""
    if s.startswith("82"):
        s = "0" + s[2:]
    if not s.startswith("0") and len(s) in (9, 10):
        s = "0" + s

    if output_format == "digits":
        return s
    if output_format == "hyphen":
        if len(s) == 11:
            return f"{s[:3]}-{s[3:7]}-{s[7:]}"
        if len(s) == 10:
            if s.startswith("02"):
                return f"{s[:2]}-{s[2:6]}-{s[6:]}"
            return f"{s[:3]}-{s[3:6]}-{s[6:]}"
    return s


def to_excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = re.sub(r"[\\/*?:\[\]]", "_", name)[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
    output.seek(0)
    return output.getvalue()


def read_uploaded(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(file)
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, encoding="cp949")
    return pd.read_excel(file)


def find_candidate_columns(df, keywords):
    cols = []
    for c in df.columns:
        cs = str(c).lower().replace(" ", "")
        if any(k.lower().replace(" ", "") in cs for k in keywords):
            cols.append(c)
    return cols


def lottery_seed(seed_text):
    if seed_text.strip():
        digest = hashlib.sha256(seed_text.strip().encode("utf-8")).hexdigest()
        return int(digest[:8], 16)
    return int(datetime.now().strftime("%Y%m%d%H%M%S")) % (2**32 - 1)


st.title("🛠️ Eeun Work Automator")
st.caption("반복적인 이벤트/엑셀 업무를 빠르게 정리하는 로컬 업무 자동화 툴")

menu = st.sidebar.radio(
    "메뉴",
    ["🧹 Excel Cleaner", "🔗 Excel Matcher", "🎁 Event Lottery"],
)

if menu == "🧹 Excel Cleaner":
    st.header("Excel Cleaner")
    st.write("전화번호, 인스타그램 ID, 공백, 중복, 빈 행 등을 한 번에 정리합니다.")
    file = st.file_uploader("Excel 또는 CSV 파일 업로드", type=["xlsx", "xls", "csv"], key="cleaner")

    if file:
        df = read_uploaded(file)
        st.subheader("미리보기")
        st.dataframe(df.head(30), use_container_width=True)

        cols = list(df.columns)
        c1, c2 = st.columns(2)
        with c1:
            phone_col = st.selectbox("전화번호 열", ["선택 안 함"] + cols)
            phone_format = st.radio("전화번호 출력 형식", ["숫자만", "하이픈 포함"], horizontal=True)
            insta_col = st.selectbox("Instagram ID 열", ["선택 안 함"] + cols)
        with c2:
            strip_all = st.checkbox("모든 텍스트 앞뒤 공백 제거", value=True)
            drop_blank_rows = st.checkbox("완전히 빈 행 제거", value=True)
            duplicate_cols = st.multiselect("중복 판단 기준 열", cols)
            drop_duplicates = st.checkbox("중복 행 제거", value=False)

        if st.button("정리 실행", type="primary"):
            out = df.copy()
            if strip_all:
                for c in out.select_dtypes(include="object").columns:
                    out[c] = out[c].map(lambda x: x.strip() if isinstance(x, str) else x)
            if phone_col != "선택 안 함":
                fmt = "digits" if phone_format == "숫자만" else "hyphen"
                out[phone_col] = out[phone_col].map(lambda x: normalize_phone(x, fmt))
            if insta_col != "선택 안 함":
                out[insta_col] = out[insta_col].map(normalize_instagram_id)
            if drop_blank_rows:
                out = out.dropna(how="all")
            dup_mask = pd.Series(False, index=out.index)
            if duplicate_cols:
                dup_mask = out.duplicated(subset=duplicate_cols, keep=False)
            if drop_duplicates and duplicate_cols:
                out = out.drop_duplicates(subset=duplicate_cols, keep="first")

            summary = pd.DataFrame({
                "항목": ["원본 행 수", "정리 후 행 수", "중복 의심 행 수"],
                "개수": [len(df), len(out), int(dup_mask.sum())],
            })
            st.success("정리 완료")
            st.dataframe(summary, use_container_width=True)
            st.dataframe(out.head(50), use_container_width=True)
            xbytes = to_excel_bytes({"cleaned": out, "summary": summary})
            st.download_button("📥 cleaned_result.xlsx", xbytes, "cleaned_result.xlsx")

elif menu == "🔗 Excel Matcher":
    st.header("Excel Matcher")
    st.write("당첨자 ID처럼 한 파일의 키를 기준으로 다른 파일에서 이름·전화번호·댓글 등을 가져옵니다.")
    left = st.file_uploader("기준 파일 업로드 (예: 당첨자 목록)", type=["xlsx", "xls", "csv"], key="left")
    right = st.file_uploader("정보 파일 업로드 (예: 전체 참여자)", type=["xlsx", "xls", "csv"], key="right")

    if left and right:
        ldf = read_uploaded(left)
        rdf = read_uploaded(right)
        c1, c2 = st.columns(2)
        with c1:
            st.caption("기준 파일")
            st.dataframe(ldf.head(20), use_container_width=True)
            lkey = st.selectbox("기준 파일 매칭 열", list(ldf.columns), key="lkey")
        with c2:
            st.caption("정보 파일")
            st.dataframe(rdf.head(20), use_container_width=True)
            rkey = st.selectbox("정보 파일 매칭 열", list(rdf.columns), key="rkey")

        normalize_as_insta = st.checkbox("매칭 키를 Instagram ID 형식으로 정규화 (@ 제거/소문자/URL 제거)", value=True)
        fields = st.multiselect("가져올 열", [c for c in rdf.columns if c != rkey], default=[c for c in rdf.columns if c != rkey][:4])

        if st.button("매칭 실행", type="primary"):
            L = ldf.copy()
            R = rdf.copy()
            lk = "__match_key__"
            if normalize_as_insta:
                L[lk] = L[lkey].map(normalize_instagram_id)
                R[lk] = R[rkey].map(normalize_instagram_id)
            else:
                L[lk] = L[lkey].astype(str).str.strip()
                R[lk] = R[rkey].astype(str).str.strip()

            R = R[[lk] + fields].drop_duplicates(subset=[lk], keep="first")
            merged = L.merge(R, on=lk, how="left", indicator=True)
            merged["매칭상태"] = merged["_merge"].map({"both": "매칭됨", "left_only": "정보 없음", "right_only": ""})
            merged = merged.drop(columns=[lk, "_merge"])
            missing = merged[merged["매칭상태"] == "정보 없음"]
            st.success(f"매칭 완료: {len(merged)-len(missing)}건 / 정보 없음 {len(missing)}건")
            st.dataframe(merged.head(100), use_container_width=True)
            xbytes = to_excel_bytes({"matched": merged, "missing": missing})
            st.download_button("📥 matched_result.xlsx", xbytes, "matched_result.xlsx")

else:
    st.header("Event Lottery")
    st.write("참여자 정리 → 팔로워 대조 → 과거 당첨자 제외 → 랜덤 추첨까지 한 번에 처리합니다.")

    participants_file = st.file_uploader("① 참여자 파일", type=["xlsx", "xls", "csv"], key="participants")
    followers_file = st.file_uploader("② 팔로워 파일 (선택)", type=["xlsx", "xls", "csv", "json"], key="followers")
    winners_file = st.file_uploader("③ 과거 당첨자 파일 (선택)", type=["xlsx", "xls", "csv"], key="winners")

    if participants_file:
        pdf = read_uploaded(participants_file)
        st.dataframe(pdf.head(25), use_container_width=True)
        pcols = list(pdf.columns)

        likely_id = find_candidate_columns(pdf, ["instagram", "인스타", "아이디", "id", "계정"])
        default_index = pcols.index(likely_id[0]) if likely_id else 0
        id_col = st.selectbox("참여자 Instagram ID 열", pcols, index=default_index)

        c1, c2, c3 = st.columns(3)
        with c1:
            dedupe = st.checkbox("중복 ID 제외", value=True)
            consent_col = st.selectbox("개인정보 동의 열", ["선택 안 함"] + pcols)
            consent_values = st.text_input("포함할 동의 값", value="동의,O,예,Y,TRUE")
        with c2:
            valid_col = st.selectbox("추첨 자격/검증 열", ["선택 안 함"] + pcols)
            valid_values = st.text_input("포함할 자격 값", value="O,예,Y,TRUE,유효")
            winner_count = st.number_input("당첨자 수", min_value=1, value=50, step=1)
        with c3:
            reserve_count = st.number_input("예비 당첨자 수", min_value=0, value=10, step=1)
            seed_text = st.text_input("재현용 Seed 문구 (선택)", placeholder="예: momntalk-2026-09")

        follower_ids = None
        if followers_file:
            if followers_file.name.lower().endswith(".json"):
                raw = pd.read_json(followers_file)
                follower_df = raw
            else:
                follower_df = read_uploaded(followers_file)
            fcols = list(follower_df.columns)
            fkey = st.selectbox("팔로워 파일 ID 열", fcols)
            follower_ids = set(follower_df[fkey].map(normalize_instagram_id))

        past_winner_ids = None
        if winners_file:
            wdf = read_uploaded(winners_file)
            wcols = list(wdf.columns)
            wkey = st.selectbox("과거 당첨자 ID 열", wcols)
            past_winner_ids = set(wdf[wkey].map(normalize_instagram_id))

        if not followers_file:
            st.info("팔로워 파일이 없으면 팔로우 여부 필터는 건너뜁니다. 먼저 '검증용 ID 목록'을 내려받아 수동 확인 후, O/X 열을 추가해서 다시 올리는 방식을 추천합니다.")
            unique_ids = pd.DataFrame({"instagram_id": sorted(set(pdf[id_col].map(normalize_instagram_id)) - {""})})
            unique_ids["팔로우확인"] = ""
            st.download_button(
                "📋 팔로우 수동검증용 목록 다운로드",
                to_excel_bytes({"follow_check": unique_ids}),
                "follow_check_list.xlsx",
            )

        if st.button("추첨 실행", type="primary"):
            work = pdf.copy()
            work["__id"] = work[id_col].map(normalize_instagram_id)
            log = []
            log.append(("전체 참여자", len(work)))

            work = work[work["__id"] != ""]
            log.append(("ID 공란 제외 후", len(work)))

            if dedupe:
                work = work.drop_duplicates(subset=["__id"], keep="first")
                log.append(("중복 ID 제외 후", len(work)))

            if consent_col != "선택 안 함":
                allowed = {x.strip().lower() for x in consent_values.split(",") if x.strip()}
                work = work[work[consent_col].astype(str).str.strip().str.lower().isin(allowed)]
                log.append(("개인정보 동의 필터 후", len(work)))

            if valid_col != "선택 안 함":
                allowed = {x.strip().lower() for x in valid_values.split(",") if x.strip()}
                work = work[work[valid_col].astype(str).str.strip().str.lower().isin(allowed)]
                log.append(("추첨 자격 필터 후", len(work)))

            if follower_ids is not None:
                work = work[work["__id"].isin(follower_ids)]
                log.append(("팔로워 필터 후", len(work)))

            if past_winner_ids is not None:
                work = work[~work["__id"].isin(past_winner_ids)]
                log.append(("과거 당첨자 제외 후", len(work)))

            total_needed = int(winner_count + reserve_count)
            if len(work) < total_needed:
                st.error(f"추첨 가능 인원 {len(work)}명으로, 필요한 {total_needed}명보다 적습니다.")
            else:
                seed = lottery_seed(seed_text)
                picked = work.sample(n=total_needed, random_state=seed)
                winners = picked.iloc[:int(winner_count)].drop(columns=["__id"])
                reserves = picked.iloc[int(winner_count):].drop(columns=["__id"])
                eligible = work.drop(columns=["__id"])
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                summary = pd.DataFrame(log, columns=["단계", "인원"])
                meta = pd.DataFrame({
                    "항목": ["추첨 시각", "Seed", "당첨자 수", "예비 당첨자 수"],
                    "값": [now, seed, int(winner_count), int(reserve_count)],
                })
                st.success(f"추첨 완료 — 최종 대상 {len(work)}명")
                st.dataframe(summary, use_container_width=True)
                st.subheader("당첨자")
                st.dataframe(winners, use_container_width=True)
                xbytes = to_excel_bytes({
                    "winners": winners,
                    "reserves": reserves,
                    "eligible_pool": eligible,
                    "process_log": summary,
                    "lottery_meta": meta,
                })
                st.download_button("📥 lottery_result.xlsx", xbytes, "lottery_result.xlsx")
