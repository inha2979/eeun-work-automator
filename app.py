import io
import re
import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Eeun Work Automator", page_icon="🛠️", layout="wide")


# -----------------------------
# Common utilities
# -----------------------------
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


def file_bytes(uploaded):
    uploaded.seek(0)
    data = uploaded.read()
    uploaded.seek(0)
    return data


def read_table(uploaded, sheet_name=None, header_row=0):
    """Read CSV/XLS/XLSX with selectable header row and sheet."""
    data = file_bytes(uploaded)
    name = uploaded.name.lower()
    bio = io.BytesIO(data)
    if name.endswith(".csv"):
        try:
            return pd.read_csv(bio, header=header_row)
        except UnicodeDecodeError:
            bio.seek(0)
            return pd.read_csv(bio, header=header_row, encoding="cp949")
    return pd.read_excel(bio, sheet_name=sheet_name if sheet_name is not None else 0, header=header_row)


def excel_sheet_names(uploaded):
    name = uploaded.name.lower()
    if not (name.endswith(".xlsx") or name.endswith(".xls")):
        return []
    data = file_bytes(uploaded)
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


def flexible_loader(uploaded, key_prefix, label="파일 구조 설정"):
    """Allow different workbook structures by selecting sheet and header row."""
    sheet = None
    if uploaded.name.lower().endswith((".xlsx", ".xls")):
        sheets = excel_sheet_names(uploaded)
        if len(sheets) > 1:
            sheet = st.selectbox(f"{label} · 시트", sheets, key=f"{key_prefix}_sheet")
        elif sheets:
            sheet = sheets[0]

    header_row_1based = st.number_input(
        f"{label} · 실제 열 이름이 있는 행",
        min_value=1,
        max_value=50,
        value=1,
        step=1,
        help="예: 1행에 제목, 2행은 안내문, 3행에 '이름/전화번호/인스타ID'가 있으면 3을 선택",
        key=f"{key_prefix}_header",
    )
    try:
        df = read_table(uploaded, sheet_name=sheet, header_row=int(header_row_1based) - 1)
        df = df.dropna(axis=1, how="all")
        return df
    except Exception as e:
        st.error(f"파일을 읽지 못했습니다: {e}")
        return None


def find_candidate_columns(df, keywords):
    cols = []
    for c in df.columns:
        cs = str(c).lower().replace(" ", "").replace("_", "")
        if any(k.lower().replace(" ", "").replace("_", "") in cs for k in keywords):
            cols.append(c)
    return cols


def auto_col(df, keywords, fallback="선택 안 함"):
    found = find_candidate_columns(df, keywords)
    return found[0] if found else fallback


def option_index(options, value):
    try:
        return options.index(value)
    except ValueError:
        return 0


def lottery_seed(seed_text):
    if seed_text.strip():
        digest = hashlib.sha256(seed_text.strip().encode("utf-8")).hexdigest()
        return int(digest[:8], 16)
    return int(datetime.now().strftime("%Y%m%d%H%M%S")) % (2**32 - 1)


def allowed_set(text):
    return {x.strip().lower() for x in text.split(",") if x.strip()}


def append_excluded(excluded_frames, before_df, after_df, reason):
    removed_idx = before_df.index.difference(after_df.index)
    if len(removed_idx):
        tmp = before_df.loc[removed_idx].copy()
        tmp["제외사유"] = reason
        excluded_frames.append(tmp)


# -----------------------------
# UI shell
# -----------------------------
st.title("🛠️ Eeun Work Automator")
st.caption("반복적인 이벤트/엑셀 업무를 빠르게 정리하는 로컬 업무 자동화 툴")

menu = st.sidebar.radio(
    "메뉴",
    ["🧹 Excel Cleaner", "🔗 Excel Matcher", "🎁 Event Lottery"],
)


# -----------------------------
# Excel Cleaner
# -----------------------------
if menu == "🧹 Excel Cleaner":
    st.header("Excel Cleaner")
    st.write("전화번호, 인스타그램 ID, 공백, 중복, 빈 행 등을 한 번에 정리합니다.")
    file = st.file_uploader("Excel 또는 CSV 파일 업로드", type=["xlsx", "xls", "csv"], key="cleaner")

    if file:
        df = flexible_loader(file, "cleaner", "파일 구조")
        if df is not None:
            st.subheader("미리보기")
            st.dataframe(df.head(30), use_container_width=True)

            cols = list(df.columns)
            phone_guess = auto_col(df, ["전화번호", "휴대폰", "핸드폰", "phone", "mobile", "연락처"])
            insta_guess = auto_col(df, ["instagram", "인스타", "snsid", "계정", "아이디", "id"])

            c1, c2 = st.columns(2)
            with c1:
                phone_options = ["선택 안 함"] + cols
                phone_col = st.selectbox(
                    "전화번호 열",
                    phone_options,
                    index=option_index(phone_options, phone_guess),
                )
                phone_format = st.radio("전화번호 출력 형식", ["숫자만", "하이픈 포함"], horizontal=True)
                insta_options = ["선택 안 함"] + cols
                insta_col = st.selectbox(
                    "Instagram ID 열",
                    insta_options,
                    index=option_index(insta_options, insta_guess),
                )
            with c2:
                strip_all = st.checkbox("모든 텍스트 앞뒤 공백 제거", value=True)
                drop_blank_rows = st.checkbox("완전히 빈 행 제거", value=True)
                duplicate_cols = st.multiselect("중복 판단 기준 열", cols)
                mark_duplicates = st.checkbox("중복 여부 열 추가", value=True)
                drop_duplicates = st.checkbox("중복 행 제거", value=False)

            if phone_guess != "선택 안 함" or insta_guess != "선택 안 함":
                detected = []
                if phone_guess != "선택 안 함":
                    detected.append(f"전화번호 → {phone_guess}")
                if insta_guess != "선택 안 함":
                    detected.append(f"Instagram ID → {insta_guess}")
                st.info("자동 감지: " + " / ".join(detected))

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
                    if mark_duplicates:
                        out["중복여부"] = dup_mask.map({True: "중복", False: ""})
                if drop_duplicates and duplicate_cols:
                    out = out.drop_duplicates(subset=duplicate_cols, keep="first")

                summary = pd.DataFrame({
                    "항목": ["원본 행 수", "정리 후 행 수", "중복 의심 행 수"],
                    "개수": [len(df), len(out), int(dup_mask.sum())],
                })
                st.success("정리 완료")
                st.dataframe(summary, use_container_width=True)
                st.dataframe(out.head(100), use_container_width=True)
                xbytes = to_excel_bytes({"cleaned": out, "summary": summary})
                st.download_button("📥 cleaned_result.xlsx", xbytes, "cleaned_result.xlsx")


# -----------------------------
# Excel Matcher
# -----------------------------
elif menu == "🔗 Excel Matcher":
    st.header("Excel Matcher")
    st.write("당첨자 ID처럼 한 파일의 키를 기준으로 다른 파일에서 이름·전화번호·댓글 등을 가져옵니다.")
    left = st.file_uploader("기준 파일 업로드 (예: 당첨자 목록)", type=["xlsx", "xls", "csv"], key="left")
    right = st.file_uploader("정보 파일 업로드 (예: 전체 참여자)", type=["xlsx", "xls", "csv"], key="right")

    if left and right:
        ldf = flexible_loader(left, "matcher_left", "기준 파일 구조")
        rdf = flexible_loader(right, "matcher_right", "정보 파일 구조")
        if ldf is not None and rdf is not None:
            c1, c2 = st.columns(2)
            with c1:
                st.caption("기준 파일")
                st.dataframe(ldf.head(20), use_container_width=True)
                lguess = auto_col(ldf, ["instagram", "인스타", "아이디", "id", "계정"], list(ldf.columns)[0])
                lkey = st.selectbox("기준 파일 매칭 열", list(ldf.columns), index=option_index(list(ldf.columns), lguess), key="lkey")
            with c2:
                st.caption("정보 파일")
                st.dataframe(rdf.head(20), use_container_width=True)
                rguess = auto_col(rdf, ["instagram", "인스타", "아이디", "id", "계정"], list(rdf.columns)[0])
                rkey = st.selectbox("정보 파일 매칭 열", list(rdf.columns), index=option_index(list(rdf.columns), rguess), key="rkey")

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


# -----------------------------
# Event Lottery
# -----------------------------
else:
    st.header("Event Lottery")
    st.write("참여자 정리 → 팔로워 대조 → 과거 당첨자 제외 → 랜덤 추첨까지 한 번에 처리합니다.")

    participants_file = st.file_uploader("① 참여자 파일", type=["xlsx", "xls", "csv"], key="participants")
    followers_file = st.file_uploader("② 팔로워 파일 (선택)", type=["xlsx", "xls", "csv", "json"], key="followers")
    winners_file = st.file_uploader("③ 과거 당첨자 파일 (선택)", type=["xlsx", "xls", "csv"], key="winners")

    if participants_file:
        pdf = flexible_loader(participants_file, "lottery_participants", "참여자 파일 구조")
        if pdf is not None:
            st.dataframe(pdf.head(25), use_container_width=True)
            pcols = list(pdf.columns)

            likely_id = find_candidate_columns(pdf, ["instagram", "인스타", "아이디", "id", "계정", "sns"])
            default_id = likely_id[0] if likely_id else pcols[0]
            id_col = st.selectbox("참여자 Instagram ID 열", pcols, index=option_index(pcols, default_id))

            consent_guess = auto_col(pdf, ["개인정보", "동의", "consent"])
            valid_guess = auto_col(pdf, ["유효응답", "유효", "검증", "자격", "valid", "eligible"])

            c1, c2, c3 = st.columns(3)
            with c1:
                dedupe = st.checkbox("중복 ID 제외", value=True)
                consent_options = ["선택 안 함"] + pcols
                consent_col = st.selectbox(
                    "개인정보 동의 열",
                    consent_options,
                    index=option_index(consent_options, consent_guess),
                )
                consent_values = st.text_input("포함할 동의 값", value="동의,O,예,Y,TRUE")
            with c2:
                valid_options = ["선택 안 함"] + pcols
                valid_col = st.selectbox(
                    "추첨 자격/검증 열",
                    valid_options,
                    index=option_index(valid_options, valid_guess),
                )
                valid_values = st.text_input("포함할 자격 값", value="O,예,Y,TRUE,유효")
                winner_count = st.number_input("당첨자 수", min_value=1, value=50, step=1)
            with c3:
                reserve_count = st.number_input("예비 당첨자 수", min_value=0, value=10, step=1)
                seed_text = st.text_input("재현용 Seed 문구 (선택)", placeholder="예: momntalk-2026-09")

            follower_ids = None
            follower_df = None
            if followers_file:
                if followers_file.name.lower().endswith(".json"):
                    try:
                        follower_df = pd.read_json(io.BytesIO(file_bytes(followers_file)))
                    except Exception as e:
                        st.error(f"팔로워 JSON을 읽지 못했습니다: {e}")
                else:
                    follower_df = flexible_loader(followers_file, "lottery_followers", "팔로워 파일 구조")
                if follower_df is not None:
                    fcols = list(follower_df.columns)
                    fguess = auto_col(follower_df, ["instagram", "인스타", "아이디", "id", "계정", "sns"], fcols[0])
                    fkey = st.selectbox("팔로워 파일 ID 열", fcols, index=option_index(fcols, fguess))
                    follower_ids = set(follower_df[fkey].map(normalize_instagram_id)) - {""}

            past_winner_ids = None
            wdf = None
            if winners_file:
                wdf = flexible_loader(winners_file, "lottery_winners", "과거 당첨자 파일 구조")
                if wdf is not None:
                    wcols = list(wdf.columns)
                    wguess = auto_col(wdf, ["instagram", "인스타", "아이디", "id", "계정", "sns"], wcols[0])
                    wkey = st.selectbox("과거 당첨자 ID 열", wcols, index=option_index(wcols, wguess))
                    past_winner_ids = set(wdf[wkey].map(normalize_instagram_id)) - {""}

            if not followers_file:
                st.info("팔로워 파일이 없으면 팔로우 여부 필터는 건너뜁니다. 아래 검증용 ID 목록을 내려받아 O/X 확인 후 다시 사용할 수 있습니다.")
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
                excluded_frames = []
                process_rows = []

                def stage(label, before_n, after_n):
                    process_rows.append({"단계": label, "제외": before_n - after_n, "잔여": after_n})

                process_rows.append({"단계": "전체 참여자", "제외": 0, "잔여": len(work)})

                before = work.copy()
                work = work[work["__id"] != ""]
                append_excluded(excluded_frames, before, work, "ID 공란")
                stage("ID 공란", len(before), len(work))

                if dedupe:
                    before = work.copy()
                    dup_mask = work.duplicated(subset=["__id"], keep="first")
                    if dup_mask.any():
                        tmp = work[dup_mask].copy()
                        tmp["제외사유"] = "중복 ID"
                        excluded_frames.append(tmp)
                    work = work[~dup_mask]
                    stage("중복 ID", len(before), len(work))

                if consent_col != "선택 안 함":
                    before = work.copy()
                    allowed = allowed_set(consent_values)
                    work = work[work[consent_col].astype(str).str.strip().str.lower().isin(allowed)]
                    append_excluded(excluded_frames, before, work, "개인정보 미동의/허용값 아님")
                    stage("개인정보 동의", len(before), len(work))

                if valid_col != "선택 안 함":
                    before = work.copy()
                    allowed = allowed_set(valid_values)
                    work = work[work[valid_col].astype(str).str.strip().str.lower().isin(allowed)]
                    append_excluded(excluded_frames, before, work, "추첨 자격 미충족")
                    stage("추첨 자격", len(before), len(work))

                if follower_ids is not None:
                    before = work.copy()
                    work = work[work["__id"].isin(follower_ids)]
                    append_excluded(excluded_frames, before, work, "비팔로워")
                    stage("팔로워", len(before), len(work))

                if past_winner_ids is not None:
                    before = work.copy()
                    work = work[~work["__id"].isin(past_winner_ids)]
                    append_excluded(excluded_frames, before, work, "과거 당첨자")
                    stage("과거 당첨자", len(before), len(work))

                summary = pd.DataFrame(process_rows)
                eligible_n = len(work)
                total_needed = int(winner_count + reserve_count)

                st.subheader("추첨 대상 정리 결과")
                st.dataframe(summary, use_container_width=True)

                excluded_counts = summary.iloc[1:][["단계", "제외"]]
                parts = [f"{row['단계']} {int(row['제외'])}명" for _, row in excluded_counts.iterrows() if int(row["제외"]) > 0]
                report_sentence = f"총 {len(pdf)}명의 참여자 중 " + (", ".join(parts) + "을/를 제외해 " if parts else "") + f"최종 {eligible_n}명을 대상으로 추첨했습니다."
                st.text_area("보고용 문장", report_sentence, height=90)

                if eligible_n < total_needed:
                    st.error(
                        f"최종 추첨 대상은 {eligible_n}명인데 당첨자+예비 당첨자를 {total_needed}명으로 설정했습니다. "
                        f"합계를 {eligible_n}명 이하로 줄여주세요."
                    )
                else:
                    seed = lottery_seed(seed_text)
                    picked = work.sample(n=total_needed, random_state=seed)
                    winners = picked.iloc[:int(winner_count)].drop(columns=["__id"])
                    reserves = picked.iloc[int(winner_count):].drop(columns=["__id"])
                    eligible = work.drop(columns=["__id"])
                    excluded = pd.concat(excluded_frames, ignore_index=True) if excluded_frames else pd.DataFrame(columns=list(pdf.columns) + ["제외사유"])
                    if "__id" in excluded.columns:
                        excluded = excluded.drop(columns=["__id"])

                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    meta = pd.DataFrame({
                        "항목": ["추첨 시각", "Seed", "당첨자 수", "예비 당첨자 수", "최종 추첨 대상"],
                        "값": [now, seed, int(winner_count), int(reserve_count), eligible_n],
                    })
                    report_df = pd.DataFrame({"보고문장": [report_sentence]})

                    st.success(f"추첨 완료 — 최종 대상 {eligible_n}명")
                    st.subheader("당첨자")
                    st.dataframe(winners.reset_index(drop=True), use_container_width=True)
                    st.subheader("예비 당첨자")
                    if len(reserves):
                        st.dataframe(reserves.reset_index(drop=True), use_container_width=True)
                    else:
                        st.caption("예비 당첨자 수를 0명으로 설정했습니다.")

                    with st.expander("제외 대상 및 제외 사유 보기"):
                        st.dataframe(excluded, use_container_width=True)

                    xbytes = to_excel_bytes({
                        "당첨자": winners,
                        "예비당첨자": reserves,
                        "최종추첨대상": eligible,
                        "제외대상": excluded,
                        "추첨과정": summary,
                        "추첨정보": meta,
                        "보고문장": report_df,
                    })
                    st.download_button("📥 lottery_result.xlsx", xbytes, "lottery_result.xlsx")
