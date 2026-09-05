import io
import re
import hashlib
from difflib import SequenceMatcher
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
    output = io.io.BytesIO()
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
    ["🧹 Excel Cleaner", "🔗 Excel Matcher", "📚 다중 엑셀 통합기", "✅ 파일 납품 전 체크기", "🧾 정책 중복 검사", "📚 이은북 카드뉴스 기획안", "✉️ 업무 보고/메일 생성기", "🎁 Event Lottery"],
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
            st.dataframe(df.head(30), width="stretch")

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
                st.dataframe(summary, width="stretch")
                st.dataframe(out.head(100), width="stretch")
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
                st.dataframe(ldf.head(20), width="stretch")
                lguess = auto_col(ldf, ["instagram", "인스타", "아이디", "id", "계정"], list(ldf.columns)[0])
                lkey = st.selectbox("기준 파일 매칭 열", list(ldf.columns), index=option_index(list(ldf.columns), lguess), key="lkey")
            with c2:
                st.caption("정보 파일")
                st.dataframe(rdf.head(20), width="stretch")
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
                st.dataframe(merged.head(100), width="stretch")
                xbytes = to_excel_bytes({"matched": merged, "missing": missing})
                st.download_button("📥 matched_result.xlsx", xbytes, "matched_result.xlsx")



# -----------------------------
# Multi Excel Merger
# -----------------------------
elif menu == "📚 다중 엑셀 통합기":
    st.header("📚 다중 엑셀 통합기")
    st.write("여러 Excel/CSV 파일을 한 번에 업로드해 하나의 통합 파일로 합칩니다.")

    uploaded_files = st.file_uploader(
        "합칠 파일 업로드",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="여러 파일을 동시에 선택할 수 있습니다.",
    )

    if uploaded_files:
        st.caption(f"업로드된 파일: {len(uploaded_files)}개")

        merge_mode = st.radio(
            "통합 방식",
            ["열 이름 기준으로 이어붙이기", "공통 열만 남겨 이어붙이기"],
            horizontal=True,
            help="열 이름 기준: 없는 열은 빈칸으로 채웁니다. / 공통 열만: 모든 파일에 공통으로 존재하는 열만 남깁니다.",
        )

        add_source = st.checkbox(
            "원본 파일명 열 추가",
            value=True,
            help="통합 결과에 각 행이 어느 파일에서 왔는지 '원본파일' 열로 표시합니다.",
        )

        sheet_choice = st.selectbox(
            "Excel 시트 처리 방식",
            ["각 파일의 첫 번째 시트만", "모든 시트 합치기"],
            help="CSV는 시트가 없으므로 그대로 처리됩니다.",
        )

        if st.button("🔗 파일 통합 실행", type="primary", width="stretch"):
            frames = []
            summary_rows = []
            errors = []

            for uf in uploaded_files:
                try:
                    name = uf.name
                    suffix = name.lower().split(".")[-1]

                    if suffix == "csv":
                        try:
                            df = pd.read_csv(uf)
                        except UnicodeDecodeError:
                            uf.seek(0)
                            df = pd.read_csv(uf, encoding="cp949")

                        if add_source:
                            df.insert(0, "원본파일", name)

                        frames.append(df)
                        summary_rows.append({
                            "파일명": name,
                            "시트": "-",
                            "행 수": len(df),
                            "열 수": len(df.columns),
                        })
                    else:
                        uf.seek(0)
                        xls = pd.ExcelFile(uf)
                        target_sheets = [xls.sheet_names[0]] if sheet_choice == "각 파일의 첫 번째 시트만" else xls.sheet_names

                        for sh in target_sheets:
                            df = pd.read_excel(xls, sheet_name=sh)
                            if add_source:
                                df.insert(0, "원본파일", name)
                                if sheet_choice == "모든 시트 합치기":
                                    df.insert(1, "원본시트", sh)

                            frames.append(df)
                            summary_rows.append({
                                "파일명": name,
                                "시트": sh,
                                "행 수": len(df),
                                "열 수": len(df.columns),
                            })
                except Exception as e:
                    errors.append(f"{uf.name}: {e}")

            if frames:
                if merge_mode == "공통 열만 남겨 이어붙이기":
                    common = set(frames[0].columns)
                    for df in frames[1:]:
                        common &= set(df.columns)

                    # Preserve first file's column order
                    ordered_common = [c for c in frames[0].columns if c in common]
                    merged = pd.concat([df[ordered_common] for df in frames], ignore_index=True)
                else:
                    merged = pd.concat(frames, ignore_index=True, sort=False)

                st.success(f"총 {len(frames)}개 파일/시트를 합쳐 {len(merged):,}행으로 통합했습니다.")

                with st.expander("📊 파일별 통합 요약", expanded=True):
                    st.dataframe(pd.DataFrame(summary_rows), width="stretch")

                st.markdown("#### 통합 결과 미리보기")
                st.dataframe(merged.head(100), width="stretch")

                if errors:
                    st.warning("일부 파일은 처리하지 못했습니다.")
                    for err in errors:
                        st.write(f"- {err}")

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    merged.to_excel(writer, index=False, sheet_name="통합결과")
                    pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="통합요약")
                    if errors:
                        pd.DataFrame({"오류": errors}).to_excel(writer, index=False, sheet_name="오류")
                output.seek(0)

                st.download_button(
                    "⬇️ 통합 Excel 다운로드",
                    data=output.getvalue(),
                    file_name="다중엑셀_통합결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )
            else:
                st.error("통합할 수 있는 파일이 없습니다.")


# -----------------------------
# Pre-delivery File Checker
# -----------------------------
elif menu == "✅ 파일 납품 전 체크기":
    st.header("✅ 파일 납품 전 체크기")
    st.write("전달하기 전 Excel/CSV 파일의 누락·중복·빈칸·전화번호·필수 열 등을 빠르게 점검합니다.")

    check_file = st.file_uploader(
        "검수할 파일 업로드",
        type=["xlsx", "xls", "csv"],
        key="delivery_check_file",
    )

    if check_file:
        # Load file
        suffix = check_file.name.lower().split(".")[-1]
        if suffix == "csv":
            try:
                cdf = pd.read_csv(check_file)
            except UnicodeDecodeError:
                check_file.seek(0)
                cdf = pd.read_csv(check_file, encoding="cp949")
            selected_sheet = "-"
        else:
            check_file.seek(0)
            xls = pd.ExcelFile(check_file)
            selected_sheet = st.selectbox("검수할 시트", xls.sheet_names)
            cdf = pd.read_excel(xls, sheet_name=selected_sheet)

        st.caption(f"현재 파일: {check_file.name} / {len(cdf):,}행 × {len(cdf.columns)}열")

        all_cols = list(cdf.columns)

        st.markdown("### 검수 조건")
        c1, c2 = st.columns(2)
        with c1:
            required_cols_text = st.text_input(
                "반드시 있어야 하는 열 (선택)",
                placeholder="예: 이름,전화번호,인스타그램ID",
                help="쉼표(,)로 구분하세요.",
            )
            required_value_cols = st.multiselect(
                "빈칸이 있으면 안 되는 열 (선택)",
                all_cols,
                help="선택한 열에서 빈 셀이 있으면 오류로 표시합니다.",
            )
            duplicate_cols = st.multiselect(
                "중복 검사 기준 열 (선택)",
                all_cols,
                help="예: 전화번호 또는 Instagram ID",
            )
        with c2:
            phone_col = st.selectbox(
                "전화번호 형식 검사 열 (선택)",
                ["선택 안 함"] + all_cols,
                help="숫자 10~11자리 기준으로 형식을 확인합니다.",
            )
            id_col = st.selectbox(
                "Instagram ID 형식 검사 열 (선택)",
                ["선택 안 함"] + all_cols,
                help="@, URL, 공백이 섞여 있는 값을 확인합니다.",
            )
            check_blank_rows = st.checkbox("완전히 빈 행 검사", value=True)

        if st.button("🔍 납품 전 검수 실행", type="primary", width="stretch"):
            issues = []
            row_issues = pd.DataFrame(index=cdf.index)
            row_issues["검수결과"] = ""
            row_issues["문제내용"] = ""

            def add_row_issue(mask, msg):
                if mask.any():
                    idxs = row_issues.index[mask]
                    for idx in idxs:
                        current = row_issues.at[idx, "문제내용"]
                        row_issues.at[idx, "문제내용"] = f"{current}; {msg}".strip("; ")
                        row_issues.at[idx, "검수결과"] = "확인필요"

            # Required columns
            required_cols = [x.strip() for x in required_cols_text.split(",") if x.strip()]
            missing_cols = [c for c in required_cols if c not in all_cols]
            if missing_cols:
                issues.append(("필수 열 누락", f"{', '.join(missing_cols)}"))

            # Blank required values
            for col in required_value_cols:
                mask = cdf[col].isna() | (cdf[col].astype(str).str.strip() == "")
                count = int(mask.sum())
                if count:
                    issues.append((f"빈칸 - {col}", f"{count}건"))
                    add_row_issue(mask, f"{col} 빈칸")

            # Duplicates
            if duplicate_cols:
                dup_mask = cdf.duplicated(subset=duplicate_cols, keep=False)
                dup_count = int(dup_mask.sum())
                if dup_count:
                    issues.append(("중복값", f"{dup_count}행"))
                    add_row_issue(dup_mask, f"{', '.join(duplicate_cols)} 기준 중복")

            # Phone
            if phone_col != "선택 안 함":
                phone_digits = cdf[phone_col].astype(str).str.replace(r"\D", "", regex=True)
                phone_mask = ~phone_digits.str.len().isin([10, 11]) & cdf[phone_col].notna()
                phone_count = int(phone_mask.sum())
                if phone_count:
                    issues.append(("전화번호 형식", f"{phone_count}건"))
                    add_row_issue(phone_mask, f"{phone_col} 전화번호 형식 확인")

            # Instagram
            if id_col != "선택 안 함":
                ids = cdf[id_col].astype(str)
                ig_mask = (
                    ids.str.contains(r"https?://|instagram\.com|@", case=False, regex=True, na=False)
                    | ids.str.contains(r"\s", regex=True, na=False)
                )
                ig_count = int(ig_mask.sum())
                if ig_count:
                    issues.append(("Instagram ID 형식", f"{ig_count}건"))
                    add_row_issue(ig_mask, f"{id_col} 형식 확인")

            # Completely blank rows
            if check_blank_rows:
                blank_row_mask = cdf.isna().all(axis=1)
                blank_count = int(blank_row_mask.sum())
                if blank_count:
                    issues.append(("완전히 빈 행", f"{blank_count}행"))
                    add_row_issue(blank_row_mask, "완전히 빈 행")

            result_df = cdf.copy()
            result_df.insert(0, "검수결과", row_issues["검수결과"].replace("", "정상"))
            result_df.insert(1, "문제내용", row_issues["문제내용"])

            if not issues:
                st.success("✅ 현재 설정 기준으로 발견된 문제가 없습니다. 전달 전 최종 육안 확인만 해주세요.")
            else:
                st.warning(f"⚠️ 확인이 필요한 항목이 {len(issues)}종류 발견됐습니다.")
                st.dataframe(
                    pd.DataFrame(issues, columns=["검사항목", "결과"]),
                    width="stretch",
                )

            issue_rows = result_df[result_df["검수결과"] == "확인필요"]
            st.markdown("#### 확인이 필요한 행")
            if len(issue_rows):
                st.dataframe(issue_rows, width="stretch")
            else:
                st.info("행 단위로 확인할 문제는 없습니다.")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="검수결과")
                pd.DataFrame(issues, columns=["검사항목", "결과"]).to_excel(
                    writer, index=False, sheet_name="검수요약"
                )
            output.seek(0)

            st.download_button(
                "⬇️ 검수 결과 Excel 다운로드",
                data=output.getvalue(),
                file_name="파일_납품전_검수결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )


# -----------------------------
# Policy Duplicate Checker
# -----------------------------
elif menu == "🧾 정책 중복 검사":
    st.header("정책 중복 검사")
    st.write("공모전·정책 제안 엑셀에서 완전 중복과 유사한 내용을 찾아 비교표로 정리합니다.")
    st.caption("엑셀 구조가 달라도 시트와 헤더 행, 제목/본문 열을 직접 선택할 수 있습니다.")

    policy_file = st.file_uploader(
        "정책/공모전 제안 파일 업로드",
        type=["xlsx", "xls", "csv"],
        key="policy_duplicate",
    )

    if policy_file:
        pdf = flexible_loader(policy_file, "policy_duplicate", "파일 구조")
        if pdf is not None and len(pdf.columns):
            st.subheader("미리보기")
            st.dataframe(pdf.head(30), width="stretch")
            pcols = list(pdf.columns)

            id_guess = auto_col(pdf, ["접수번호", "제안번호", "번호", "id", "no"], "선택 안 함")
            title_guess = auto_col(pdf, ["정책명", "제안명", "제목", "아이디어명", "공모명", "title"], "선택 안 함")
            body_candidates = find_candidate_columns(
                pdf,
                ["정책내용", "제안내용", "내용", "상세", "설명", "필요성", "배경", "목적", "기대효과", "body", "description"],
            )

            c1, c2 = st.columns(2)
            with c1:
                id_options = ["선택 안 함"] + pcols
                id_col = st.selectbox(
                    "식별용 열 (선택)",
                    id_options,
                    index=option_index(id_options, id_guess),
                    help="접수번호 등이 있으면 결과표에 함께 표시합니다.",
                )
                title_options = ["선택 안 함"] + pcols
                title_col = st.selectbox(
                    "제목 열 (선택)",
                    title_options,
                    index=option_index(title_options, title_guess),
                )
            with c2:
                default_bodies = body_candidates[:2]
                body_cols = st.multiselect(
                    "비교할 내용 열",
                    pcols,
                    default=default_bodies,
                    help="정책 내용, 필요성, 기대효과 등 여러 열을 동시에 선택할 수 있습니다.",
                )
                threshold = st.slider(
                    "유사 중복 기준",
                    min_value=0.60,
                    max_value=1.00,
                    value=0.78,
                    step=0.01,
                    help="0.78이면 문장 표현이 일부 달라도 상당히 유사한 제안을 후보로 표시합니다.",
                )

            include_title = st.checkbox("제목도 유사도 계산에 포함", value=True)
            ignore_short = st.number_input(
                "너무 짧은 내용 제외 (최소 글자 수)",
                min_value=0,
                max_value=500,
                value=20,
                step=5,
            )

            selected_cols = list(body_cols)
            if include_title and title_col != "선택 안 함":
                selected_cols = [title_col] + selected_cols

            if not selected_cols:
                st.warning("제목 열 또는 비교할 내용 열을 하나 이상 선택해주세요.")
            else:
                if len(pdf) > 1500:
                    st.warning("행이 1,500개를 넘으면 유사도 비교에 시간이 걸릴 수 있습니다. 먼저 필요한 행만 필터링하면 더 빠릅니다.")

                def normalize_policy_text(value):
                    text = normalize_text(value).lower()
                    text = re.sub(r"\s+", " ", text)
                    text = re.sub(r"[^0-9a-zA-Z가-힣 ]", "", text)
                    return text.strip()

                def row_policy_text(row):
                    parts = [normalize_text(row.get(c, "")) for c in selected_cols]
                    return normalize_policy_text(" ".join(p for p in parts if p))

                if st.button("중복 검사 실행", type="primary"):
                    work = pdf.copy().reset_index(drop=True)
                    work["__source_row"] = work.index + 2
                    work["__compare_text"] = work.apply(row_policy_text, axis=1)
                    work["__text_len"] = work["__compare_text"].str.len()
                    valid = work[work["__text_len"] >= int(ignore_short)].copy()

                    exact_map = {}
                    for idx, text in valid["__compare_text"].items():
                        exact_map.setdefault(text, []).append(idx)
                    pair_rows = []
                    indices = list(valid.index)
                    for pos, i in enumerate(indices):
                        a = valid.at[i, "__compare_text"]
                        la = len(a)
                        for j in indices[pos + 1:]:
                            b = valid.at[j, "__compare_text"]
                            lb = len(b)
                            if not a or not b:
                                continue
                            # Very different lengths are unlikely to be meaningful duplicates.
                            if min(la, lb) / max(la, lb) < 0.45:
                                continue
                            exact = a == b
                            score = 1.0 if exact else SequenceMatcher(None, a, b, autojunk=False).ratio()
                            if score >= threshold:
                                row = {
                                    "유형": "완전중복" if exact else "유사중복",
                                    "유사도": round(score * 100, 1),
                                    "A_엑셀행": int(work.at[i, "__source_row"]),
                                    "B_엑셀행": int(work.at[j, "__source_row"]),
                                }
                                if id_col != "선택 안 함":
                                    row["A_ID"] = work.at[i, id_col]
                                    row["B_ID"] = work.at[j, id_col]
                                if title_col != "선택 안 함":
                                    row["A_제목"] = work.at[i, title_col]
                                    row["B_제목"] = work.at[j, title_col]
                                row["A_내용미리보기"] = " ".join(str(work.at[i, c]) for c in selected_cols if not pd.isna(work.at[i, c]))[:300]
                                row["B_내용미리보기"] = " ".join(str(work.at[j, c]) for c in selected_cols if not pd.isna(work.at[j, c]))[:300]
                                pair_rows.append(row)

                    pairs = pd.DataFrame(pair_rows)
                    if len(pairs):
                        pairs = pairs.sort_values(["유사도", "유형"], ascending=[False, True]).reset_index(drop=True)

                    result = pdf.copy().reset_index(drop=True)
                    # pandas 3.x에서는 문자열 열에 숫자를 나중에 넣으면 TypeError가 날 수 있으므로
                    # 처음부터 각 결과 열의 dtype을 명시합니다.
                    result["중복검사_최고유사도"] = pd.Series(pd.NA, index=result.index, dtype="Float64")
                    result["중복검사_상대행"] = pd.Series(pd.NA, index=result.index, dtype="Int64")
                    result["중복검사_판정"] = pd.Series(pd.NA, index=result.index, dtype="string")

                    if len(pairs):
                        best = {}
                        for _, r in pairs.iterrows():
                            aidx = int(r["A_엑셀행"]) - 2
                            bidx = int(r["B_엑셀행"]) - 2
                            score = float(r["유사도"])
                            typ = r["유형"]
                            if aidx not in best or score > best[aidx][0]:
                                best[aidx] = (score, int(r["B_엑셀행"]), typ)
                            if bidx not in best or score > best[bidx][0]:
                                best[bidx] = (score, int(r["A_엑셀행"]), typ)
                        for idx, (score, other_row, typ) in best.items():
                            result.at[idx, "중복검사_최고유사도"] = score
                            result.at[idx, "중복검사_상대행"] = other_row
                            result.at[idx, "중복검사_판정"] = typ

                    exact_pairs = pairs[pairs["유형"] == "완전중복"].copy() if len(pairs) else pd.DataFrame()
                    similar_pairs = pairs[pairs["유형"] == "유사중복"].copy() if len(pairs) else pd.DataFrame()
                    summary = pd.DataFrame({
                        "항목": ["전체 제안", "비교 대상", "완전 중복 쌍", "유사 중복 쌍", "유사도 기준"],
                        "값": [len(pdf), len(valid), len(exact_pairs), len(similar_pairs), f"{threshold*100:.0f}% 이상"],
                    })

                    st.success(f"검사 완료 — 완전 중복 {len(exact_pairs)}쌍 / 유사 중복 {len(similar_pairs)}쌍")
                    st.dataframe(summary, width="stretch")
                    st.subheader("중복 의심 쌍")
                    if len(pairs):
                        st.dataframe(pairs, width="stretch")
                    else:
                        st.info("현재 기준에서는 중복 또는 유사 중복 후보가 발견되지 않았습니다.")

                    settings = pd.DataFrame({
                        "항목": ["식별 열", "제목 열", "비교 열", "제목 포함", "유사도 기준", "최소 글자 수"],
                        "값": [id_col, title_col, ", ".join(map(str, body_cols)), include_title, threshold, int(ignore_short)],
                    })
                    xbytes = to_excel_bytes({
                        "검사결과": result,
                        "중복의심쌍": pairs,
                        "완전중복": exact_pairs,
                        "유사중복": similar_pairs,
                        "요약": summary,
                        "검사설정": settings,
                    })
                    st.download_button("📥 policy_duplicate_result.xlsx", xbytes, "policy_duplicate_result.xlsx")


# -----------------------------
# Eeunbook Card News Prompt
# -----------------------------
elif menu == "📚 이은북 카드뉴스 기획안":
    st.header("📚 이은북 카드뉴스 기획안")
    st.write("도서 정보를 입력하면 ChatGPT에 그대로 붙여넣을 수 있는 카드뉴스 기획 프롬프트를 만들어줍니다.")
    st.info("이 기능은 AI API를 직접 호출하지 않습니다. 여기서는 프롬프트만 만들고, 완성된 내용을 본인의 ChatGPT에 붙여넣어 사용합니다.")

    book_title = st.text_input("도서명", placeholder="예: AI 광고 콘텐츠 제작법")
    book_intro = st.text_area(
        "책 소개 / 참고 원문",
        height=280,
        placeholder="출판사 책 소개, 상세페이지 내용, 보도자료 등 카드뉴스 기획에 참고할 내용을 붙여넣으세요.",
    )

    c1, c2 = st.columns(2)
    with c1:
        card_count = st.number_input("카드 수", min_value=4, max_value=15, value=7, step=1,
                                     help="표지와 마지막 마무리 페이지를 포함한 전체 카드 수입니다.")
        target_reader = st.text_input("주요 타깃 독자 (선택)", placeholder="예: AI 활용에 관심 있는 마케터·기획자")
    with c2:
        tone = st.selectbox("톤", ["깔끔한 정보형", "흥미 유도형", "친근한 소개형", "전문적인 소개형"])
        emphasis = st.text_area("특히 강조할 내용 (선택)", height=100,
                                placeholder="예: 실무에서 바로 따라 할 수 있다는 점")

    extra_request = st.text_area("추가 요청사항 (선택)", height=90,
                                 placeholder="예: 2P는 질문형으로 시작 / 저자 소개는 제외")

    if st.button("✨ ChatGPT용 프롬프트 만들기", type="primary", width="stretch"):
        if not book_title.strip():
            st.error("도서명을 입력해주세요.")
        elif not book_intro.strip():
            st.error("책 소개 / 참고 원문을 입력해주세요.")
        else:
            target_text = target_reader.strip() or "제공된 자료를 바탕으로 가장 적절한 핵심 독자를 판단"
            emphasis_text = emphasis.strip() or "제공된 자료에서 가장 중요한 도서의 차별점과 독자 효용을 판단"
            extra_text = extra_request.strip() or "없음"

            prompt = f"""당신은 출판사 이은북의 SNS 카드뉴스를 기획하는 콘텐츠 기획자입니다.

아래 도서 자료를 바탕으로 총 {int(card_count)}장의 도서 소개 카드뉴스 기획안을 작성해주세요.

[도서명]
{book_title.strip()}

[주요 타깃 독자]
{target_text}

[원하는 톤]
{tone}

[특히 강조할 내용]
{emphasis_text}

[추가 요청사항]
{extra_text}

[책 소개 / 참고 원문]
{book_intro.strip()}

[작성 원칙]
- 반드시 제공된 자료를 근거로 작성하고, 자료에 없는 사실·수치·저자 정보·책의 내용을 임의로 만들지 마세요.
- 단순히 책 소개를 잘게 나누지 말고 처음부터 끝까지 하나의 카드뉴스로 자연스럽게 읽히는 흐름을 설계하세요.
- 1P는 표지입니다. 책의 핵심 매력을 짧고 강하게 보여주는 헤드카피를 제안하세요.
- 초반에는 독자의 고민이나 관심사와 책을 연결해 흥미를 유도하세요.
- 중간에는 책의 핵심 내용과 차별점을 여러 포인트로 나누어 전달하세요.
- 마지막 페이지는 도서명과 핵심 메시지를 다시 연결해 자연스럽게 마무리하세요.
- 실제 카드에 들어갈 문장은 짧고 직관적으로 작성하세요.
- 과도한 광고 문구, 근거 없는 최상급 표현, 뻔한 AI식 표현은 피하세요.
- 페이지마다 같은 표현과 메시지를 반복하지 마세요.
- 자연스러운 한국어를 사용하세요.

[출력 형식]
먼저 전체 기획 의도를 2~3문장으로 설명해주세요.

그 다음 1P부터 {int(card_count)}P까지 각 페이지를 아래 형식으로 작성해주세요.
- 페이지 역할:
- 헤드카피:
- 본문:
- 이미지/디자인 방향:

마지막으로 아래 항목도 작성해주세요.
- 전체 카드뉴스 흐름 요약:
- 제작 시 주의할 점:
"""

            st.success("완성됐습니다. 아래 프롬프트를 전체 복사해서 ChatGPT에 붙여넣으면 됩니다.")
            st.text_area("완성된 ChatGPT 프롬프트", value=prompt, height=600)
            st.download_button(
                "📄 프롬프트 TXT로 저장",
                data=prompt.encode("utf-8"),
                file_name=f"{book_title.strip()}_카드뉴스_기획_프롬프트.txt",
                mime="text/plain",
                width="stretch",
            )


# -----------------------------
# Work Report / Email Generator
# -----------------------------
elif menu == "✉️ 업무 보고/메일 생성기":
    st.header("✉️ 업무 보고/메일 생성기")
    st.write("완료한 업무와 전달할 내용을 입력하면 메일·카톡 보고 문장을 바로 만들어줍니다.")

    report_type = st.radio(
        "만들 문서",
        ["📧 이메일", "💬 카카오톡/메신저", "📝 짧은 업무 보고"],
        horizontal=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        recipient = st.selectbox(
            "받는 사람",
            ["팀장님", "차장님", "선배님", "사장님", "담당자님", "직접 입력"],
        )
        if recipient == "직접 입력":
            recipient_custom = st.text_input("호칭 직접 입력", placeholder="예: 부장님")
            recipient_text = recipient_custom.strip() or "담당자님"
        else:
            recipient_text = recipient

        sender = st.text_input(
            "보내는 사람",
            value="김인하 인턴",
            help="이메일 서명에 사용할 이름/직함입니다.",
        )

    with c2:
        subject = st.text_input(
            "메일 제목 (선택)",
            placeholder="예: [맘앤톡] 온리원 이벤트 당첨자 파일 전달드립니다",
            help="이메일 선택 시 제목 추천에 사용합니다.",
        )
        tone = st.selectbox(
            "톤",
            ["공손하고 간결하게", "조금 더 부드럽게", "아주 짧게", "진행상황 중심"],
        )

    task_name = st.text_input(
        "업무명",
        placeholder="예: 맘앤톡 온리원 이벤트 당첨자 정리 / AI 인물편집 교정",
    )

    completed = st.text_area(
        "완료한 내용",
        height=120,
        placeholder="예: 전체 참여자 검토 후 중복·미동의자를 제외하고 최종 추첨을 완료했습니다.",
    )

    attachments = st.text_area(
        "첨부/전달 파일 (선택)",
        height=85,
        placeholder="예: 당첨자 목록.xlsx, 추첨 영상.mp4",
    )

    remaining = st.text_area(
        "남은 업무 / 다음 진행사항 (선택)",
        height=90,
        placeholder="예: 나머지 2건은 월요일에 마무리할 예정입니다.",
    )

    request_text = st.text_area(
        "확인 요청 / 질문 (선택)",
        height=90,
        placeholder="예: 확인 부탁드립니다. / 예약 발행으로 진행해도 될지 확인 부탁드립니다.",
    )

    deadline = st.text_input(
        "일정/기한 (선택)",
        placeholder="예: 월요일 오전까지 / 오늘 중 / 9월 8일까지",
    )

    if st.button("✨ 보고 문장 만들기", type="primary", width="stretch"):
        task = task_name.strip()
        done = completed.strip()
        attach = attachments.strip()
        remain = remaining.strip()
        req = request_text.strip()
        due = deadline.strip()

        if not task:
            st.error("업무명을 입력해주세요.")
        elif not done:
            st.error("완료한 내용을 입력해주세요.")
        else:
            # 입력문을 그대로 붙이지 않고, 보고/메일체로 다듬기 위한 간단한 규칙
            import re

            def clean_fragment(s):
                s = (s or "").strip()
                s = re.sub(r"\s+", " ", s)
                return s.rstrip(".!?")

            def polite_sentence(s, mode="done"):
                s = clean_fragment(s)
                if not s:
                    return ""

                replacements = [
                    (r"\b함$", "했습니다"),
                    (r"\b완료함$", "완료했습니다"),
                    (r"\b진행함$", "진행했습니다"),
                    (r"\b전달함$", "전달드렸습니다"),
                    (r"\b보냄$", "전달드렸습니다"),
                    (r"\b확인함$", "확인했습니다"),
                    (r"\b정리함$", "정리했습니다"),
                    (r"\b추첨 함$", "추첨했습니다"),
                    (r"\b추첨함$", "추첨했습니다"),
                    (r"\b제거함$", "제외했습니다"),
                    (r"\b제외함$", "제외했습니다"),
                    (r"\b예정임$", "예정입니다"),
                    (r"\b해야함$", "진행할 예정입니다"),
                ]
                for pat, repl in replacements:
                    s = re.sub(pat, repl, s)

                # 짧은 메모체를 자연스러운 보고체로 보정
                if not re.search(r"(습니다|드립니다|했습니다|입니다|예정입니다|부탁드립니다)$", s):
                    if mode == "request":
                        s = s + " 확인 부탁드립니다"
                    elif mode == "remain":
                        s = s + " 진행할 예정입니다"
                    else:
                        s = s + " 진행했습니다"

                return s + "."

            done_s = polite_sentence(done, "done")
            remain_s = polite_sentence(remain, "remain")
            req_s = polite_sentence(req, "request") if req else ""

            if tone == "아주 짧게":
                intro_done = done_s
            elif tone == "진행상황 중심":
                intro_done = f"{task} 진행 상황 공유드립니다.\n\n{done_s}"
            elif tone == "조금 더 부드럽게":
                intro_done = f"{task} 관련하여 진행한 내용 전달드립니다.\n\n{done_s}"
            else:
                intro_done = f"{task} 관련하여 {done_s}"

            if report_type == "📧 이메일":
                if subject.strip():
                    subject_line = subject.strip()
                else:
                    subject_line = f"[{task}] 작업 결과 전달드립니다"

                body_parts = [
                    f"{recipient_text}, 안녕하세요.",
                    f"{sender}입니다.",
                    "",
                    intro_done,
                ]

                if attach:
                    body_parts += ["", f"관련 파일은 아래와 같이 함께 전달드립니다.\n{attach}"]

                if remain:
                    remain_line = remain_s
                    if due and due not in remain_line:
                        remain_line = f"{remain_line} 일정은 {due}입니다."
                    body_parts += ["", remain_line]
                elif due:
                    body_parts += ["", f"관련 일정은 {due}입니다."]

                if req:
                    body_parts += ["", req_s]
                else:
                    body_parts += ["", "확인 부탁드립니다."]

                body_parts += ["", "감사합니다.", f"{sender.split()[0] if sender.strip() else '김인하'} 드림"]

                generated = "\n".join(body_parts)

                st.markdown("#### 추천 메일 제목")
                st.code(subject_line, language=None)
                st.markdown("#### 메일 본문")
                st.text_area("복사용 메일", generated, height=360)

            elif report_type == "💬 카카오톡/메신저":
                parts = [f"{recipient_text}, 안녕하세요!"]
                parts.append(intro_done)

                if attach:
                    parts.append(f"관련 파일도 함께 전달드렸습니다: {attach}")
                if remain:
                    line = remain_s
                    if due and due not in line:
                        line += f" 일정은 {due}입니다."
                    parts.append(line)
                elif due:
                    parts.append(f"관련 일정은 {due}입니다.")

                if req:
                    parts.append(req_s)
                else:
                    parts.append("확인 부탁드립니다. 감사합니다!")

                generated = "\n\n".join(parts)
                st.text_area("복사용 메신저 문장", generated, height=280)

            else:
                parts = [f"{task}: {done_s}"]
                if attach:
                    parts.append(f"전달 파일: {attach}")
                if remain:
                    parts.append(f"남은 업무: {remain_s}")
                if due:
                    parts.append(f"일정: {due}")
                if req:
                    parts.append(f"확인 요청: {req_s}")

                generated = "\n".join(parts)
                st.text_area("복사용 업무 보고", generated, height=230)

            st.download_button(
                "📄 문장 TXT로 저장",
                data=generated.encode("utf-8"),
                file_name="업무보고_문장.txt",
                mime="text/plain",
                width="stretch",
            )


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
            st.dataframe(pdf.head(25), width="stretch")
            pcols = list(pdf.columns)

            likely_id = find_candidate_columns(pdf, ["instagram", "인스타", "아이디", "id", "계정", "sns"])
            default_id = likely_id[0] if likely_id else pcols[0]
            id_col = st.selectbox("참여자 Instagram ID 열", pcols, index=option_index(pcols, default_id))

            consent_guess = auto_col(pdf, ["개인정보", "동의", "consent"])
            valid_guess = auto_col(pdf, ["유효응답", "유효", "검증", "자격", "valid", "eligible"])

            st.markdown("### 추첨 조건 설정")
            st.caption("아래에서 어떤 참여자를 추첨 대상에 포함할지 정합니다. 필요 없는 조건은 '선택 안 함'으로 두면 됩니다.")

            c1, c2, c3 = st.columns(3)
            with c1:
                dedupe = st.checkbox(
                    "중복 Instagram ID 제외",
                    value=True,
                    help="같은 Instagram ID가 여러 번 참여했을 경우 한 명으로 처리합니다.",
                )
                consent_options = ["선택 안 함"] + pcols
                consent_col = st.selectbox(
                    "개인정보 동의 여부가 적힌 열",
                    consent_options,
                    index=option_index(consent_options, consent_guess),
                    help="예: 개인정보동의, 개인정보 수집동의. 개인정보 동의를 추첨 조건으로 사용하지 않으면 '선택 안 함'을 고르세요.",
                )
                consent_values = st.text_input(
                    "동의한 것으로 인정할 값",
                    value="동의,O,예,Y,TRUE",
                    help="선택한 열에서 이 값에 해당하는 사람만 추첨 대상에 포함합니다. 여러 값은 쉼표(,)로 구분하세요.",
                )
            with c2:
                valid_options = ["선택 안 함"] + pcols
                valid_col = st.selectbox(
                    "추첨 가능 여부가 적힌 열",
                    valid_options,
                    index=option_index(valid_options, valid_guess),
                    help="예: 유효응답, 참여조건충족, 검증결과. 별도의 자격 조건이 없으면 '선택 안 함'을 고르세요.",
                )
                valid_values = st.text_input(
                    "추첨 가능으로 인정할 값",
                    value="O,예,Y,TRUE,유효",
                    help="선택한 열에서 이 값에 해당하는 사람만 추첨 대상에 포함합니다. 여러 값은 쉼표(,)로 구분하세요.",
                )
                winner_count = st.number_input(
                    "최종 당첨 인원",
                    min_value=1,
                    value=50,
                    step=1,
                    help="실제로 최종 당첨자로 선정할 인원입니다.",
                )
            with c3:
                reserve_count = st.number_input(
                    "예비 당첨 인원",
                    min_value=0,
                    value=10,
                    step=1,
                    help="최종 당첨자가 조건을 충족하지 못하거나 연락이 되지 않을 때 사용할 예비 인원입니다.",
                )
                seed_text = st.text_input(
                    "같은 추첨 결과를 다시 만들기 위한 기준값 (선택)",
                    placeholder="예: momntalk-2026-09",
                    help="같은 데이터와 같은 문구로 다시 실행하면 동일한 무작위 추첨 결과를 재현할 수 있습니다. 비워두어도 됩니다.",
                )

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

            manual_follow_mode = False
            verified_follow_file = None
            extra_candidates = 50
            if not followers_file:
                st.warning("팔로워 파일이 없어도 진행할 수 있습니다. 전체 참여자를 확인하지 않고, 먼저 후보 순번을 만든 뒤 그 후보만 Instagram에서 O/X 확인하는 방식을 권장합니다.")
                follow_mode = st.radio(
                    "팔로워 파일이 없을 때 처리 방식",
                    ["선추첨 → 후보만 수동검증 (권장)", "팔로우 조건 없이 추첨"],
                    horizontal=True,
                )
                manual_follow_mode = follow_mode.startswith("선추첨")
                if manual_follow_mode:
                    extra_candidates = st.number_input(
                        "추가 검증 후보 수",
                        min_value=0,
                        value=max(10, int(winner_count)),
                        step=1,
                        help="팔로워 파일이 없을 때 비팔로워가 나올 경우를 대비해, 당첨자와 예비 당첨자 외에 추가로 확인할 후보 수입니다.",
                    )
                    total_verify = int(winner_count) + int(reserve_count) + int(extra_candidates)
                    st.info(
                        f"현재 설정: 최종 당첨 {int(winner_count)}명 + 예비 {int(reserve_count)}명 "
                        f"+ 추가 검증 {int(extra_candidates)}명 = 최대 {total_verify}명을 먼저 뽑아 팔로우 여부를 확인합니다."
                    )
                    verified_follow_file = st.file_uploader(
                        "④ O/X 확인을 끝낸 검증 파일 다시 업로드 (최종 확정용)",
                        type=["xlsx", "xls", "csv"], key="verified_follow",
                    )
                    st.caption("1차로 후보 파일을 만든 뒤, `팔로우확인` 열에 O/X를 입력해 이곳에 다시 올리면 순번대로 최종 당첨자를 확정합니다.")

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
                st.dataframe(summary, width="stretch")

                excluded_counts = summary.iloc[1:][["단계", "제외"]]
                parts = [f"{row['단계']} {int(row['제외'])}명" for _, row in excluded_counts.iterrows() if int(row["제외"]) > 0]
                report_sentence = f"총 {len(pdf)}명의 참여자 중 " + (", ".join(parts) + "을/를 제외해 " if parts else "") + f"최종 {eligible_n}명을 대상으로 추첨했습니다."
                st.text_area("보고용 문장", report_sentence, height=90)

                # 팔로워 파일이 없을 때: 후보를 먼저 순번화하고, 확인된 O/X 파일로 최종 확정
                if not followers_file and manual_follow_mode:
                    seed = lottery_seed(seed_text)
                    candidate_n = min(eligible_n, total_needed + int(extra_candidates))
                    if candidate_n < total_needed:
                        st.error(f"기본 조건을 통과한 후보가 {eligible_n}명뿐이라 당첨자+예비 {total_needed}명을 확보할 수 없습니다.")
                    else:
                        ranked = work.sample(n=candidate_n, random_state=seed).copy().reset_index(drop=True)
                        ranked.insert(0, "검증순번", range(1, len(ranked) + 1))
                        ranked["팔로우확인"] = ""
                        ranked = ranked.drop(columns=["__id"])
                        st.info(f"팔로우 확인이 필요한 후보 {candidate_n}명만 뽑았습니다. 아래 파일에서 Instagram 팔로우 여부를 O/X로 확인하세요.")
                        st.dataframe(ranked.head(30), width="stretch")
                        st.download_button(
                            "📋 1차 팔로우 검증 후보.xlsx 다운로드",
                            to_excel_bytes({"팔로우검증후보": ranked, "추첨과정": summary}),
                            "follow_verification_candidates.xlsx",
                        )

                    if verified_follow_file is not None:
                        checked = read_table(verified_follow_file, header_row=0)
                        if "팔로우확인" not in checked.columns:
                            st.error("업로드한 검증 파일에 `팔로우확인` 열이 없습니다. 1차 후보 파일을 사용해 O/X를 입력해주세요.")
                        else:
                            if "검증순번" in checked.columns:
                                checked = checked.sort_values("검증순번")
                            ok_values = {"o", "y", "yes", "true", "팔로우", "확인", "1"}
                            confirmed = checked[checked["팔로우확인"].astype(str).str.strip().str.lower().isin(ok_values)].copy()
                            need = int(winner_count + reserve_count)
                            if len(confirmed) < need:
                                st.warning(f"팔로워로 확인된 후보가 {len(confirmed)}명입니다. 필요한 {need}명보다 {need-len(confirmed)}명 부족합니다. 추가 후보를 생성해 더 확인해주세요.")
                            else:
                                final_winners = confirmed.iloc[:int(winner_count)].copy()
                                final_reserves = confirmed.iloc[int(winner_count):need].copy()
                                st.success(f"팔로우 검증 완료 — 당첨자 {len(final_winners)}명 / 예비 {len(final_reserves)}명 확정")
                                st.subheader("최종 당첨자")
                                st.dataframe(final_winners, width="stretch")
                                st.subheader("최종 예비 당첨자")
                                st.dataframe(final_reserves, width="stretch")
                                st.download_button(
                                    "📥 팔로우 검증 최종 결과.xlsx",
                                    to_excel_bytes({
                                        "최종당첨자": final_winners,
                                        "최종예비당첨자": final_reserves,
                                        "검증전체": checked,
                                        "팔로워확인": confirmed,
                                        "추첨과정": summary,
                                    }),
                                    "lottery_verified_result.xlsx",
                                )
                    st.stop()

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
                    st.dataframe(winners.reset_index(drop=True), width="stretch")
                    st.subheader("예비 당첨자")
                    if len(reserves):
                        st.dataframe(reserves.reset_index(drop=True), width="stretch")
                    else:
                        st.caption("예비 당첨자 수를 0명으로 설정했습니다.")

                    with st.expander("제외 대상 및 제외 사유 보기"):
                        st.dataframe(excluded, width="stretch")

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
