# Eeun Work Automator

반복적인 이벤트 운영 및 Excel 정리 업무를 자동화하는 Streamlit 기반 로컬 웹앱입니다.

## 주요 기능

### 1. Excel Cleaner
- 전화번호 형식 통일
- Instagram ID 정규화
- 공백 제거
- 빈 행 제거
- 중복 탐지/제거
- 정리 결과 Excel 다운로드

### 2. Excel Matcher
- 두 Excel/CSV 파일을 공통 키로 매칭
- 당첨자 ID를 기준으로 이름, 전화번호, 댓글 등 가져오기
- 매칭되지 않은 데이터 별도 확인

### 3. Event Lottery
- 참여자 ID 정규화
- 중복 참여자 제외
- 개인정보 동의자 필터
- 유효 응답/추첨 자격 필터
- 팔로워 목록 대조
- 과거 당첨자 제외
- 당첨자/예비 당첨자 랜덤 추첨
- Seed 및 추첨 시각 기록
- 추첨 과정 로그 포함 결과 Excel 생성

## 팔로워 파일이 없는 경우
Instagram에 반복 로그인하여 팔로워 여부를 자동 크롤링하는 기능은 계정 제한 및 서비스 정책 문제 때문에 기본 기능으로 포함하지 않았습니다.

대신 Event Lottery 메뉴에서 `팔로우 수동검증용 목록`을 다운로드할 수 있습니다.

1. 중복 제거된 Instagram ID 목록 다운로드
2. 팔로우 여부를 O/X 등으로 확인
3. 참여자 원본에 확인 결과 열을 추가
4. Event Lottery의 `추첨 자격/검증 열` 기능으로 O만 포함해 추첨

Meta에서 공식적으로 제공하는 계정 데이터 내보내기 파일을 확보할 수 있다면, 해당 팔로워 목록 파일을 직접 업로드하여 자동 대조할 수 있습니다.

## 설치

Python 3.10 이상 권장

```bash
pip install -r requirements.txt
```

## 실행

```bash
streamlit run app.py
```

브라우저가 자동으로 열리지 않는 경우 터미널에 표시되는 Local URL을 브라우저에서 엽니다.

## GitHub 업로드 예시

```bash
git init
git add .
git commit -m "Initial release: Excel cleaner, matcher and event lottery"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/eeun-work-automator.git
git push -u origin main
```

## 주의

실제 회사 데이터에는 이름, 전화번호, Instagram ID 등 개인정보가 포함될 수 있습니다. 예제 데이터나 실제 참여자 파일은 GitHub에 업로드하지 마세요.

`.gitignore`를 활용해 `.xlsx`, `.csv` 등의 실제 데이터 파일을 저장소에서 제외하는 것을 권장합니다.
