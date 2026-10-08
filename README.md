# 크제비(K-JEBI) 추천 공고 리포트

매일 오는 크제비 맞춤검색 알림메일(`no-replay@kjebi.com`)을 읽어서 입찰공고·사전규격·발주계획을 뽑고, 우리 회사 기준으로 점수를 매겨 **강력 추천 / 검토 추천 / 참고 / 제외**로 나눈 리포트를 만듭니다.

## 동작 방식

1. Gmail(IMAP)에서 크제비 알림메일을 가져옵니다.
2. 메일 HTML에서 공고번호, 제목, 예산(추정가격·배정예산액), 수요기관, 참여지역, 마감일을 추출합니다.
3. 정정공고(`-001` 등)는 최신 차수만, `[재발송]` 메일과 같은 제목의 중복 공고는 하나만 남깁니다.
4. `config.yaml` 기준으로 점수를 매깁니다.
   - 제목 키워드 가점/감점 (UX·플랫폼·챗봇 +, 유지관리·연구용역 − 등)
   - 제외 키워드 (장비·공사·행사·해커톤 등이 들어가면 제외)
   - 예산 구간, 입찰 마감까지 남은 일수, 참여지역 제한
5. 결과를 터미널에 출력하고, 옵션에 따라 Markdown 파일로 저장하거나 이메일로 발송합니다.

## 로컬 실행

```bash
pip install -r requirements.txt

# Gmail에서 오늘 받은 메일 분석
export GMAIL_ADDRESS=you@hnine.com
export GMAIL_APP_PASSWORD=xxxxxxxxxxxxxxxx   # Google 계정 > 보안 > 앱 비밀번호
python -m kjebi_recommender

# 최근 3일치를 분석해서 메일로 받기 (받는 사람 기본값: 본인)
REPORT_TO=team@hnine.com python -m kjebi_recommender --days 3 --send --out reports

# 메일을 .eml로 저장해 두었다면 IMAP 없이 실행
python -m kjebi_recommender --input ~/Downloads/kjebi.eml
```

> Google Workspace 관리자가 IMAP이나 앱 비밀번호를 막아 둔 경우에는 관리자에게 허용을 요청해야 합니다.

## 매일 자동 실행 (GitHub Actions)

`.github/workflows/kjebi-daily.yml`이 **매일 09:30 KST**에 실행되어 리포트 메일을 보냅니다.

저장소 **Settings → Secrets and variables → Actions**에 다음 값을 등록하세요.

| Secret | 설명 |
|---|---|
| `GMAIL_ADDRESS` | 크제비 메일을 받는 Gmail 주소 |
| `GMAIL_APP_PASSWORD` | 해당 계정의 앱 비밀번호 (16자리) |
| `REPORT_TO` | (선택) 리포트 받을 주소, 쉼표로 여러 명 가능 |

Actions 탭에서 `Run workflow`로 바로 실행해 볼 수도 있습니다. 실행 요약 화면과 아티팩트에도 리포트가 남습니다.

## 추천 기준 바꾸기

`config.yaml`만 수정하면 됩니다.

- `positive` / `negative`: 제목 키워드별 가중치
- `exclude`: 포함되면 무조건 제외할 키워드
- `budget`: 최소 예산, 선호 예산 구간
- `deadline.min_days_left`: 입찰 준비에 필요한 최소 일수
- `regions_ok`: 참여 가능한 지역
- `thresholds`: 강력/검토 추천 기준 점수

## 테스트

```bash
pip install pytest
python -m pytest
```
