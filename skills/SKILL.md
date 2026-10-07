---
name: screen-spec-pptx
description: 기능정의서·요구사항정의서로 모바일/관리자 화면설계서 PPTX를 만들거나 고칠 때 사용. 전체화면(전 장표)·해당화면(화면목록·간지·와이어프레임만) 범위 지원. 작성규칙 v0.45 기준, JSON 명세만 쓰고 고정 스크립트로 생성·검증한다.
---

# 화면설계서 PPTX 생성 (작성규칙 v0.45)

레이아웃·좌표·색·마커·Description 표·분할·페이지 번호·검증은 **모두 `build_deck.py`에 고정**되어 있다.
작성자는 **`screen_spec.json`만 쓴다.** PPTX 코드를 직접 짜지 않는다. 작성규칙 MD 전문과 pptx 스킬 문서도 읽지 않는다(이 문서가 요약본).

## 0. 준비 (매번 1회, 토큰 거의 0)

```bash
D=/home/claude/screen-spec-pptx
[ -d $D ] || git clone --depth 1 https://github.com/sxaimmed/screen-spec-pptx $D
S=$D/skills                       # scripts/build_deck.py, examples/*.json
ls $S/scripts/build_deck.py
```
- clone이 거부되면 add_repo(owner sxaimmed, repo screen-spec-pptx, access read) 후 다시 clone.
- 폴더가 없으면 사용자에게 `screen-spec-pptx.zip`을 첨부해 달라고 하고 압축을 풀어 쓴다.
- 필요: python-pptx, lxml (없으면 `pip install python-pptx --break-system-packages`), 렌더는 soffice·pdftoppm.
- 스키마가 헷갈릴 때만 `$S/examples/admin_sample.json`(관리자) / `mobile_sample.json`(모바일)을 연다. 스크립트 본문은 열지 않는다.

## 1. 작성 범위 — 전체화면 / 해당화면

요청 문구로 범위를 정한다. 범위에 따라 **들어가는 장표가 다르다.**

| 범위 | 요청 표현 | 포함 장표 | 제외 장표 |
|---|---|---|---|
| **전체화면** (`full`) | "전체화면", "전체" | 표지 · Revision History · 요구사항 및 질문답변 · 문서 개요·작성 기준 · 화면 목록(제외 화면 포함) · 플로우 · 간지 · 와이어프레임 상세 · 정책 · 선행 결정사항 — **모든 장표** | 없음 |
| **해당화면** (`partial`) | "해당화면", "일부화면", "이 화면만", "OO 화면" | **화면 목록 · 간지 · Description이 들어간 와이어프레임 상세만** | 표지, Revision History, 요구사항 및 질문답변, 문서 개요·작성 기준, 플로우, 제외 화면, 정책, 선행 결정사항 |

- **전체 장표는 "전체화면"이라고 할 때만** 만든다. **축소 구성은 "해당화면/일부화면"이라고 할 때만** 만든다.
- 둘 다 말하지 않았으면 생성 전에 한 번 묻는다(AskUserQuestion: 전체화면 / 해당화면). 물을 수 없는 상황(무인 실행)이면 전체화면으로 만들고 전달 메시지 첫 줄에 그렇게 판단했다고 적는다.
- 해당화면에서 **대상 화면**은 사용자가 지목한 화면(화면명·화면ID·메뉴)이다. 지목이 모호하면 대상 화면 목록(화면명·화면ID)을 먼저 제시해 확인받는다.
- 해당화면의 화면 목록·간지는 **대상 화면만** 담는다(대상이 없는 섹션은 간지째 빠짐). 장표 번호는 1부터 다시 매긴다.
- 해당화면이어도 Description 내용 규칙(§3)·참조ID 표기·마커 짝 맞춤은 전체화면과 똑같이 지킨다. 선행 결정사항 장표가 없으므로, 신규 참조ID(T-n 등)는 전달 메시지에 목록으로 알린다.
- 대상 화면에서 덱 밖 화면으로 가는 `{{p:KEY}}` 토큰은 "전체본 참조"로 자동 표기된다(경고 아님). 이동 설명의 **화면명 + [화면ID]**는 그대로 필수.

**스크립트 지정** — `doc.scope`(`"full"`|`"partial"`, 한글 `"전체화면"`|`"해당화면"`도 허용)와 `doc.targets`(화면 key 또는 screenId 배열), 또는 CLI `--scope`·`--targets`(CLI가 우선).
```bash
python3 $S/scripts/build_deck.py screen_spec.json -o out --scope partial --targets ISSUE_A,DocList
```
- 전체화면용 spec이 이미 있으면 **spec을 새로 쓰지 말고** `--scope partial --targets ...`로 같은 spec에서 뽑는다.
- spec이 없고 해당화면만 요청받았으면 대상 화면만 담은 spec을 쓴다. 이때 `revisions`·`qna`·`overview`·`excluded`·`flow`·`policies`는 생략해도 되고, `decisions`는 본문 참조ID 정합 점검용으로 채운다.
- targets에 없는 key/screenId는 경고로 나온다(오타 확인).

## 2. 절차

1. **입력 파싱** — 기능정의서(1차 원천) > 화면목록 > 요구사항정의서 > 기타 순으로 신뢰. 어긋나면 임의로 고르지 않고 검토사항(C-n/T-n)에 올린다. 첨부 없이 언급만 된 파일은 요청한다.
   - 요구사항정의서만 있으면: 화면목록 초안(화면명·가정 화면ID·유형·REQ-ID)을 표로 먼저 제시 → 승인 후 생성. 부재 시 가정 항목 전부 `[TBD]`.
2. **모드 판정** — C열 분류 Mobile web/App → `"MO"`, Admin/Desktop web → `"ADM"`. 섞이면 모드별 파일 분리(또는 screen에 `"mode"` 지정).
3. **spec 작성** — `/home/claude/work/screen_spec.json` (아래 §4 스키마). 큰 화면은 섹션 단위로 나눠 써도 된다.
4. **생성·검증**
   ```bash
   python3 $S/scripts/build_deck.py screen_spec.json -o out [--scope full|partial] [--targets K1,K2] [--xlsx]
   ```
   출력: PPTX(파일명 규칙 자동) + `report.json` + 콘솔 요약(범위·경고·검증 실패). `--xlsx`는 화면목록 xlsx 동시 생성.
5. **경고 0·검증 실패 0이 될 때까지 spec만 고쳐 재실행.** 경고 예: 마커 위치 없는 항목, 묶음 순서 위반, 760px 초과, 참조ID 불일치.
6. **육안 확인은 샘플만** — 새로 쓴 블록 유형이 들어간 상세 장표 1~2장만 `--render 12,15`로 PNG 확인(전 장표 렌더 금지).
7. **전달** — PPTX를 다운로드 전용 첨부로 보낸다. 메시지는 범위(전체화면/해당화면)·장표 수·화면 수·참조ID 건수·동기화 불일치·신규 T-n 건수만(해당화면은 신규 참조ID 목록 포함). spec JSON은 요청 시 함께.

**수정 요청**: spec의 해당 부분만 Edit로 고친 뒤 재생성한다(전체 재작성 금지). 확인이 필요한 장표만 `--render`.
**업데이트 모드**: 기존 PPTX는 python-pptx로 텍스트만 읽어 spec을 복원 → 변경 범위만 수정, 변경 요소에 `vmemo`·`mark`, Revision History 맨 위 행 추가. 버전은 승인 전 유지(날짜만 갱신). `--only 13,14`로 해당 장표만 뽑아 확인 가능.

## 3. 내용 규칙 (스크립트가 못 하는 것)

- **화면ID·Component ID·화면명은 기능정의서 값 그대로.** 없으면 가채번(`{SVC}-ADM-{메뉴}-{기능}`) + `[TBD]`.
- **Component ID 표시 — 산출물 형식별로 다르다.**
  - **PPTX 화면설계서**: Description 영역에 Component ID(예: `TDC-ADM-CARE-LIST-sel-4`)를 **표시하지 않는다.** 항목은 번호·`[유형] 항목명`·태그·설명 글머리만. 스크립트가 자동으로 뺀다.
  - **HTML 화면설계서**: Description 영역의 각 항목에 Component ID를 **반드시 표시한다**(항목명 바로 아래 줄).
  - 두 형식 모두 spec의 `cid`는 **항상 채운다**(동기화 키·HTML 표시용). `lines`·`overview` 문장 안에 Component ID를 직접 쓰지 않는다 — 쓰면 스크립트가 경고한다.
- 기능 요구사항 밖 화면·버튼·문구·필수(*)·기본값은 **추가 금지**(승인 전 `[TBD]` 또는 검토사항).
- Description 변환: Property의 `1. 2.` → 글머리 1개씩, `-` 하위 → 앞에 공백 2칸(하위 글머리). Property `None` → "동작 정의 없음 [TBD 참조ID]". 비고(P열) → 하위 글머리.
- 문장 규칙: 동작은 `[트리거] → (조건) → 결과`, 이동은 **화면명 + [화면ID]** 필수, 팝업은 버튼별 결과 모두, 문구는 큰따옴표 원문, DB 값 `%변수%`, 근거 요구사항 `(REQ-03)`, 원본 변경 `(v2.88 변경)`.
- 용어: 노출→표시/보이기·감추기, 모바일은 **터치**·관리자는 **클릭**, 날짜+시간=일시·날짜=일자, Confirm Y/N → `[예]` `[아니요]`(원문이 Y/N뿐이면 문구 [TBD]).
- **미확정 문장에는 참조ID 필수**: `[TBD C-2]` `[결정필요 C-3]` `[개발확인 C-14]`. C-n 기능정의서 검토사항, N-n 미해결 넘버링 점검, T-n 비고 [TBD] 중 번호 없는 것(신규 부여 → 전달 시 기능정의서 등재 대상으로 알림), Q-n 요구사항정의서 질의. 모든 참조ID는 `decisions`에 등재(페이지는 자동 수집).
- 번호: 0은 화면설명(자동), 1부터 위→아래·좌→우, 하위 `3-1`. 원본 번호가 틀렸으면 기능정의서 기준으로 재부여(동기화 키 = Component ID). 모드 변형 장표는 같은 컴포넌트에 같은 번호.
- 장표 단위: Screen/Page → 화면 1개. Popup·Alert은 부모 화면에 함께(ADM `popup`, MO frame `popup`) 또는 ID 따로 있고 항목 5개 이상이면 독립 화면. 한 Screen ID의 모드 변형(예: 관리자발급/접수발급)은 화면을 나누고 이름에 `(변형명)`, 화면ID는 같게. 여러 버튼이 공유하는 Confirm은 `common`에 한 장(screenId `-`).
- 공통 장표·GNB는 **입력 문서에 근거가 있을 때만**.
- 묶음 폼 요소 순서: **선택(셀렉트·기준 라디오) → 컨트롤러(기간·토글 버튼) → 폼요소(달력·인풋)**. Description 설명 순서도 동일.
- 버튼: 버튼 묶음마다 강조(`primary`) 1개(주 실행: 검색·저장·발급하기·확인·예). 선택된 기간/토글은 `on`.
- 파일명(자동): `{deliverable}_{service}_{version}_{YYYYMMDD}.pptx`. 짝 문서 버전이 있으면(예 v2.88) 그대로, 없으면 `v01`(문서 안 v0.1).

## 4. spec 스키마 (요약)

```jsonc
{
 "doc": {"project":"표지·헤더 프로젝트명","service":"파일명 서비스","deliverable":"화면설계서-관리자|화면설계서-모바일",
         "version":"v2.88","date":"YYYY-MM-DD","author":"-","mode":"ADM|MO","docId":"","source":"기준 기능정의서 파일명",
         "scope":"full|partial", "targets":["화면 key 또는 screenId"]},   // scope 생략=full, targets는 partial에서만 의미
 "revisions": [["v2.88","2026-09-29","전체","Notes","작성자"]],          // 최신이 맨 위
 "qna": [["Date/V","중요도","페이지","항목","Notes","요청자","의견","답변자"]], // 없으면 []
 "overview": {"purpose":"", "basis":[["기능정의서","파일명"]], "scope":["추가 줄"], "idRules":["..."]},
 "excluded": [["화면명","화면ID","제외 사유"]],
 "flow": {"nodes":[{"key":"LIST"},{"key":"X","label":"외부 화면","id":"-","grid":[2,0]}],
          "edges":[["LIST","ISSUE","F-1"]], "triggers":[["F-1","트리거","도착 [화면ID]"]], "notes":[]},   // 여러 개면 "flows":[...]
 "common": [screen...],                       // 공통 화면 (간지 자동)
 "sections": [{"title":"B2C > 서류신청관리","screens":[screen...]}],   // 섹션 간지 자동
 "policies": [{"title":"정책 · 상태값","tables":[{"caption":"","cols":["상태","정의"],"rows":[[..]],"widths":[2,10.53]}],"notes":[]}],
 "decisions": [["C-13","TBD|결정필요|개발확인","화면ID","내용","확인 주체"]]
}
```
본문 어디서나 `{{p:KEY}}`(첫 페이지)·`{{pr:KEY}}`(범위) 토큰 → 실제 페이지로 치환.

**screen**
```jsonc
{"key":"ISSUE_A","screenId":"DocIssue","name":"서류발급(관리자발급)","path":["B2C","서류신청관리","서류발급"],
 "type":"페이지|팝업|탭|바텀시트","status":"신규|수정|확인필요|정책변경|기존","req":"REQ-03","updated":"YYYY-MM-DD",
 "frameLabel":"관리자발급", "mode":"ADM|MO(생략 시 doc.mode)",
 "overview":["목적","진입 경로","사전조건"],
 "items":[{"no":"9","type":"선택","name":"발급방법","cid":"DocIssue-radio-2"/*PPTX 미표시·HTML 표시*/,"tag":true,"status":"신규",
           "lines":["단일 선택: e-mail / FAX","  하위 글머리(공백 2칸)","규칙 [TBD C-13]"]}],
 "wire": {...} }        // 760px 넘는 관리자 화면은 "wire":[{상단},{..., "suffix":" (하단)"}]
```
type 태그: 버튼 입력 선택 텍스트 영역 목록 표 탭 팝업 알럿 바텀시트 토스트 로직 공통. `tag:true` → 주황 마커. Description 분할·`(1/3)`·"p.N에서 이어짐"은 자동. **모든 item no는 wire 어딘가의 `no`와 짝**(아니면 경고).

**ADM wire** — `{"title","crumb"(생략 시 path), "blocks":[...], "popup":{...}}` 1280px 기준 px, 위에서 아래로 자동 적층.

| block | 형식 |
|---|---|
| `search` | `{"t":"search","rows":[row...]}` 열 [130,1150], 행 38 |
| `form` | `{"t":"form","rows":[row...]}` 열 [160 레이블,100 보조,1020 값], 행 34 (`cols`·`rh`로 변경) |
| row | `{"label":"신청자명*","no":"2","v":값,"aux":"보조열 텍스트","h":34,"vmemo":"v2.87 20260831 …","memo":"노란 메모","mark":"신규"}` / 여러 줄: `{"label","no","rows":[{"v","aux","h"}...]}`(레이블 세로 병합, 하위 행 32) / 자유: `{"cells":[{"label"},{"v","cs":2}]}` |
| 값 `v` | 문자열(셀 텍스트: 라디오 `●○`·체크 `☑□` 포함) 또는 배열 `["앞 텍스트", {컨트롤}..., "뒤 텍스트"]` — 간격·여백·겹침 방지 자동 |
| 컨트롤 | `{"sel":"신청일","w":110,"ph":true}` `{"inp":"placeholder","w":220,"filled":false}` `{"date":["2026.08.29","2026.09.29"]}` `{"btn":"1개월","on":true}` `{"btn":"검색","primary":true}` `{"btn":"x","disabled":true}` `{"area":"텍스트","w":760,"h":54}` — 공통 선택 키 `"no"` `"tbd":true`(주황 점선) `"mark":"수정"` |
| `buttons` | `{"t":"buttons","items":[{"btn":"초기화","no":"5"},{"btn":"검색","primary":true}],"align":"c","vmemo":""}` |
| `caption` | `{"t":"caption","text":"총 (4) 건","no":"7","right":[{"btn":"Excel 다운로드"},{"sel":"20/쪽"}]}` |
| `list` | `{"t":"list","cols":[["No",38],["발급번호",118]...],"rows":[[...]],"link":["발급번호"],"hl":["신규열"],"mk":{"8":"발급번호"},"rh":48}` — 셀 줄바꿈 `\n`, mk = 열 마커 |
| 기타 | `paging`(`text`,`right`) · `tabs`(`items`,`sel`) · `text` · `memo` · `gap`(`h`) · `omit`(~ 생략 ~) |
| `popup` | `{"w":560,"title":"알림","body":["문구",["라벨",{"inp":".."}]],"buttons":[{"btn":"아니요"},{"btn":"예","primary":true}],"no":"12"}` — 딤 자동 |

**MO wire** — `{"frames":[frame ≤3]}` 393×852, 왼쪽부터 배치.
frame: `{"label":"기본","title":"화면 타이틀","back":true,"nav":true,"right":"편집","navNo":"1","body":[item...],"bottom":[item...],"tabbar":{"items":[..],"sel":0,"no":""},"popup":{"title","body":[..],"buttons":[..],"no"},"sheet":{"title","h":360,"body":[..],"no"}}`
item(t): `title` `text`(size,bold,color,align) `input`/`select`(label,ph,val,req) `btn`(text,primary,disabled) `btns`(items) `radio`/`check`(text) `card`(lines) `list`(items) `kv`(rows) `image`(h) `tabs`(items,sel) `divider` `space`(h) `memo`(text) `toast`(text) — 공통 `no` `mark` `tbd` `after`(간격).

## 5. 전달 전 체크 (스크립트가 자동 검증하는 것 외)

- 범위가 요청과 맞는가: "전체화면"일 때만 전 장표, "해당화면/일부화면"일 때는 화면 목록·간지·와이어프레임 상세만(콘솔 `범위` 표시로 확인)

- 입력에 없는 요소를 넣지 않았는가, 이동 설명마다 [화면ID]가 있는가, `[TBD]`마다 참조ID가 있는가
- PPTX Description에 Component ID가 없는가 / HTML 버전을 함께 만들었다면 HTML Description에는 모든 항목에 Component ID가 있는가
- 동기화: 화면ID·화면명·Component ID·순서가 기능정의서와 같은가 (화면목록·플로우는 spec 하나에서 자동 생성되므로 spec만 맞으면 된다)
