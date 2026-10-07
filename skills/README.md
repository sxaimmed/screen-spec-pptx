# screen-spec-pptx

화면설계서 PPTX 작성규칙 v0.45 생성기. `SKILL.md`가 Claude용 절차·스키마이고, `scripts/build_deck.py`가 JSON → PPTX 생성·검증기다.

```
python3 scripts/build_deck.py examples/admin_sample.json -o out --render 10 --xlsx
```
- `--only 13,14` 지정 장표만 저장(페이지 번호 유지)
- `--render all|13,14` PNG 렌더(soffice·pdftoppm 필요)
- `--xlsx` 화면목록 xlsx 동시 생성
- `--scope full|partial` 전체화면(전 장표, 기본) / 해당화면(화면목록·간지·와이어프레임만), `--targets K1,K2` 해당화면 대상 화면

규칙이 바뀌면(v0.46 등) 이 스크립트의 상수·함수와 SKILL.md를 같이 고친다.
