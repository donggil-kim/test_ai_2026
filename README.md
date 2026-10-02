# 통합 AI 보안위협 매트릭스 (UTM)

## 산출물 (`output/`)

| 파일 | 설명 |
| --- | --- |
| `통합_AI보안위협_매트릭스_v3.2_LITE.xlsx` | v3.2 — 분류 체계 개선(Lv3 124개), 전체 Lv3를 참고 양식 문체(요약설명·참조: AI 관점·실제 사례)로 재작성 |
| `통합_AI보안위협_매트릭스_v3.1_LITE.xlsx` | v3.1 원본에서 핵심 시트·열만 남긴 경량본 |

## v3.2 원본 데이터 (`data/v3.2/`)

| 파일 | 내용 |
| --- | --- |
| `base.yaml` | Lv3 행(분류·위험 평가 입력·ATLAS/OWASP 매핑). v3.1에서 이관 후 분류 체계 변경 적용 |
| `lv2.yaml` · `domains.yaml` | Lv2 명칭·핵심 통제, Lv1 도메인 정의 |
| `cases.yaml` | ATLAS 사례 73건과 Lv3 매핑(발생가능성·근거 수준 산정에 사용) |
| `rewrite/*.yaml` | v3.2 문체 재작성본(Lv3 ID 키) — 도메인별로 추가 |
| `changelog.yaml` | 변경이력·Lv3 ID 대응표 |

## 빌드

```bash
python3 scripts/build_v3.2.py        # data/v3.2 → output/..._v3.2_LITE.xlsx
python3 scripts/build_lite_v3.1.py   # reference/utm/..._v3.1.xlsx → output/..._v3.1_LITE.xlsx
```

`scripts/migrate_v3.2.py`는 v3.1 → v3.2 분류 체계 이관(1회성) 스크립트입니다. 실행하면 `base.yaml`·`lv2.yaml`·`domains.yaml`·`cases.yaml`을 다시 생성하므로, 이후 분류 변경은 이 파일들을 직접 수정합니다(`rewrite/`·`changelog.yaml`은 영향 없음).

참조 자료는 [`reference/README.md`](reference/README.md)를 참고하세요.
