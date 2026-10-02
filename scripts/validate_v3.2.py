"""v3.2 데이터 검증: 문체 작성 현황, 참조 ID 유효성, 사례 라벨-ATLAS 유형 일치, OWASP 인용 근거 수 정합성."""
import re
import sys
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data/v3.2"
base = yaml.safe_load((DATA / "base.yaml").read_text(encoding="utf-8"))
cases = {c["id"]: c for c in yaml.safe_load((DATA / "cases.yaml").read_text(encoding="utf-8"))}
rw = {}
for f in sorted((DATA / "rewrite").glob("*.yaml")):
    rw.update(yaml.safe_load(f.read_text(encoding="utf-8")) or {})
ids = {r["id"] for r in base}
problems, pending = [], []
for r in base:
    e = {**r, **rw.get(r["id"], {})}
    if r["id"] not in rw or rw[r["id"]].get("style") == "v3.1":
        pending.append(r["id"])
        continue
    summary, ref = e.get("summary") or "", e.get("reference") or ""
    if len([l for l in summary.split("\n") if l.startswith("- ")]) != 2:
        problems.append((r["id"], "요약설명 2줄 아님"))
    if "■ AI 관점" not in ref or "■ 실제 사례" not in ref:
        problems.append((r["id"], "참조 단락 누락"))
    text = summary + "\n" + ref
    refs = set(re.findall(r"UT-\d{2}\.\d", text))
    for m in re.finditer(r"UT-\d{2}\.\d((?:·\d{2}\.\d)+)", text):
        refs |= {"UT-" + x for x in re.findall(r"\d{2}\.\d", m.group(1))}
    problems += [(r["id"], "없는 Lv3 참조", x) for x in sorted(refs - ids)]
    owasp_inc = owasp_vul = 0
    for line in ref.split("\n"):
        m = re.match(r"- \[([^\]]+)\]", line)
        if not m:
            continue
        label, cs = m.group(1), re.search(r"AML\.CS\d{4}", line)
        if cs:
            if cs.group(0) not in cases:
                problems.append((r["id"], "없는 사례", cs.group(0)))
            elif (cases[cs.group(0)]["type"] == "Incident") != (label == "실제 사고"):
                problems.append((r["id"], "라벨-유형 불일치", cs.group(0), label))
        else:
            owasp_inc += label == "실제 사고"
            owasp_vul += label == "공개 취약점"
    if owasp_inc > e["owasp_incidents"]:
        problems.append((r["id"], f"OWASP 인용 실제 사고 {owasp_inc}건 > 집계 {e['owasp_incidents']}"))
    if owasp_vul > e["owasp_vulns"]:
        problems.append((r["id"], f"OWASP 인용 취약점 {owasp_vul}건 > 집계 {e['owasp_vulns']}"))
print(f"v3.2 문체 {len(base) - len(pending)}/{len(base)}", "미작성:", pending or "없음")
for p in problems:
    print("  !", *p)
sys.exit(1 if problems else 0)
