#!/usr/bin/env python3
"""
data/benefit_tags.json (categories/rewardPrograms/cards, 이미 raw+혜택 병합됨)
  -> data/cards.json (id 슬러그 붙인 배열 형태, meta 포함)
  -> index.html 안의 <script id="card-data" type="application/json"> 블록 교체

사용법:
    python3 scripts/build.py
    (먼저 scripts/import_from_mapping.py로 data/benefit_tags.json을 최신화해둘 것)
"""
import json
import re
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TAGS_PATH = ROOT / "data" / "benefit_tags.json"
OUT_PATH = ROOT / "data" / "cards.json"
HTML_PATH = ROOT / "index.html"


def slugify(name, seen):
    base = re.sub(r"[^\w가-힣]+", "-", name.strip()).strip("-").lower()
    base = base or "card"
    slug = base
    n = 2
    while slug in seen:
        slug = f"{base}-{n}"
        n += 1
    seen.add(slug)
    return slug


def build():
    tags = json.loads(TAGS_PATH.read_text(encoding="utf-8"))

    seen_slugs = set()
    cards_out = []
    for name, c in tags["cards"].items():
        benefits = c.get("benefits", [])
        raw_fields = c.get("raw", {})
        cards_out.append({
            "id": slugify(name, seen_slugs),
            "name": name,
            "issuer": c.get("issuer"),
            "group": c.get("group"),
            "annualFee": c.get("annualFee"),
            "annualFeeRaw": c.get("annualFeeRaw"),
            "familyCardFee": c.get("familyCardFee"),
            "familyCardFeeRaw": c.get("familyCardFeeRaw"),
            "raw": raw_fields,
            "benefits": benefits,
            "hasStructuredBenefits": bool(benefits),
            "researched": bool(raw_fields) or bool(benefits),
        })

    final = {
        "meta": {
            "generatedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "sourceFile": "benefit_tags.json",
            "cardCount": len(cards_out),
            "taggedCardCount": sum(1 for c in cards_out if c["hasStructuredBenefits"]),
        },
        "categories": tags["categories"],
        "rewardPrograms": tags["rewardPrograms"],
        "cards": cards_out,
    }

    OUT_PATH.write_text(json.dumps(final, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: data/cards.json ({len(cards_out)}개 카드, 구조화됨 {final['meta']['taggedCardCount']}개)")
    return final


def inject_into_html(data):
    if not HTML_PATH.exists():
        print("index.html이 아직 없어서 데이터 주입은 건너뜁니다.")
        return
    html = HTML_PATH.read_text(encoding="utf-8")
    # </script>가 JSON 문자열 값 안에 등장해도 HTML 파서가 스크립트 블록을 조기 종료하지
    # 않도록 이스케이프한다 ("\/"는 JSON 표준 이스케이프라 JSON.parse가 그대로 복원함).
    payload = json.dumps(data, ensure_ascii=False).replace("</script", "<\\/script")
    pattern = re.compile(
        r'(<script id="card-data" type="application/json">)(.*?)(</script>)',
        re.DOTALL,
    )
    if not pattern.search(html):
        print("경고: index.html에서 id=card-data 스크립트 블록을 찾지 못했습니다.")
        return
    html = pattern.sub(lambda m: m.group(1) + payload + m.group(3), html, count=1)
    HTML_PATH.write_text(html, encoding="utf-8")
    print("OK: index.html에 최신 데이터 주입 완료")


if __name__ == "__main__":
    data = build()
    inject_into_html(data)
