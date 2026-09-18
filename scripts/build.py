#!/usr/bin/env python3
"""
data/cards_raw.json (xlsx 원문) + data/benefit_tags.json (수작업 구조화 오버레이)
  -> data/cards.json (최종 데이터, 카드 이름으로 매칭해서 병합)
  -> index.html 안에 <script id="card-data" type="application/json"> 블록을 최신 데이터로 교체

사용법:
    python3 scripts/build.py
    (먼저 scripts/xlsx_to_raw.py로 새 xlsx를 data/cards_raw.json 으로 변환해둘 것)
"""
import json
import re
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RAW_PATH = ROOT / "data" / "cards_raw.json"
TAGS_PATH = ROOT / "data" / "benefit_tags.json"
OUT_PATH = ROOT / "data" / "cards.json"
HTML_PATH = ROOT / "index.html"

TOP_LEVEL_FIELDS = {"카드사", "연회비", "가족카드"}


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


def to_number_or_none(v):
    if isinstance(v, (int, float)):
        return v
    return None  # "x" 등 텍스트는 숫자화하지 않음 (원문은 raw에 보존)


def build():
    raw = json.loads(RAW_PATH.read_text(encoding="utf-8"))
    tags = json.loads(TAGS_PATH.read_text(encoding="utf-8"))
    tag_cards = tags.get("cards", {})

    seen_slugs = set()
    cards_out = []
    for c in raw["cards"]:
        name = c["name"]
        fields = c.get("fields", {})
        raw_fields = {k: v for k, v in fields.items() if k not in TOP_LEVEL_FIELDS}
        tag_entry = tag_cards.get(name)
        benefits = tag_entry["benefits"] if tag_entry else []

        cards_out.append({
            "id": slugify(name, seen_slugs),
            "name": name,
            "issuer": fields.get("카드사"),
            "group": c.get("group"),
            "annualFee": to_number_or_none(fields.get("연회비")),
            "annualFeeRaw": fields.get("연회비"),
            "familyCardFee": to_number_or_none(fields.get("가족카드")),
            "familyCardFeeRaw": fields.get("가족카드"),
            "raw": raw_fields,
            "benefits": benefits,
            "hasStructuredBenefits": bool(benefits),
            "researched": bool(raw_fields) or bool(benefits),
        })

    final = {
        "meta": {
            "generatedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "sourceFile": raw.get("sourceFile"),
            "cardCount": len(cards_out),
            "taggedCardCount": sum(1 for c in cards_out if c["hasStructuredBenefits"]),
        },
        "spendCategories": tags["spendCategories"],
        "metaCategories": tags["metaCategories"],
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
    payload = json.dumps(data, ensure_ascii=False)
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
