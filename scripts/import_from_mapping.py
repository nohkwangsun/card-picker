#!/usr/bin/env python3
"""
카테고리·혜택매핑 xlsx (안내/카테고리/혜택매핑/혜택유형·조건/변경내역 5개 시트)
  + 카드 비교 xlsx (연회비 등 카드 기본정보)
  -> data/benefit_tags.json (계층형 카테고리 코드 기반, 새 스키마)

사용법:
    python3 scripts/import_from_mapping.py <카테고리매핑.xlsx> <카드비교.xlsx>

이 스크립트가 하는 일:
  1. '카테고리' 시트를 읽어 코드 계층 트리(categories[])를 만든다.
  2. '혜택유형·조건' 시트에서 리워드 프로그램 목록을 뽑는다.
  3. '혜택매핑' 시트의 각 행(카드 하나의 혜택 문장이 카테고리 코드 단위로
     쪼개진 것)을 파싱해서 카드별 benefits[]로 만든다. 혜택률/금액, 조건,
     한도 텍스트에서 숫자를 최대한 뽑아내고, 애매한 건 원문 그대로 note에 남긴다.
  4. xlsx_to_raw.extract()로 카드 비교 시트에서 연회비 등 기본정보를 가져와 병합한다.
  5. 알려진 카드명 오타(사용자 확인됨)를 정리한다.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlsx_to_raw import extract as extract_raw  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT_PATH = ROOT / "data" / "benefit_tags.json"

# 카드명 수동 보정 (현재는 없음 — 카드비교/혜택매핑 두 시트의 카드명이 그대로 일치함)
NAME_FIXES = {}

TYPE_MAP = {
    "청구할인": "discount", "결제일할인": "discount", "할인": "discount",
    "할인(정액)": "fixedDiscount", "청구할인(정액)": "fixedDiscount",
    "할인(리터당)": "perLiterDiscount",
    "캐시백": "discount", "캐시백(정액)": "perTxCashback",
    "포인트 적립": "accrual", "마일리지 적립": "mileage",
    "바우처": "voucher",
    "무료 이용": "freeUse",
    "수수료 면제": "feeDiscount", "수수료 할인": "feeDiscount",
    "우대 서비스": "serviceOnly",
}

REWARD_PROGRAM_BY_NOTE = {
    "대한항공": "skypass_korean", "스카이패스": "skypass_korean",
    "아시아나": "asiana_club",
    "M포인트": "hyundai_mpoint",
    "Membership Rewards": "amex_mr", "MR": "amex_mr",
    "마이신한": "shinhan_mypoint",
    "KB포인트리": "kb_pointrewards",
    "네이버페이": "naverpay_point",
}


def parse_num_won(text):
    """'150,000원' / '30,000 Mi' / '1.5만/2만원' 등에서 숫자(원 단위)를 뽑는다.
    여러 숫자가 있으면(예: 구간별) 가장 작은 값을 우선(보수적) 반환.
    'N원 ×M' 형태(정액 혜택을 연 M회 제공)는 N*M(연간 총액)로 계산한다."""
    if not text:
        return None
    text = str(text)
    m = re.search(r"([\d,.]+)\s*(만|천|억)?\s*원?\s*[×xX*]\s*(\d+)", text)
    if m:
        digits = m.group(1).replace(",", "")
        try:
            v = float(digits)
        except ValueError:
            v = None
        if v is not None:
            unit = m.group(2)
            if unit == "만":
                v *= 10000
            elif unit == "천":
                v *= 1000
            elif unit == "억":
                v *= 100000000
            return v * int(m.group(3))
    # '×2', 'x2'처럼 배수/횟수를 나타내는 숫자는 그 자체로 금액 후보가 아니므로 스캔에서 제외
    text_wo_multiplier = re.sub(r"[×xX*]\s*\d+", "", text)
    nums = []
    for m in re.finditer(r"([\d,.]+)\s*(만|천|억)?", text_wo_multiplier):
        digits = m.group(1).replace(",", "")
        if not digits or digits == ".":
            continue
        try:
            v = float(digits)
        except ValueError:
            continue
        unit = m.group(2)
        if unit == "만":
            v *= 10000
        elif unit == "천":
            v *= 1000
        elif unit == "억":
            v *= 100000000
        nums.append(v)
    if not nums:
        return None
    return min(nums)


def parse_rate_or_amount(rtype, value_text):
    """혜택률/금액 컬럼을 유형에 맞게 파싱. (rate, amount, rewardProgramHint, parseNote) 반환."""
    if not value_text:
        return None, None, None
    s = str(value_text).strip()

    if rtype in ("mileage", "accrual"):
        # 'X/1000 MPW' 형태 (원당 마일/포인트 비율)
        m = re.match(r"([\d.]+)\s*/\s*1000", s)
        if m:
            return float(m.group(1)) / 1000, None, None
        # '1.5%' 형태 (구간형은 첫 값)
        m = re.match(r"([\d.]+)\s*%", s)
        if m:
            return float(m.group(1)) / 100, None, None
        return None, None, s  # 파싱 실패 -> 원문만 note로

    if rtype in ("discount", "feeDiscount"):
        m = re.match(r"\+?\s*([\d.]+)\s*%", s)
        if m:
            return float(m.group(1)) / 100, None, None
        if rtype == "feeDiscount" and "면제" in s:
            # 구체적 수수료율 명시 없이 '면제'만 표기된 경우: 국내 카드사 평균 해외이용수수료
            # 약 1%를 절감하는 것으로 근사 (원문에 실제 수수료율 없음, 보수적 가정)
            return 0.01, None, "면제(구체적 수수료율 명시 없어 1% 근사)"
        amt = parse_num_won(s)
        if amt is not None:
            return None, amt, None
        return None, None, s

    if rtype == "perLiterDiscount":
        m = re.match(r"([\d.]+)\s*원\s*/\s*L", s)
        if m:
            return None, float(m.group(1)), None
        return None, None, s

    if rtype in ("fixedDiscount", "perTxCashback"):
        amt = parse_num_won(s)
        if amt is not None:
            return None, amt, None
        return None, None, s

    if rtype == "voucher":
        m = re.search(r"([\d,]+)\s*Mi", s)
        if m:
            return None, float(m.group(1).replace(",", "")), None
        amt = parse_num_won(s)
        if amt is not None:
            return None, amt, None
        return None, None, s

    if rtype == "freeUse":
        if "무제한" in s:
            return None, None, None  # unlimited은 visitsPerYear로 별도 처리
        return None, None, s

    return None, None, s


def parse_visits_per_year(value_text):
    if not value_text:
        return None, False
    s = str(value_text)
    if "무제한" in s:
        return None, True
    m = re.search(r"연\s*(\d+)\s*회", s)
    if m:
        return int(m.group(1)), False
    m = re.search(r"(\d+)\s*회", s)
    if m:
        return int(m.group(1)), False
    return None, False


def parse_min_card_spend(condition_text):
    """'전월 30만/50만/100만' -> 가장 낮은 구간(30만원)을 minCardSpend로."""
    if not condition_text:
        return 0
    s = str(condition_text)
    if "무실적" in s:
        return 0
    m = re.search(r"전월\s*([\d.,만/]+)", s)
    if not m:
        return 0
    first = re.split(r"[/]", m.group(1))[0]
    amt = parse_num_won(first)
    return int(amt) if amt else 0


def parse_cap(limit_text):
    """한도 텍스트에서 (대표 cap, 통합여부, capGroupKey원본) 추출.
    구간형(3천/7천/1만)은 가장 작은 값을 기본 cap으로 사용(보수적)."""
    if not limit_text:
        return None, False
    s = str(limit_text)
    is_shared = "통합" in s
    amt = parse_num_won(s)
    cap = int(amt) if amt else None
    return cap, is_shared


def parse_monthly_count_cap(condition_text, per_unit_amount):
    """'일1회·월10회'처럼 건당 정액 혜택(perTxCashback)의 월 최대 횟수만 조건에
    적혀있고 한도(원) 컬럼이 비어있는 경우, 월 횟수 x 건당 금액으로 월 한도를 근사."""
    if not condition_text or per_unit_amount is None:
        return None
    m = re.search(r"월\s*(\d+)\s*회", str(condition_text))
    if not m:
        return None
    return int(per_unit_amount) * int(m.group(1))


def load_categories(wb):
    ws = wb["카테고리"]
    categories = []
    for r in range(2, ws.max_row + 1):
        code = ws.cell(row=r, column=1).value
        if not code:
            continue
        code = str(code).strip()
        group = ws.cell(row=r, column=2).value
        level = ws.cell(row=r, column=6).value
        kind = ws.cell(row=r, column=7).value
        path = ws.cell(row=r, column=8).value
        keywords = ws.cell(row=r, column=9).value
        label = path.split(">")[-1].strip() if path else code
        parent_code = code.rsplit("-", 1)[0] if "-" in code else None
        categories.append({
            "code": code,
            "group": group,
            "label": label,
            "path": path,
            "level": level,
            "kind": kind,
            "parentCode": parent_code,
            "keywords": keywords,
        })
    return categories


def load_reward_programs(wb):
    ws = wb["혜택유형·조건"]
    programs = []
    seen = set()
    for r in range(2, ws.max_row + 1):
        dim = ws.cell(row=r, column=1).value
        val = ws.cell(row=r, column=2).value
        if dim != "리워드 프로그램" or not val:
            continue
        pid = re.sub(r"[^\w가-힣]+", "_", val.strip()).strip("_")
        if pid in seen:
            continue
        seen.add(pid)
        is_point = "포인트" in val
        programs.append({
            "id": pid,
            "label": val,
            "defaultRate": 1 if is_point else 20,
            "unit": "원/포인트" if is_point else "원/마일",
        })
    return programs


def guess_reward_program(note_text):
    if not note_text:
        return None
    for key, pid in REWARD_PROGRAM_BY_NOTE.items():
        if key in note_text:
            return pid
    return None


def load_benefit_rows(wb):
    ws = wb["혜택매핑"]
    by_card = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name:
            continue
        name = NAME_FIXES.get(name.strip(), name.strip())
        raw_type = ws.cell(row=r, column=3).value
        code = str(ws.cell(row=r, column=4).value).strip()
        note = ws.cell(row=r, column=7).value
        value_text = ws.cell(row=r, column=8).value
        condition = ws.cell(row=r, column=9).value
        limit_text = ws.cell(row=r, column=10).value
        select_group = ws.cell(row=r, column=11).value

        rtype = TYPE_MAP.get(raw_type, "discount")
        # 혜택률/금액 컬럼이 비어있으면(예: '택1' 항목의 금액이 note에만 서술된 경우)
        # note 텍스트에서라도 숫자를 뽑아본다 (여러 옵션 중 최솟값 = 보수적 근사).
        rate, amount, parse_fail_note = parse_rate_or_amount(rtype, value_text or note)
        min_card_spend = parse_min_card_spend(condition)
        cap, cap_shared = parse_cap(limit_text)
        if cap is None and rtype == "perTxCashback":
            # 건당 정액 캐시백인데 한도(원)는 안 적혀있고 조건에 '월N회'만 있는 경우,
            # 월 횟수 x 건당 금액으로 월 한도를 근사한다.
            derived_cap = parse_monthly_count_cap(condition, amount)
            if derived_cap is not None:
                cap = derived_cap

        benefit = {
            "code": code,
            "type": rtype,
            "note": note,
            "conditionRaw": condition,
            "limitRaw": limit_text,
            "minCardSpend": min_card_spend,
        }
        if rate is not None:
            benefit["rate"] = rate
        if amount is not None:
            benefit["amount"] = amount
        if cap is not None:
            benefit["cap"] = cap
        if cap_shared:
            benefit["capGroup"] = f"{name}::{limit_text}"
        if select_group:
            benefit["selectGroup"] = select_group
        if rtype in ("mileage", "accrual"):
            prog = guess_reward_program(note)
            if prog:
                benefit["rewardProgram"] = prog

        freeuse_parsed = False
        if rtype == "freeUse":
            visits, unlimited = parse_visits_per_year(value_text)
            if unlimited:
                benefit["unlimited"] = True
                freeuse_parsed = True
            elif visits is not None:
                benefit["visitsPerYear"] = visits
                freeuse_parsed = True

        if parse_fail_note:
            if rtype == "serviceOnly":
                # serviceOnly(카드 브랜드/네트워크 자체 서비스 등)는 원래 금액/횟수로
                # 환산되지 않는 서술형 정보이므로 파싱 실패로 취급하지 않는다
                pass
            elif rtype == "freeUse":
                # freeUse는 금전 rate/amount가 원래 없는 유형 -> visitsPerYear/unlimited로
                # 이용 횟수가 잡혔으면 정상 처리된 것이지 파싱 실패가 아님
                if not freeuse_parsed:
                    benefit["unparsedValue"] = parse_fail_note
            elif rate is None and amount is None:
                # 값 자체를 못 뽑아냄 -> 계산에서 제외해야 함 (엔진이 이 필드로 판단)
                benefit["unparsedValue"] = parse_fail_note
            else:
                # rate/amount는 이미 채워졌고, 근사/가정에 대한 설명만 남기는 것
                # (예: '면제'를 1%로 근사) -> 계산은 정상 진행되어야 하므로 별도 필드로 분리
                benefit["approxNote"] = parse_fail_note

        by_card.setdefault(name, []).append(benefit)

    _auto_share_duplicate_caps(by_card)
    return by_card


def _auto_share_duplicate_caps(by_card):
    """같은 카드 안에서 note/조건/한도가 완전히 동일한 혜택 행이 여러 카테고리 코드에
    걸쳐 반복되면(엑셀에서 한 문장을 카테고리별로 쪼개 여러 행으로 매핑한 경우),
    사실은 하나의 한도를 여러 카테고리가 나눠 쓰는 것이므로 capGroup을 자동으로
    공유시킨다 (limitRaw에 '통합'이라고 명시되지 않았어도 적용)."""
    for name, benefits in by_card.items():
        buckets = {}
        for b in benefits:
            if b.get("cap") is None or b.get("capGroup"):
                continue
            key = (b["type"], b.get("note"), b.get("conditionRaw"), b.get("limitRaw"))
            buckets.setdefault(key, []).append(b)
        for key, group in buckets.items():
            if len(group) < 2:
                continue
            cap_group = f"{name}::{key[1]}::{key[3]}"
            for b in group:
                b["capGroup"] = cap_group


def merge_raw(card_compare_path, benefits_by_card):
    raw = extract_raw(card_compare_path)
    cards_out = {}
    unmatched = []
    for c in raw["cards"]:
        name = NAME_FIXES.get(c["name"], c["name"])
        fields = dict(c.get("fields", {}))
        benefits = benefits_by_card.get(name, [])
        if not benefits:
            unmatched.append(name)
        cards_out[name] = {
            "issuer": fields.get("카드사"),
            "group": c.get("group"),
            "annualFee": fields.get("연회비") if isinstance(fields.get("연회비"), (int, float)) else None,
            "annualFeeRaw": fields.get("연회비"),
            "familyCardFee": fields.get("가족카드") if isinstance(fields.get("가족카드"), (int, float)) else None,
            "familyCardFeeRaw": fields.get("가족카드"),
            "raw": {k: v for k, v in fields.items() if k not in ("카드사", "연회비", "가족카드")},
            "benefits": benefits,
        }
    extra_benefit_cards = set(benefits_by_card) - set(cards_out)
    return cards_out, unmatched, extra_benefit_cards, raw


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    mapping_path, compare_path = sys.argv[1], sys.argv[2]

    wb = openpyxl.load_workbook(mapping_path, data_only=True)
    categories = load_categories(wb)
    reward_programs = load_reward_programs(wb)
    benefits_by_card = load_benefit_rows(wb)

    cards_out, unmatched, extra_benefit_cards, raw = merge_raw(compare_path, benefits_by_card)

    result = {
        "_readme": (
            "계층형 카테고리 코드(예: 17-01-01 = 주유·충전>주유소>GS칼텍스) 기반 스키마. "
            "scripts/import_from_mapping.py가 카테고리+혜택매핑 xlsx와 카드 비교 xlsx를 "
            "병합해서 생성함 (손으로 고칠 때는 재실행하면 덮어써지니 원본 xlsx를 고치는 게 우선)."
        ),
        "categories": categories,
        "rewardPrograms": reward_programs,
        "cards": cards_out,
    }
    OUT_PATH.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: 카테고리 {len(categories)}개, 리워드프로그램 {len(reward_programs)}개, "
          f"카드 {len(cards_out)}개 -> {OUT_PATH}")
    if unmatched:
        print(f"혜택 매핑 없는 카드({len(unmatched)}개, 미구조화 상태로 남음):", unmatched)
    if extra_benefit_cards:
        print(f"경고: 혜택매핑엔 있는데 카드비교 시트엔 없는 카드명:", extra_benefit_cards)


if __name__ == "__main__":
    main()
