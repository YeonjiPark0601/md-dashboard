"""
고정 지출결의 자동화 - 메인 실행기

사용법:
    python expense_runner.py --month 3 --item all
    python expense_runner.py --month 3 --item lotte_rental
    python expense_runner.py --month 3 --item dongbo
    python expense_runner.py --month 3 --item sindaerim --amounts "2:577914,3:32493"
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

import os
import asyncio
import argparse
import glob
from datetime import date, datetime
from pathlib import Path

import yaml

# 프로젝트 루트
PROJECT_ROOT = Path(__file__).parent.parent
CONFIG_DIR = PROJECT_ROOT / "config"
SCREENSHOT_DIR = PROJECT_ROOT / "screenshots"

from browser import DaouBrowser
from expense_form import ExpenseForm, format_amount, format_pay_date


def load_config():
    with open(CONFIG_DIR / "items.yaml", "r", encoding="utf-8") as f:
        items_config = yaml.safe_load(f)
    with open(CONFIG_DIR / "settings.yaml", "r", encoding="utf-8") as f:
        settings = yaml.safe_load(f)
    return items_config, settings


def get_item_config(items_config: dict, item_id: str) -> dict:
    for item in items_config["items"]:
        if item["id"] == item_id:
            return item
    raise ValueError(f"항목을 찾을 수 없습니다: {item_id}")


def find_attachments(item_config: dict, settings: dict, year: int, month: int) -> list[str]:
    """첨부파일 경로 찾기"""
    base_path = settings["attachment_base_path"]
    folder = item_config["attachment_folder"]
    full_path = os.path.join(base_path, folder)

    if not os.path.exists(full_path):
        print(f"  [경고] 첨부파일 폴더 없음: {full_path}")
        return []

    item_id = item_config["id"]
    ym = f"{year % 100:02d}{month:02d}"

    if item_id == "lotte_rental":
        folder_path = os.path.join(full_path, f"{year % 100}년 {month}월")
        if os.path.exists(folder_path):
            return sorted(glob.glob(os.path.join(folder_path, f"{ym}*")))
        return []
    elif item_id == "dongbo":
        return sorted(glob.glob(os.path.join(full_path, f"{ym}*동보빌딩*")))
    elif item_id == "sindaerim":
        return sorted(glob.glob(os.path.join(full_path, f"{ym}*신대림빌딩*")))
    else:
        return sorted(glob.glob(os.path.join(full_path, f"{ym}*")))


def make_title(item_config: dict, payment_date: date, total_amount: int) -> str:
    """제목 생성: [MD] YYYYMMDD_업체명_금액원"""
    date_str = payment_date.strftime("%Y%m%d")
    vendor = item_config["vendor"]
    amount_str = f"{total_amount:,}원"
    return f"[MD] {date_str}_{vendor}_{amount_str}"


async def process_item(
    item_id: str,
    year: int,
    month: int,
    variable_amounts: dict = None,
    action: str = "temp_save",
    payment_date_override: date = None,
):
    """단일 항목 지출결의서 처리"""
    items_config, settings = load_config()
    item_config = get_item_config(items_config, item_id)
    prev_month = month - 1 if month > 1 else 12

    print(f"\n{'='*50}")
    print(f"[처리] {item_config['name']} ({item_config['vendor']})")
    print(f"{'='*50}")

    # 1. 지급요청일 결정
    if payment_date_override:
        payment_date = payment_date_override
    elif item_config["payment_date_type"] == "fixed_day":
        pay_month = month + 1 if item_config.get("payment_next_month") else month
        pay_year = year
        if pay_month > 12:
            pay_month = 1
            pay_year += 1
        payment_date = date(pay_year, pay_month, item_config["payment_day"])
    else:
        # calendar_first: 외부에서 전달 필요
        payment_date = date(year, month, 10)  # 폴백
    print(f"  [지급요청일] {payment_date}")

    # 2. 첨부파일
    attachments = find_attachments(item_config, settings, year, month)
    print(f"  [첨부파일] {len(attachments)}개")
    for f in attachments:
        print(f"    - {os.path.basename(f)}")

    # 3. 행 데이터 준비
    sub_items = item_config["sub_items"]
    rows = []
    for sub in sub_items:
        desc = sub.get("description") or sub.get("description_template", "")
        desc = desc.replace("{prev_month}", str(prev_month))

        amount = sub.get("amount")
        if sub["type"] == "variable" and variable_amounts:
            amount = variable_amounts.get(sub["no"], amount)
        if amount is None:
            amount = 0

        bank_info = sub.get("bank_info") or item_config.get("bank_info")
        transfer = item_config.get("transfer_method", "계좌이체")

        rows.append({
            "vendor": item_config["vendor"],
            "description": desc,
            "amount": format_amount(amount),
            "pay_date": format_pay_date(payment_date),
            "bank": "자동이체" if transfer == "자동이체" else (bank_info.get("bank", "") if bank_info else ""),
            "account": "" if transfer == "자동이체" else (bank_info.get("account", "") if bank_info else ""),
            "holder": "" if transfer == "자동이체" else (bank_info.get("holder", "") if bank_info else ""),
        })

    # 4. 제목
    total_amount = sum(int(r["amount"].replace(",", "")) for r in rows)
    title = make_title(item_config, payment_date, total_amount)

    # 5. 브라우저 자동화
    browser = DaouBrowser(headless=False)
    try:
        await browser.start()
        await browser.login()
        await browser.goto_new_form()

        form = ExpenseForm(browser.page)

        # 제목
        await form.fill_title(title)

        # 행 추가 (1행은 기본 존재)
        if len(rows) > 1:
            await form.add_rows(len(rows) - 1)

        # 데이터 입력
        for i, row in enumerate(rows):
            await form.fill_row(i, **row)

        # 첨부파일
        await form.attach_files(attachments)

        # 참조자 추가
        ref_members = [m for m in settings.get("approval_line", []) if m.get("role") == "참조"]
        for ref in ref_members:
            await form.set_approval_ref(ref["name"])

        # 스크린샷
        screenshot_path = str(SCREENSHOT_DIR / f"{item_id}_{year}{month:02d}.png")
        await browser.screenshot(screenshot_path)

        # 임시저장 또는 결재요청
        if action == "submit":
            await form.submit()
        else:
            await form.temp_save()

        print(f"\n[성공] {item_config['name']} 지출결의서 {action} 완료!")
        return True

    except Exception as e:
        print(f"\n[실패] {item_config['name']}: {e}")
        try:
            await browser.screenshot(str(SCREENSHOT_DIR / f"error_{item_id}.png"))
        except:
            pass
        raise
    finally:
        await browser.close()


async def process_all(year: int, month: int, variable_amounts: dict = None, action: str = "temp_save"):
    """전체 항목 순차 처리"""
    items_config, _ = load_config()
    item_ids = [item["id"] for item in items_config["items"]]

    print(f"\n{'#'*50}")
    print(f"  {month}월 고정 지출결의 자동화 시작")
    print(f"  대상: {len(item_ids)}건 - {', '.join(item_ids)}")
    print(f"{'#'*50}")

    results = []
    for item_id in item_ids:
        try:
            amounts = variable_amounts.get(item_id, {}) if variable_amounts else None
            await process_item(item_id, year, month, amounts, action)
            results.append({"item": item_id, "status": "success"})
        except Exception as e:
            results.append({"item": item_id, "status": "failed", "error": str(e)})

    print(f"\n{'#'*50}")
    print(f"  처리 결과")
    print(f"{'#'*50}")
    for r in results:
        status = "OK" if r["status"] == "success" else f"FAIL: {r.get('error', '')}"
        print(f"  [{status}] {r['item']}")

    return results


def parse_amounts(amounts_str: str) -> dict:
    if not amounts_str:
        return {}
    result = {}
    for pair in amounts_str.split(","):
        no, amount = pair.split(":")
        result[int(no.strip())] = int(amount.strip())
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="고정 지출결의 자동화")
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument("--year", type=int, default=datetime.now().year)
    parser.add_argument("--item", type=str, default="all")
    parser.add_argument("--amounts", type=str, default=None)
    parser.add_argument("--action", type=str, default="temp_save", choices=["temp_save", "submit"])

    args = parser.parse_args()

    if args.item == "all":
        asyncio.run(process_all(args.year, args.month, action=args.action))
    else:
        amounts = parse_amounts(args.amounts) if args.amounts else None
        asyncio.run(process_item(args.item, args.year, args.month, amounts, args.action))
