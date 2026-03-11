"""
자금집행일정 캘린더에서 출금일/작성마감일 파싱

캘린더 로직:
- "X일출금 지결작성마감"이 적힌 날짜 = 작성 데드라인
- X일 = 실제 출금일
- 월초 상신 항목은 해당 월 첫 번째 출금일을 지급요청일로 사용
"""
import re
from datetime import datetime, date
from typing import Optional


def parse_calendar_data(calendar_values: list[list[str]], target_year: int, target_month: int) -> list[dict]:
    """
    캘린더 시트 데이터에서 특정 월의 출금 스케줄을 파싱

    Returns:
        [{"deadline": date, "payment_date": date, "label": str}, ...]
        deadline 기준 오름차순 정렬
    """
    schedules = []
    month_str = f"{target_month}/"

    for row_idx, row in enumerate(calendar_values):
        for col_idx, cell in enumerate(row):
            if not cell:
                continue

            # "X일출금 지결작성마감" 패턴 찾기
            match = re.search(r'(\d+)일출금\s*\n?지결작성마감', cell)
            if not match:
                # "X/Y 출금 지결작성마감" 패턴도 체크 (월 넘어가는 경우)
                match = re.search(r'(\d+)/(\d+)\s*출금\s*\n?지결작성마감', cell)
                if match:
                    pay_month = int(match.group(1))
                    pay_day = int(match.group(2))
                else:
                    continue
            else:
                pay_day = int(match.group(1))
                pay_month = target_month

            # 이 셀의 날짜 찾기 (같은 열의 위쪽 행에서 날짜 헤더 탐색)
            deadline_date = _find_cell_date(calendar_values, row_idx, col_idx, target_year, target_month)
            if not deadline_date:
                continue

            # 출금일 계산
            if pay_month != target_month:
                # 다음달 출금
                if pay_month > 12:
                    payment_date = date(target_year + 1, pay_month - 12, pay_day)
                else:
                    next_month = target_month + 1
                    next_year = target_year
                    if next_month > 12:
                        next_month = 1
                        next_year += 1
                    payment_date = date(next_year, pay_month, pay_day)
            else:
                payment_date = date(target_year, target_month, pay_day)

            schedules.append({
                "deadline": deadline_date,
                "payment_date": payment_date,
                "label": cell.strip(),
            })

    # deadline 기준 정렬
    schedules.sort(key=lambda x: x["deadline"])
    return schedules


def _find_cell_date(
    values: list[list[str]], row_idx: int, col_idx: int,
    target_year: int, target_month: int
) -> Optional[date]:
    """셀 위치에서 해당 날짜를 역추적"""
    # 위쪽으로 올라가면서 날짜 패턴 (M/D) 찾기
    for r in range(row_idx, max(row_idx - 4, -1), -1):
        if r >= len(values) or col_idx >= len(values[r]):
            continue
        cell = values[r][col_idx]
        if not cell:
            continue
        # "3/5" 같은 패턴
        date_match = re.match(r'^(\d{1,2})/(\d{1,2})$', cell.strip())
        if date_match:
            m = int(date_match.group(1))
            d = int(date_match.group(2))
            if m == target_month:
                try:
                    return date(target_year, m, d)
                except ValueError:
                    continue
    return None


def get_first_payment_date(schedules: list[dict]) -> Optional[dict]:
    """월 첫 번째 출금 스케줄 반환"""
    if not schedules:
        return None
    return schedules[0]


def get_payment_for_deadline(schedules: list[dict], submit_date: date) -> Optional[dict]:
    """특정 작성일 기준으로 해당하는 출금 스케줄 반환"""
    for schedule in schedules:
        if submit_date <= schedule["deadline"]:
            return schedule
    return None
