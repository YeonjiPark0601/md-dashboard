"""
다우오피스 지출결의서 웹폼 자동 입력

검증된 셀렉터 (2026-03-11):
- 제목: input#subject
- 1행 필드: editorForm_5(업체명), 6(내역), 7(금액), 8(날짜), 9(은행), 10(계좌), 11(예금주)
- 추가행: dynamic_table1_{행번호}_{1~7}
- 행 추가 버튼: a#plus1
- 첨부파일: input[type="file"]
- 임시저장: a.btn_tool:has-text("임시저장")
- 결재요청: a.btn_tool:has-text("결재요청")
- 결재라인: 기본 설정됨 (박연지 → 전진영 → 조대현)
- 참고사항 에디터: dext_frame_editorForm_13
"""
import os
from datetime import date
from playwright.async_api import Page


# 1행(기본행) 필드 ID
ROW1_FIELDS = ['editorForm_5', 'editorForm_6', 'editorForm_7', 'editorForm_8',
               'editorForm_9', 'editorForm_10', 'editorForm_11']


def get_row_field_ids(row_idx: int) -> list[str]:
    """행 인덱스(0-based)에 따른 필드 ID 목록 반환"""
    if row_idx == 0:
        return ROW1_FIELDS
    else:
        return [f'dynamic_table1_{row_idx}_{i}' for i in range(1, 8)]


class ExpenseForm:
    def __init__(self, page: Page):
        self.page = page
        # confirm 다이얼로그 자동 수락
        self.page.on("dialog", lambda dialog: dialog.accept())

    async def _set_field_js(self, field_id: str, value: str):
        """JS로 필드 값 설정 (readonly/datepicker 대응)"""
        escaped = value.replace("\\", "\\\\").replace("'", "\\'")
        await self.page.evaluate(f"""
            (() => {{
                const el = document.getElementById('{field_id}');
                if (el) {{
                    el.readOnly = false;
                    el.disabled = false;
                    el.value = '{escaped}';
                    el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                }}
            }})()
        """)

    async def fill_title(self, title: str):
        """제목 입력"""
        await self.page.fill('input#subject', title)
        print(f"  [입력] 제목: {title}")

    async def add_rows(self, count: int):
        """행 추가 (count개)"""
        for _ in range(count):
            await self.page.click('a#plus1')
            await self.page.wait_for_timeout(500)
        print(f"  [행 추가] {count}개 추가")

    async def fill_row(self, row_idx: int, vendor: str, description: str,
                       amount: str, pay_date: str, bank: str = "",
                       account: str = "", holder: str = ""):
        """행 데이터 입력 (row_idx: 0-based)"""
        field_ids = get_row_field_ids(row_idx)
        values = [vendor, description, amount, pay_date, bank, account, holder]

        for field_id, value in zip(field_ids, values):
            await self._set_field_js(field_id, value)

        print(f"  [행 {row_idx + 1}] {vendor} / {description[:35]} / {amount}")

    async def fill_note(self, html_content: str):
        """참고사항 에디터에 HTML 삽입"""
        frame = self.page.frame(name="dext_frame_editorForm_13")
        if frame:
            escaped = html_content.replace('`', '\\`').replace('${', '\\${')
            await frame.evaluate(f"document.body.innerHTML = `{escaped}`;")
            print("  [입력] 참고사항 삽입")

    async def attach_files(self, file_paths: list[str]):
        """첨부파일 업로드"""
        if not file_paths:
            print("  [스킵] 첨부파일 없음")
            return

        for file_path in file_paths:
            if not os.path.exists(file_path):
                print(f"  [경고] 파일 없음: {file_path}")
                continue

            file_input = await self.page.query_selector('input[type="file"]')
            if file_input:
                await file_input.set_input_files(file_path)
                await self.page.wait_for_timeout(2000)
                print(f"  [첨부] {os.path.basename(file_path)}")

    async def temp_save(self):
        """임시저장"""
        save_btn = await self.page.query_selector('a.btn_tool:has-text("임시저장")')
        if save_btn:
            await save_btn.click()
            await self.page.wait_for_timeout(3000)
            print("  [완료] 임시저장")
            return True
        print("  [경고] 임시저장 버튼 못 찾음")
        return False

    async def submit(self):
        """결재요청 (상신)"""
        submit_btn = await self.page.query_selector('a.btn_tool:has-text("결재요청")')
        if submit_btn:
            await submit_btn.click()
            await self.page.wait_for_timeout(3000)
            print("  [완료] 결재요청")
            return True
        print("  [경고] 결재요청 버튼 못 찾음")
        return False

    async def set_approval_ref(self, name: str, member_id: str = ""):
        """결재 정보 팝업에서 참조자 추가 (드래그 앤 드롭)

        Args:
            name: 참조자 이름 (예: "이예림")
            member_id: 멤버 ID (예: "MEMBER_327"). 비어있으면 이름으로 검색.
        """
        page = self.page

        # 1) 결재 정보 팝업 열기
        await page.click('a.btn_tool:has-text("결재 정보")')
        await page.wait_for_timeout(3000)

        # 2) 참조자 탭 클릭
        await page.evaluate("""
            () => {
                const allLi = document.querySelectorAll('li');
                for (const li of allLi) {
                    if (li.textContent.trim().includes('참조자')) {
                        li.querySelector('a') ? li.querySelector('a').click() : li.click();
                        return;
                    }
                }
            }
        """)
        await page.wait_for_timeout(2000)

        # 3) 참조자 트리(aside-referer-orgtree-tree)에서 대상 찾기
        tree_id = 'aside-referer-orgtree-tree'

        # member_id가 없으면 이름으로 검색
        if not member_id:
            member_id = await page.evaluate(f"""
                () => {{
                    const tree = document.getElementById('{tree_id}');
                    if (!tree) return '';
                    const links = tree.querySelectorAll('a');
                    for (const a of links) {{
                        if (a.textContent.trim() === '{name}' && a.id.startsWith('MEMBER_')) {{
                            return a.id;
                        }}
                    }}
                    return '';
                }}
            """)

        if not member_id:
            print(f"  [경고] 참조자 '{name}' 을(를) 조직도에서 찾지 못함")
            return False

        # 4) 트리 스크롤 컨테이너에서 대상이 보이도록 스크롤
        await page.evaluate(f"""
            () => {{
                const tree = document.getElementById('{tree_id}');
                if (!tree) return;
                const el = tree.querySelector('a#{member_id}');
                if (!el) return;

                let container = tree.parentElement;
                while (container) {{
                    const cs = window.getComputedStyle(container);
                    if (cs.overflow === 'auto' || cs.overflow === 'scroll' ||
                        cs.overflowY === 'auto' || cs.overflowY === 'scroll') {{
                        break;
                    }}
                    container = container.parentElement;
                }}
                if (!container) container = tree.parentElement;

                const containerRect = container.getBoundingClientRect();
                const elRect = el.getBoundingClientRect();
                const scrollNeeded = elRect.top - containerRect.top - containerRect.height / 2;
                container.scrollTop += scrollNeeded;
            }}
        """)
        await page.wait_for_timeout(500)

        # 5) 드래그 좌표 계산
        coords = await page.evaluate(f"""
            () => {{
                const tree = document.getElementById('{tree_id}');
                const src = tree ? tree.querySelector('a#{member_id}') : null;
                if (!src) return null;
                const srcR = src.getBoundingClientRect();
                if (srcR.width === 0) return null;

                const droppables = document.querySelectorAll('tr.appr-activity.ui-droppable');
                for (const d of droppables) {{
                    if (d.offsetParent !== null) {{
                        const dR = d.getBoundingClientRect();
                        if (dR.width > 0) {{
                            return {{
                                sx: srcR.x + srcR.width / 2,
                                sy: srcR.y + srcR.height / 2,
                                tx: dR.x + dR.width / 2,
                                ty: dR.y + dR.height / 2
                            }};
                        }}
                    }}
                }}
                return null;
            }}
        """)

        if not coords:
            print(f"  [경고] 참조자 드래그 좌표를 계산할 수 없음")
            return False

        # 6) 마우스 드래그 (jQuery UI draggable → droppable)
        sx, sy, tx, ty = coords['sx'], coords['sy'], coords['tx'], coords['ty']
        await page.mouse.move(sx, sy)
        await page.wait_for_timeout(200)
        await page.mouse.down()
        await page.wait_for_timeout(500)

        steps = 30
        for i in range(steps + 1):
            t = i / steps
            await page.mouse.move(sx + (tx - sx) * t, sy + (ty - sy) * t)
            await page.wait_for_timeout(20)

        await page.wait_for_timeout(500)
        await page.mouse.up()
        await page.wait_for_timeout(2000)

        # 7) 추가 확인
        added = await page.evaluate(f"""
            () => {{
                const popup = document.getElementById('gpopupLayer');
                if (!popup) return false;
                return popup.textContent.includes('{name}');
            }}
        """)

        if not added:
            print(f"  [경고] 참조자 '{name}' 추가 실패")
            return False

        # 8) 확인 버튼 클릭
        await page.evaluate("""
            () => {
                const popup = document.getElementById('gpopupLayer');
                if (!popup) return;
                const btns = popup.querySelectorAll('a.btn_major_s');
                for (const btn of btns) {
                    if (btn.textContent.trim() === '확인' && btn.offsetParent !== null) {
                        btn.click();
                        return;
                    }
                }
            }
        """)
        await page.wait_for_timeout(2000)
        print(f"  [참조자] {name} 추가 완료")
        return True


def format_amount(amount: int) -> str:
    """금액을 콤마 포함 문자열로"""
    return f"{amount:,}"


def format_pay_date(d: date) -> str:
    """날짜를 YYYY-MM-DD 문자열로"""
    return d.strftime("%Y-%m-%d")
