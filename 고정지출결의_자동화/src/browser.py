"""
다우오피스 브라우저 관리 + 로그인
"""
import os
import asyncio
from playwright.async_api import async_playwright, Browser, Page


class DaouBrowser:
    def __init__(self, headless: bool = False):
        self.headless = headless
        self.base_url = "https://gw.integrationcorp.co.kr"
        self.playwright = None
        self.browser: Browser = None
        self.page: Page = None

    async def start(self):
        """브라우저 시작"""
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(
            headless=self.headless,
            slow_mo=300,
        )
        self.page = await self.browser.new_page()
        self.page.set_default_timeout(30000)
        return self.page

    async def login(self):
        """다우오피스 로그인"""
        daou_id = os.environ.get("DAOU_ID")
        daou_pw = os.environ.get("DAOU_PW")

        if not daou_id or not daou_pw:
            raise ValueError(
                "환경변수 DAOU_ID, DAOU_PW가 설정되지 않았습니다.\n"
                "설정 방법: export DAOU_ID='아이디' && export DAOU_PW='비밀번호'"
            )

        await self.page.goto(f"{self.base_url}/login")
        await self.page.wait_for_load_state("networkidle")

        # 아이디/비밀번호 입력
        await self.page.fill('input#username', daou_id)
        await self.page.fill('input#password', daou_pw)

        # 로그인 버튼 클릭 (a 태그)
        await self.page.click('a#login_submit')
        await self.page.wait_for_timeout(5000)

        # 로그인 확인
        if "/login" in self.page.url:
            raise RuntimeError("로그인 실패 - 아이디/비밀번호를 확인해주세요")

        print("[OK] 다우오피스 로그인 성공")
        return True

    async def goto_new_form(self, dept_id: int = 50, form_id: int = 2219):
        """지출결의서 새 문서 작성 페이지로 이동"""
        url = f"{self.base_url}/app/approval/document/new/{dept_id}/{form_id}"
        await self.page.goto(url)
        await self.page.wait_for_timeout(8000)  # SPA 로딩 대기
        print(f"[OK] 지출결의서 작성 페이지 이동")
        return self.page

    async def screenshot(self, path: str):
        """스크린샷 저장"""
        await self.page.screenshot(path=path, full_page=True)
        print(f"[OK] 스크린샷 저장: {path}")

    async def close(self):
        """브라우저 종료"""
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()
        print("[OK] 브라우저 종료")
