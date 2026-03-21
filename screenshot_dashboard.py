import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        await page.goto('http://127.0.0.1:8050')
        await page.wait_for_timeout(5000)  # Wait for Dash loading
        await page.screenshot(path='dashboard_screenshot.png', full_page=True)
        await browser.close()
        print("Screenshot saved to dashboard_screenshot.png")

if __name__ == '__main__':
    asyncio.run(main())
