from playwright.sync_api import sync_playwright

# Your live GitHub Pages dashboard URL
URL = "https://suwicha-cst.github.io/F-B-Dashboard/"

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={"width": 1400, "height": 900})
    page.goto(URL, wait_until="networkidle")
    page.wait_for_timeout(3000)  # give charts a moment to finish rendering
    page.screenshot(path="dashboard.png", full_page=True)
    browser.close()

print("Screenshot saved as dashboard.png")
