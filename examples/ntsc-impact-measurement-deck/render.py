#!/usr/bin/env python3
"""Render each .slide (1280x720) to a high-res PNG and build a PDF deck."""
import asyncio, os, glob
from playwright.async_api import async_playwright

HTML = "file://" + os.path.abspath("ntsc-impact-model.html")
OUT = "slides"; os.makedirs(OUT, exist_ok=True)
SCALE = 2  # retina

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path="/home/user/deck_redesign/chrome-linux64/chrome",
            args=["--no-sandbox","--disable-gpu","--force-color-profile=srgb"])
        page = await browser.new_page(viewport={"width":1280,"height":720},
                                      device_scale_factor=SCALE)
        await page.goto(HTML, wait_until="networkidle")
        # ensure webfonts ready
        await page.evaluate("document.fonts.ready")
        await page.wait_for_timeout(1200)
        slides = await page.query_selector_all(".slide")
        print(f"slides found: {len(slides)}")
        pngs=[]
        for i, s in enumerate(slides, 1):
            f=f"{OUT}/slide-{i:02d}.png"
            await s.screenshot(path=f)
            pngs.append(f); print("captured", f)
        await browser.close()
    # build PDF from PNGs
    from PIL import Image
    imgs=[Image.open(f).convert("RGB") for f in sorted(pngs)]
    imgs[0].save("NTSC-Impact-Measurement-Model.pdf", save_all=True,
                 append_images=imgs[1:], resolution=150.0)
    print("PDF written: NTSC-Impact-Measurement-Model.pdf")

asyncio.run(main())
