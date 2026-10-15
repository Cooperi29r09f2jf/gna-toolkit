"""Drive the task pane in headless Chromium against tests/mock-office.js.

    python3 tests/ui_test.py [screenshot_dir]
"""
import functools
import http.server
import json
import socketserver
import sys
import threading
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "tests" / "screens"
OUT.mkdir(parents=True, exist_ok=True)
MOCK = (ROOT / "tests" / "mock-office.js").read_text()


class Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a):
        pass


def serve():
    httpd = socketserver.TCPServer(("127.0.0.1", 0), functools.partial(Quiet, directory=str(ROOT)))
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd


def calls(page, kind=None):
    c = page.evaluate("window.__calls")
    return [x for x in c if kind is None or x[0] == kind]


def deck(page):
    return page.evaluate("window.__deck")


def main():
    httpd = serve()
    base = f"http://127.0.0.1:{httpd.server_address[1]}/"
    errors = []
    with sync_playwright() as p:
        b = p.chromium.launch()
        page = b.new_page(viewport={"width": 340, "height": 900})
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.route("**/appsforoffice.microsoft.com/**", lambda r: r.fulfill(status=200, content_type="text/javascript", body=MOCK))
        page.route("**/fonts.googleapis.com/**", lambda r: r.fulfill(status=200, content_type="text/css", body=""))

        # ---- Library
        page.goto(base + "taskpane.html?tab=library")
        page.wait_for_selector(".card")
        n_cards = page.locator("#lib-groups .card").count()
        assert n_cards == 87, n_cards
        page.wait_for_timeout(400)
        page.screenshot(path=str(OUT / "1-library.png"))
        page.fill("#lib-search", "waterfall")
        assert page.locator("#lib-groups .card").count() == 2
        page.fill("#lib-search", "")
        page.click('#lib-groups .card[data-n="16"]')
        page.wait_for_function("window.__calls.some(c => c[0] === 'insertSlidesFromBase64')")
        ins = calls(page, "insertSlidesFromBase64")[0]
        assert ins[2]["targetSlideId"] == "257#", ins
        assert ins[2]["formatting"] == "UseDestinationTheme", ins
        assert ins[1] > 1000
        page.wait_for_selector("#toast:not([hidden])")
        assert "Inserted" in page.inner_text("#toast"), page.inner_text("#toast")
        assert page.locator("#lib-recent .card").count() == 1
        page.screenshot(path=str(OUT / "2-library-inserted.png"))
        print("library: ok (87 cards, search, insert after selected slide, destination theme, recent)")

        # ---- Starters
        page.click('.tabs button[data-tab="starters"]')
        page.screenshot(path=str(OUT / "3-starters.png"), full_page=True)
        page.locator(".starter").first.locator("button.primary").click()
        page.wait_for_function("window.__calls.some(c => c[0] === 'createPresentation')")
        page.locator(".starter").nth(2).locator("button:not(.primary)").click()
        page.wait_for_function("window.__calls.filter(c => c[0] === 'insertSlidesFromBase64').length === 2")
        print("starters: ok (new deck + insert)")

        # ---- Tools (selected slide 257#)
        page.evaluate("window.__deck.selected = ['257#']")
        page.click('.tabs button[data-tab="tools"]')
        page.screenshot(path=str(OUT / "4-tools.png"), full_page=True)
        page.click('#sticker-row button:has-text("For discussion")')
        page.wait_for_timeout(150)
        page.click('#sticker-row button:has-text("Draft")')
        page.wait_for_timeout(150)
        s = [sh for sh in deck(page)["slides"] if sh["id"] == "257#"][0]
        stickers = [sh for sh in s["shapes"] if sh["name"] == "GNA sticker"]
        assert len(stickers) == 1 and stickers[0]["text"] == "Draft", stickers
        assert stickers[0]["font"]["italic"] is True and stickers[0]["font"]["size"] == 14
        page.click("#sticker-remove")
        page.wait_for_timeout(150)
        s = [sh for sh in deck(page)["slides"] if sh["id"] == "257#"][0]
        assert not [sh for sh in s["shapes"] if sh["name"] == "GNA sticker"]

        page.check('input[name="src"][value="gna"]')
        page.click("#src-apply")
        page.wait_for_timeout(150)
        s = [sh for sh in deck(page)["slides"] if sh["id"] == "257#"][0]
        srcs = [sh for sh in s["shapes"] if sh["text"].startswith("Source")]
        assert len(srcs) == 1 and srcs[0]["text"] == "Source: GNA analysis", srcs

        page.fill("#sec-text", "Situation")
        page.click("#sec-apply")
        page.wait_for_timeout(150)
        s = [sh for sh in deck(page)["slides"] if sh["id"] == "257#"][0]
        assert [sh for sh in s["shapes"] if sh["name"] == "Text Placeholder 8"][0]["text"] == "Situation"

        page.fill("#needed-text", "client revenue, FY2025")
        page.click("#needed-insert")
        page.wait_for_timeout(150)
        s = [sh for sh in deck(page)["slides"] if sh["id"] == "257#"][0]
        nf = [sh for sh in s["shapes"] if sh["name"] == "GNA needed flag"][0]
        assert nf["text"] == "[NEEDED: client revenue, FY2025]" and nf["fill"] == "#FFFF00", nf

        page.click('.dot-btn[data-dot="amber"]')
        page.wait_for_timeout(150)
        assert calls(page, "addGeometricShape")[-1][2] == "Ellipse"
        page.click('#harvey-row button[title="75%"]')
        page.wait_for_timeout(150)
        hv = calls(page, "setSelectedDataAsync")[-1]
        assert hv[1] == "xmlSvg" and hv[2].startswith("<svg"), hv
        print("tools: ok (sticker add/replace/remove, source, section tag, NEEDED, dot, Harvey)")

        # ---- Check: reset slide 257# source to an empty template one so the checker has work
        page.evaluate("""() => {
          const s = window.__deck.slides.find(x => x.id === '257#');
          s.shapes = s.shapes.filter(x => x.name !== 'GNA needed flag');
          s.shapes.find(x => x.text.startsWith('Source')).text = 'Source: xx';
        }""")
        page.click('.tabs button[data-tab="check"]')
        page.click("#check-run")
        page.wait_for_selector("#check-summary:not([hidden])")
        page.wait_for_timeout(200)
        page.screenshot(path=str(OUT / "5-check.png"), full_page=True)
        res = page.inner_text("#check-results")
        for expect in ["Em dash", "Font is Arial", "placeholder not filled", "Title Case", "full stop", "No source line", "Section tag"]:
            assert expect in res, (expect, res)
        assert "[NEEDED: Q3 close rate]" in page.inner_text("#needed-list")
        # "Add" on a slide whose source is the template's "Source: xx" must reuse that box
        page.locator('.slide-block', has_text="Slide 2").locator('.issue', has_text="No source line").locator("button").click()
        page.wait_for_timeout(300)
        s2 = [x for x in deck(page)["slides"] if x["id"] == "257#"][0]
        srcs = [sh for sh in s2["shapes"] if sh["text"].startswith("Source")]
        assert len(srcs) == 1 and srcs[0]["text"] == "Source: [NEEDED: source]", srcs
        fix_label = page.inner_text("#check-fix")
        page.click("#check-fix")
        page.wait_for_function("document.querySelector('#check-fix').hidden || !document.querySelector('#check-fix').textContent.startsWith('Fixing')")
        page.wait_for_timeout(400)
        d = deck(page)
        s2 = [x for x in d["slides"] if x["id"] == "257#"][0]
        body = [sh for sh in s2["shapes"] if sh["name"] == "TextBox 26"][0]
        assert body["text"] == "Dispatch is manual, crews wait for calls", body["text"]
        assert body["font"]["name"] == "Lora"
        title = [sh for sh in s2["shapes"] if sh["name"] == "Title 1"][0]
        assert not title["text"].endswith("."), title["text"]
        s3 = [x for x in d["slides"] if x["id"] == "258#"][0]
        table = [sh for sh in s3["shapes"] if sh["type"] == "Table"][0]
        assert table["table"][1][0] == "Close rate, Q3", table["table"]
        grp = [sh for sh in s3["shapes"] if sh["type"] == "Group"][0]
        assert grp["children"][0]["text"] == "Owner, Clay", grp["children"][0]["text"]
        page.screenshot(path=str(OUT / "6-check-after-fix.png"), full_page=True)
        res2 = page.inner_text("#check-results")
        assert "Em dash" not in res2 and "Font is Arial" not in res2
        print(f"check: ok ({fix_label} applied; dashes in shapes, groups and tables fixed; fonts to Lora; title period removed)")

        # ---- Guide
        page.click('.tabs button[data-tab="guide"]')
        page.wait_for_timeout(300)
        page.screenshot(path=str(OUT / "7-guide.png"))

        # ---- commands.html loads and registers without errors
        page2 = b.new_page()
        page2.on("pageerror", lambda e: errors.append("commands: " + str(e)))
        page2.route("**/appsforoffice.microsoft.com/**", lambda r: r.fulfill(status=200, content_type="text/javascript", body=MOCK))
        page2.goto(base + "commands.html")
        page2.wait_for_timeout(200)
        page2.evaluate("window.__deck.selected = ['258#']")
        page2.evaluate("new Promise(res => window.stickerDiscussion({ completed: res }))")
        d = page2.evaluate("window.__deck")
        assert [sh for sh in [x for x in d["slides"] if x["id"] == "258#"][0]["shapes"] if sh["name"] == "GNA sticker"]
        page2.evaluate("new Promise(res => window.sourceNeeded({ completed: res }))")
        print("commands: ok (ribbon functions run and complete)")
        b.close()
    httpd.shutdown()
    assert not errors, errors
    print("UI test passed. Screenshots in", OUT)


if __name__ == "__main__":
    main()
