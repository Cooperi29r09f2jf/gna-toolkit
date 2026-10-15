# GNA Toolkit

A PowerPoint add-in for Gheen & Noyes Advisory consultants. It adds a **GNA Toolkit** tab next to Home with:

| Group | Button | What it does |
|---|---|---|
| Build | **Slide library** | 87 template slides in 9 categories, with thumbnails and search. Click a slide to insert it after the current one. Recently used slides stay at the top. |
| Build | **Deck starters** | Proposal, Findings, SteerCo and Workshop lineups from the deck rules. Open one as a new deck, or insert it into the deck you have open. |
| Quick tools | **Sticker** | For discussion, Preliminary or Draft on the selected slides, placed where the template puts it. Remove sticker too. |
| Quick tools | **Source line** | "Source: GNA analysis", "Source: [NEEDED: source]", or a client or named source from the Tools pane. Updates the existing line or adds one. |
| Quick tools | **[NEEDED] flag** | Yellow `[NEEDED: ...]` box on the current slide. |
| Quick tools | **More tools** | Section tag, status dots (green, amber, red, grey), Harvey balls (0 to 100), wrap selected text as `[NEEDED: ...]`. |
| Check | **Style checker** | Scans every slide for em dashes, non-Lora fonts, leftover `[placeholders]` and `xx` text, missing source lines, all-caps or Title Case titles, titles ending in a full stop, tiny text, empty section tags, style guide slides left in, and extra slide masters. One click fixes the safe ones (dashes to commas, fonts to Lora, title full stops, the Deliverables typo). Lists every open `[NEEDED]` flag with a Copy button. |
| Check | **Style guide** | The house rules and template style guide slides 1 to 3. |

Works in PowerPoint for Microsoft 365 on Windows, Mac and the web.

---

## One-time setup (about 30 minutes)

### 1. Put the files on GitHub Pages

The add-in is a small website that PowerPoint loads. GitHub Pages hosts it for free.

1. Sign in to GitHub as **Cooperi29r09f2jf**. If that is not your exact username, see "Changing the URL" below before continuing.
2. Create a new **public** repository named **`gna-toolkit`** (no README, no .gitignore).
3. Install **GitHub Desktop** (desktop.github.com), sign in, then **File > Clone repository** and pick `gna-toolkit`.
4. Copy everything in this folder (including the hidden `.nojekyll` file) into the cloned folder.
5. In GitHub Desktop: write a summary like "GNA Toolkit 1.0.0", click **Commit to main**, then **Push origin**.
6. On github.com, open the repo: **Settings > Pages**. Under "Build and deployment" choose **Deploy from a branch**, branch **main**, folder **/ (root)**, then **Save**.
7. Wait a minute or two, then open https://cooperi29r09f2jf.github.io/gna-toolkit/taskpane.html in a browser. You should see the toolkit page with slide thumbnails (buttons won't work outside PowerPoint, that's expected).

Note: Pages sites are public. Anyone with the exact URL could download the template slides. There is no client data in them.

### 2. Deploy to everyone at GNA (Microsoft 365 admin center)

1. Go to **admin.microsoft.com** and sign in as a Global admin.
2. Left menu: **Show all > Settings > Integrated apps**, then the **Add-ins** link near the top of the page.
3. **Deploy Add-in > Next**, and choose to upload a custom add-in.
4. Under **Choose how to upload app**, pick the manifest file option, **Choose File**, select `manifest.xml` from this folder, then **Upload**.
5. Assign users: **Just me** so you can test first. Then **Deploy**.
6. Restart PowerPoint. The **GNA Toolkit** tab appears next to Home. Microsoft says this can take 24 to 72 hours, though it's often faster.
7. Once you have tested it, open the add-in in the same page and change the assignment to **Everyone** (or a GNA consultants group, which Microsoft recommends so new hires get it automatically).

The admin center may offer a newer upload screen under **Integrated apps > Upload custom apps**; it accepts the same `manifest.xml`.

Everyone signed in to PowerPoint with an @gheennoyesadvisory.com work account then gets the tab automatically. Guests and personal Office accounts don't.

### 3. Test checklist (in real PowerPoint)

Open a blank deck and try, in order:

- [ ] Deck starters > Proposal > **Open as new deck**: a new window opens with 15 slides.
- [ ] In that deck, select slide 3 and insert *Waterfall chart* from the library: it lands as slide 4 and looks like the template.
- [ ] Settings at the bottom of the library: leave on **Auto**. If inserted slides ever look off, switch to **Keep template formatting** and tell Cooper.
- [ ] Sticker > For discussion, then Draft: one sticker, top right under the title.
- [ ] Source line > Source: GNA analysis on a slide that has "Source: xx": the line updates.
- [ ] Style checker on the starter deck: placeholder, `xx` and source issues show up; **Fix safe issues** runs; **Go to** jumps to the slide.
- [ ] Repeat a couple of steps on Mac and in PowerPoint on the web.

---

## Day-to-day changes

| You want to... | Do this |
|---|---|
| Change code or styling | Edit, run `python3 tools/bump_version.py 1.0.1`, commit and push. Users get it the next time they open the pane (PowerPoint may cache for a few hours; restarting PowerPoint helps). |
| Change ribbon buttons, names, icons or URLs | Edit `manifest.xml` (it's generated by hand; keep the structure), bump the version, push, then in the admin center's add-ins page select GNA Toolkit, upload the new manifest and click **Update**. Users get it the next time they start PowerPoint. |
| Update the slide template | `python3 tools/build_assets.py "path/to/THE Slide Template.pptx"` then commit and push. Needs Python 3. Thumbnails need LibreOffice; without it, add `--skip-thumbs` and the old thumbnails stay. If slide numbers change, update `tools/catalog.json` first. |
| Rename a slide, re-categorize it, change starter lineups | Edit `tools/catalog.json`, run `python3 tools/build_assets.py --catalog-only`. If you changed a starter lineup, run the full build so its .pptx is regenerated. |
| Move hosting (e.g. to Azure) | Upload the folder to the new host, run `python3 tools/set_url.py https://new-host/path/`, bump the version, and **Update** the manifest in Integrated apps. |

### Changing the URL

The manifest points at `https://cooperi29r09f2jf.github.io/gna-toolkit/`. If your GitHub username or repo name is different:

```
python3 tools/set_url.py https://YOUR-USERNAME.github.io/YOUR-REPO/
```

---

## What's in the folder

```
manifest.xml          add-in definition (ribbon, URLs); upload this in the admin center
taskpane.html         the side pane (Library, Starters, Tools, Check, Guide tabs)
commands.html         loads the one-click ribbon commands
index.html            short help page (linked from the add-in's support link)
src/
  gna-core.js         everything that talks to PowerPoint (insert, tools, checker)
  rules.js            house-style rules as plain functions (unit tested)
  taskpane.js/.css    pane UI
  commands.js         ribbon commands
  catalog.js          generated slide catalog
assets/
  slides/sNN.pptx     one file per template slide (layouts trimmed to keep them small)
  starters/*.pptx     the four starter decks (full master and layouts)
  thumbs/, guide/     thumbnails and style guide images
  icons/              ribbon icons
tools/
  catalog.json        slide names, categories, starter lineups (edit this)
  build_assets.py     regenerates assets from the template
  set_url.py          repoints the manifest
  bump_version.py     cache-busting version bump
  make_icons.py       redraws the ribbon icons
tests/
  rules.test.js       node tests/rules.test.js
  ui_test.py          python3 tests/ui_test.py (needs Playwright), drives the pane against a mock PowerPoint
```

## How it works, briefly

- Every library slide and starter deck carries hidden PowerPoint tags (`GNA_TEMPLATE_ID`, `GNA_TEMPLATE_VERSION`). The checker uses them to know which slides are covers, agendas, dividers and closings (no source line needed).
- **Auto formatting** inserts slides with the open deck's theme when the deck already has the GNA master, so duplicate masters don't pile up. In any other deck it keeps the template formatting.
- The template fixes are applied at build time: "Delive-rables" on slide 57 becomes "Deliverables". Embedded add-in panes and co-authoring history are stripped from the asset files.

## Known limits

- **Text highlight**: add-ins can't apply yellow text highlighting yet. The [NEEDED] box uses a yellow fill; for wrapped text, apply Home > Text Highlight Color yourself.
- **Harvey balls** go in as small pictures, not editable pie shapes.
- **Amber** is not in the template palette. The status dots use Office amber (#FFC000); change `GNA.COLORS.amber` in `src/gna-core.js` if GNA picks a different one.
- **Not checkable by software**: one message per slide, whether titles tell the story, one teal highlight per slide, white backgrounds. The checker reminds you to read the titles in order.
- **Table fixes** (dashes inside tables) need a recent PowerPoint build; on older ones the checker reports them for a manual fix.
- **Ribbon commands** (Sticker, Source line, [NEEDED] flag) show no error message if nothing is selected. Use the Tools pane to see errors.
- **Template slide 58** has "Deliver-ables" hyphenated in a narrow box. It's flagged by the checker as a typo; widening that box in the template is the real fix.

## Roadmap (v2)

- Claude drafting from the pane (action titles, executive summary from intake notes) through a small backend so the API key stays private.
- Agenda builder that fills section tags from the agenda slide.
- Sharing per-user favorites across the team.
