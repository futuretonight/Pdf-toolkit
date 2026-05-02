# Changelog — pdf toolkit V3 (PC Edition)

---

## [2.0.0] — 2025-05 · "Structural DNA" release

This is a major architecture update. It introduces a fully custom
hyperdimensional PDF viewer, a multi-stage OpenCV restoration pipeline,
live before/after previews in Compress and Restore, and a wave of UI
stability fixes. Every new feature is optional and degrades gracefully
if its dependencies are not installed.

---

### New — 📄 Viewer tab (HDC-powered PDF viewer)

A completely custom PDF viewer built around a **Hyperdimensional
Computing (HDC) engine** rather than a traditional page-tree.  No
third-party viewer library is used.

#### HDC Engine (`HDCVec`, `PDFStructuralParser`, `HDCDocumentIndex`)

- **Dimension: 8 192 bits** (128 × `uint64`).  
  Chosen over 10 000 (not cache-aligned, ~20 % slower SIMD XOR) and
  over 4 096 (majority-vote accuracy degrades faster with many
  components per page, Hamming resolution halved).  
  Memory cost: exactly **1 KB per page** — a 1 000-page document
  occupies 1 MB of vectors.

- **Structural DNA encoding** — no pixel data is ever stored in the
  vector.  For each page the parser extracts:
  - Font names / types (up to 20)
  - Text word tokens (up to 60, deterministically sub-sampled)
  - Quantised page geometry (rounded to 72 pt grid for layout stability)
  - Image XREF IDs (up to 16)

  These are encoded into a single 8 192-bit `PageVec` using the HDC
  bind / bundle algebra:

  ```
  PageVec = bind(KEY_FONT,   bundle(font vecs))
          ⊕ bind(KEY_TEXT,   bundle(text vecs))
          ⊕ bind(KEY_LAYOUT, layout_vec)
          ⊕ bind(KEY_IMAGE,  bundle(image vecs))
  ```

  Each component vector is generated deterministically from its string
  identity via **SHAKE-256** (variable-length hash, single call for
  exactly 1 024 bytes of entropy).

- **`HDCVec` operations**
  - `bind(other)` — XOR, reversible association
  - `bundle(vecs)` — majority vote with deterministic tie-breaking
  - `hamming(other)` — differing-bit count via `numpy.unpackbits`
  - `similarity(other)` — `1 − hamming / 8192`  (0.5 = random, 1.0 = identical)
  - `permute(k)` — cyclic bit-rotation, encodes sequence / order
  - `from_string(s)` — SHAKE-256 deterministic construction
  - `to_bytes()` / `from_bytes()` — 1 024-byte SQLite serialisation

- **`HDCDocumentIndex`** — builds, caches and queries the full index.
  - Cache key: `(canonical_path, mtime, fsize, sha256[:16 of first 4 KB])`  
    File changes auto-invalidate the cache row.
  - Storage: `~/.pdftoolkit/hdc_cache.db` (SQLite, ~1 KB × n_pages).
  - `nearest_pages(page, n)` — O(n_pages × 128) XOR ops, typically < 2 ms
    for a 500-page document.
  - `xor_navigate(page, query_vec)` — semantic jump: XOR current PageVec
    toward a query concept, Hamming-search for the closest real page.
    Example: XOR with `KEY_IMAGE` jumps to the most image-heavy page
    structurally similar to the current one.
  - `similarity_map(page)` — full `{page_num: float}` similarity dict,
    used to colour thumbnail similarity bars.

- **`PDFStructuralParser`** — renderer-agnostic DNA extractor.
  - PyMuPDF path: `page.get_fonts()` + `page.get_text("words")` +
    `page.get_images()` — fast, no subprocess.
  - pypdf fallback: `/Resources/Font` dict + `extract_text()`.

#### Predictive Pre-Renderer (`PredictivePreRenderer`)

- **Optional** — toggled by the ⚡ Pre-render checkbox in the viewer nav.
- `ThreadPoolExecutor(max_workers=2)` runs in the background.
- `prime(current_page)` combines two prediction signals:
  - **Sequential**: pages ±1, +2, +3 (most likely to be visited next)
  - **Semantic**: `nearest_pages()` top-4 (structurally similar pages)
- `evict_far(current, radius=12)` — drops cached pages more than 12
  positions away to cap RAM usage.
- When a page is already in the pre-render cache, the status bar shows
  `[⚡ pre-rendered]` and the page is displayed with zero render latency.

#### Viewer UI (`PDFViewerPane`)

- **Thumbnail strip** (left panel)
  - Lazy background rendering via `ThreadPoolExecutor`.
  - Per-page similarity bar (4 px, cyan = structurally similar to
    current page, dim = unrelated).  Updates on every page navigation.
  - Page label shows similarity % for pages above 25 % threshold.
  - Click any thumbnail to navigate.  Strip auto-scrolls to keep the
    current page visible.
  - Mousewheel scrolling bound on all three platform event variants
    (`<MouseWheel>`, `<Button-4>`, `<Button-5>`).

- **Main canvas**
  - Mousewheel zoom towards cursor (factor 1.12×/step, range 5–2000 %)
  - Click-drag pan
  - Double-click → fit to window
  - Keyboard: `← →` / `PgUp PgDn` for page navigation
  - `+` / `−` buttons, live zoom % label

- **HDC status bar** (bottom)
  - Shows dimension, total memory usage, current page, top-5 semantic
    similarity %, and a colour-coded bar.

- **Cache prompt** — after building a fresh index the app asks whether
  to save the vectors.  Answering Yes writes to SQLite in a daemon
  thread.  Next open loads in < 5 ms from cache.

---

### New — ✦ Restore tab: full OpenCV restoration pipeline

#### `apply_restoration_pipeline(img, ...)`

Replaces the single-pass Pillow `UnsharpMask` with a four-stage
OpenCV pipeline.  Each stage is independently controlled by a slider
and can be disabled individually.

| Stage | Library | Control |
|---|---|---|
| Denoise | `cv2.fastNlMeansDenoisingColored` | 0–30 (h-param) |
| CLAHE | `cv2.createCLAHE` in LAB space | on/off toggle |
| Sharpening | Laplacian unsharp (document) / `cv2.detailEnhance` (mixed) / Pillow UnsharpMask (photo) | 0–250 % |
| Bilateral | `cv2.bilateralFilter` | 0–15 (diameter) |

- **Mode selector**: `document` (text / diagrams, hard Laplacian
  kernel), `photo` (natural images, Pillow soft mask), `mixed`
  (OpenCV `detailEnhance` with tunable sigma).
- **CLAHE**: runs in CIE LAB `L` channel to boost local contrast
  without blowing out highlights.  `clipLimit` is 2.5 for document
  mode, 1.8 for photo mode.
- **Fallback**: if OpenCV / numpy are absent, the pipeline gracefully
  falls back to Pillow contrast + UnsharpMask only.

#### `restore_pdf_quality(src, out, ...)`

- Renderer priority: **PyMuPDF → pypdfium2 → pdf2image/poppler**
  (same chain as the viewer).
- All four pipeline parameters plus render DPI are now exposed and
  wired through from the UI sliders.
- Output filename includes both DPI and mode:
  `original_restored_300dpi_document.pdf`

#### Restore tab live preview

- Full `PreviewPane` (before / after, zoom / pan) on the right half.
- Before panel: raw page render at preview DPI.
- After panel: same render passed through the full OpenCV pipeline
  with the current slider values — what you see is exactly what the
  output file will look like.
- All sliders trigger `schedule_refresh(500 ms)` — debounced to avoid
  hammering the CPU on every tick.

---

### New — ⊟ Compress tab: live preview + free DPI slider

#### GS DPI slider

- Replaces the old `screen / ebook / printer / prepress` combobox.
- Free range 30–600 DPI, resolution 5.
- Live label updates with a human hint:
  `screen (tiny, very pixelated)` → `web / ebook` → `balanced` →
  `printer quality` → `prepress / hi-res`.
- Slider uses `.place(relwidth=1.0)` inside a `pack_propagate(False)`
  container — see UI Fixes below.

#### Compress tab live preview

- Before panel: page rendered at preview DPI.
- After panel: same render passed through `simulate_gs_compression()`,
  which performs a real JPEG re-encode via OpenCV at the matching
  quality level, producing accurate artefact preview.
- Click any file in the batch listbox to preview that file.

---

### New — Renderer stack (`render_pdf_page_pil`)

All rendering throughout the app (viewer, compress preview, restore
preview, pre-renderer, thumbnails) now uses a unified renderer with
priority fallback:

1. **PyMuPDF** (`import pymupdf as fitz`) — in-process, no system
   deps, sub-pixel AA, fastest
2. **pypdfium2** — Google PDFium via Python wheel, also no system deps
3. **pdf2image / poppler** — original fallback, needs system poppler

`count_pdf_pages()` also uses the same priority chain to avoid
spawning a subprocess just to count pages.

---

### New — `ScrollableFrame`

A vertically scrollable `ttk.Frame` wrapper used by both the Compress
and Restore left panels.  Exposes `.inner` as the child widget parent.
Mousewheel bound on all three platform events.  `_on_canvas_configure`
stretches `.inner` to the canvas width so `fill="x"` children work
normally.

---

### Fixed — PyMuPDF import collision (#1)

`import fitz` could resolve to an unrelated `frontend` package
(common in environments with web-UI libraries installed), causing:

```
RuntimeError: Directory 'static/' does not exist
```

**Fix**: now imports `pymupdf as fitz` first (PyMuPDF ≥ 1.24 canonical
name, avoids the collision entirely), with a guard on the old `fitz`
path that checks `fitz.open` is callable before accepting it.

---

### Fixed — Ghostscript platform detection on x64 Windows (#2)

The old code:

```python
gs_bin = "gswin64c" if sys.platform == "win32" else "gs"
```

`sys.platform` returns `"win32"` on both 32-bit and 64-bit Windows,
but if only `gswin32c` or a PATH-installed `gs` was present the old
code would silently fail.

**Fix**: `_find_ghostscript()` tries `gswin64c → gswin32c → gs` in
order using `shutil.which()`.  If none are found, the error message
names all three candidates and links to the GS download page.

---

### Fixed — UI jitter on window resize (#3)

**Root cause**: `_on_canvas_resize` was calling `schedule_refresh()`
(background re-render) on every pixel of window resize, causing
constant flicker and CPU churn.

**Fix**: two operations that were conflated are now separate:

- **Window resize** → `_on_canvas_configure` calls `_redraw_one()` /
  `_draw_main()` only, which re-fits the already-cached PIL image into
  the new canvas size.  Pure Python maths + one `ImageTk.PhotoImage`
  allocation, < 5 ms, runs on the main thread, zero jitter.
- **Slider change** → `schedule_refresh(500 ms)` debounces and then
  launches the actual background re-render.

`tk.Scale` widgets are now placed with `.place(relwidth=1.0)` inside
a `pack_propagate(False)` container.  This completely stops the
`Scale` from participating in the Tk geometry cascade — the primary
source of the jitter loop.

A render generation counter (`_render_gen`) ensures stale results from
a render that was overtaken by a newer slider move are silently dropped
rather than flashing onto the screen.

---

### Fixed — Compress/Restore left panel scroll (#4)

Controls in the Restore tab overflowed off-screen on small windows
with no way to scroll to them.

**Fix**: Both left panels are now wrapped in `ScrollableFrame`.  The
canvas `itemconfigure` call stretches the inner frame to the canvas
width on resize, so controls remain full-width as the sash is dragged.

---

### Fixed — `PreviewPane` magnification / zoom (#5)

Previous version had no zoom or pan capability on preview canvases.

**Added to both Compress and Restore preview panes**:
- **Mousewheel** → zoom ×1.10 per tick, centred on cursor position
- **Click-drag** → pan
- **Double-click** → reset zoom to 100 % and re-centre
- Zoom % label and pixel dimensions shown below each canvas
  (turns cyan when not at 100 %)
- Hint: `scroll=zoom · drag=pan · dbl-click=reset`

PIL images are cached as `_pil_b` / `_pil_a` so zoom/pan is O(1)
(no re-render, just resize the cached PIL and swap the `PhotoImage`).

---

### Dependencies added

| Package | Use | Install |
|---|---|---|
| `pymupdf` | Primary PDF renderer (PyMuPDF ≥ 1.24) | `pip install pymupdf` |
| `pypdfium2` | Secondary renderer (Google PDFium) | `pip install pypdfium2` |
| `numpy` | HDC vector algebra, OpenCV bridge | `pip install numpy` |
| `opencv-python` | Denoise, CLAHE, sharpen, bilateral | `pip install opencv-python` |

Stdlib additions: `hashlib`, `sqlite3`, `struct`, `concurrent.futures`,
`pathlib` (all standard, no install needed).

All new dependencies are optional with graceful fallback — the app
continues to run with degraded functionality if any are absent.

---

### New files / paths created at runtime

| Path | Contents |
|---|---|
| `~/.pdftoolkit/hdc_cache.db` | SQLite HDC vector cache (~1 KB/page) |
| `<source>_compressed.pdf` | Batch compress output |
| `<source>_restored_<DPI>dpi_<mode>.pdf` | Restore output |