# Changelog

All notable changes to this project will be documented in this file.

### Added
- **Intelligent Engine Discovery**
    - New `_find_pdf2john()` utility: Replaces basic path checks with a robust multi-platform search. It now scans standard installation directories across Windows (`Program Files`), Linux (`/usr/share/john`, `/snap`), and macOS.
    - Support for all `pdf2john` variants: Automatically detects and handles native binaries, Perl scripts (`.pl`), and Python scripts (`.py`).
    - Added `probe_gpu()`: Automatically detects OpenCL, CUDA, Metal, and specific hardware (Nvidia, AMD, Intel Arc) via `hashcat -I` to provide real-time hardware availability feedback.

- **Advanced Hash Management**
    - New `extract_pdf_hash()`: A two-stage extraction process that validates output, writes to secure temporary files, and provides a localized hash preview.
    - Automated Hashcat Mode Detection: `detect_pdf_hash_mode()` parses PDF version fields to automatically map hashes to their specific Hashcat kernels:
        - `10400` (PDF 1.1–1.3)
        - `10500` (PDF 1.4–1.6)
        - `10600` (v4r4)
        - `10700` (v5r5/r6, AES-256)
        - `25400` (PDF 2.0)

- **Cracking Engine Overhaul**
    - **John the Ripper (CPU):** Integrated support for Wordlist (with rules), Incremental, Markov (customizable levels), and Mask modes.
    - **Hashcat (GPU):** Integrated Attack Modes 0 (Wordlist), 3 (Brute-force), and 6 (Hybrid). Includes workload profile management (1–4) and live status polling.
    - **Live Streaming:** Real-time `stdout` streaming via `Popen` allowing users to see progress line-by-line without UI freezing.

- **New "Crack Tab" UI/UX**
    - **Step-by-Step Workflow:** Guided interface from target selection to hash extraction and engine configuration.
    - **Target Probing:** Integrated `PyMuPDF`/`pikepdf` button to inspect PDF encryption details and permissions before attempting a crack.
    - **Dynamic UI States:** Conditional visibility for attack parameters (e.g., Markov spinners or Mask inputs only appear when the relevant mode is selected).
    - **Mask Presets:** Quick-action buttons for common patterns (e.g., 8-char brute force, `Cap+5lower+2digit`, `6-digits`).
    - **Dependency Badges:** Visual `✓/✗` status indicators for `john`, `hashcat`, and `pdf2john` in the header.

- **PDFRepairEngine (Forensic Recovery)**
    - **Heuristic Rebuilding:** Integrated `pikepdf` recovery mode (`attempt_recovery=True`). The engine now performs a full binary scan for `obj` headers to rebuild corrupted XREF tables, allowing access to files that fail to open in standard viewers.
    - **Deep Triage System:** Automated scanning for linearization errors, object number gaps (truncation detection), and structural warnings.
    - **Orphan Analysis:** Implemented a Depth-First Search (DFS) walker starting from the `/Root` catalog. It identifies and quantifies "orphan" objects that are unreachable via the object graph, enabling cleaner, smaller file saves.
    - **Stream Intelligence:** Added a filter summary tool that tallies stream types (e.g., `FlateDecode`, `DCTDecode`), providing a high-level view of text vs. image data density.

- **PDFFlagPanel (Inspection & Modification UI)**
    - Added a collapsible, multi-tabbed diagnostic panel located below the main HDC bar, featuring a real-time "Health Badge" in the header.
    - **⚕ Health / XREF Tab:** 
        - Visual triage report with color-coded severity icons.
        - "Repair & Save" workflow to export recovered versions of damaged documents.
        - Owner-level password support for triaging encrypted streams.
    - **🔒 Permissions Tab:** 
        - Full exposure of the 32-bit `/P` integer field via ISO 32000-1 compliant checkboxes.
        - Granular control over Print (Low/High), Modify, Copy, Annotate, Form Fill, and Accessibility flags.
        - "Unlock All" / "Lock All" macros for rapid restriction stripping.
        - Live Binary/Hex preview of the `/P` field for cryptographic verification.
    - **🗂 Objects Tab:** 
        - Searchable `ttk.Treeview` listing every object, its type, subtype, and byte count.
        - Real-time type-filtering (e.g., "Image", "Font") and clickable column headers for sorting.
    - **📋 Metadata Tab:** 
        - Editable interface for the eight primary `/Info` fields.
        - "Privacy Scrub" feature: One-click metadata wiping to sanitize documents before distribution.

### Changed
- **Session Handling:** Moved to a robust `--session` and `--restore` system. The UI now includes a dedicated "Restore" checkbox to resume interrupted jobs using JtR `.rec` files.
- **Security:** Temporary files are no longer stored in the source directory to prevent clutter and accidental data exposure.
- **Progress Monitoring:** Refined the "Stop" process with a confirmation dialog explaining how session files preserve progress.
- **Optimization:** The PDFFlagPanel now initializes in a background thread with a 100ms delay after file load, ensuring heavy object-graph walking never blocks the primary page rendering or UI responsiveness.
- **Saving Logic:** 
    - Default saving now disables linearization (`linearize=False`) for cleaner surgical edits.
    - Enabled `ObjectStreamMode.generate` by default to re-pack compressed streams and resolve stream-level corruption during export.

### Fixed
- Fixed issue where the UI would hang during long-running shell executions.
- Resolved pathing issues on Windows when calling Perl-based `pdf2john` scripts.
- Corrected hashcat mode mapping for Acrobat 9/X/XI (AES-256) files.
- Resolved UI "freezing" when opening large PDFs with deep object nesting.
- Fixed a bug where files with broken XREF trailers would fail to load despite containing valid object data.
- Many other bug fixes.