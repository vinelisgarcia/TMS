# Calendar report generators

Locally hosted, lazy-loaded browser builds. No report data is sent to a third party.

- ExcelJS 4.4.0: https://github.com/exceljs/exceljs (MIT)
- jsPDF 4.2.1: https://github.com/parallax/jsPDF (MIT)
- jsPDF-AutoTable 5.0.9: https://github.com/simonbengtsson/jsPDF-AutoTable (MIT)
- PDF.js 6.5.438, legacy browser build and matching worker:
  https://github.com/mozilla/pdf.js (Apache-2.0). Used for local text extraction and page rendering. Matching `cmaps`,
  `standard_fonts` and image-decoding `wasm` resources are under
  `pdfjs-assets-6.5.438/`; upstream notices are preserved there.

Downloaded from version-pinned npm distributions (npm registry or jsDelivr). License files
are alongside the distributions; AutoTable also includes its license header.

# Packing list scan reader

Self-hosted and loaded only when OCR is needed; document images stay in the browser.
Official npm tarballs, versions pinned:

- `tesseract.js@7.0.0`: `dist/tesseract.min.js`, `dist/worker.min.js`, Apache-2.0.
  https://github.com/naptha/tesseract.js
- `tesseract.js-core@7.0.0`: all six embedded-WASM browser variants, Apache-2.0.
  https://github.com/naptha/tesseract.js-core
- `@tesseract.js-data/eng@1.0.0` and `@tesseract.js-data/spa@1.0.0`:
  `4.0.0_best_int/{eng,spa}.traineddata.gz`. npm packages declare MIT; the
  underlying Tesseract training data uses Apache-2.0 (notice included).
  https://github.com/naptha/tessdata
  https://github.com/tesseract-ocr/tessdata_best

Core scripts embed their WASM, so separate `.wasm` binaries are not needed.
Language files are the only data cached by the OCR runtime. No source documents
or recognized text are placed in that cache.
