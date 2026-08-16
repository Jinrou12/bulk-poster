# Bulk Poster Generator (កម្មវិធីបង្កើត Poster ថ្លែងអំណរគុណជាកញ្ចប់)

A modern, fast, web-based application for creating bulk thank-you, appreciation, and ceremonial posters with Excel/CSV data merging, dynamic text and avatar frames, Khmer typography support, and batch export to ZIP.

## 🌟 Features
- **Upload Background Template**: Support for custom image backgrounds (PNG, JPG, WebP) or built-in Cambodian traditional ceremony template.
- **Visual Canvas Editor**: Draggable, resizable text boxes & photo frames with real-time overlay drag handles.
- **Khmer Typography Support**: Rich selection of Google Khmer Fonts (Moul, Battambang, Siemreap, Kantumruy Pro, Hanuman, Bayon, Koh Santepheap, Dangrek) with custom stroke outlines and text shadows.
- **Excel Data Import**: SheetJS integration supporting `.xlsx`, `.xls`, `.csv` files. Supports files with or without header rows.
- **Dynamic Merge Fields**: Auto-detects Excel column names and generates merge field tags (e.g., `{{ Name }}`, `{{ Number }}`).
- **Live Record Preview**: Browse through individual records record-by-record with instant live rendering on canvas.
- **Batch Export to ZIP**: Renders all high-resolution poster images asynchronously and packages them into a single downloadable ZIP file.

## 🚀 Live Demo
Deployable on Vercel, Netlify, or GitHub Pages as a static web application.

## 🛠 Tech Stack
- HTML5, CSS3 (Glassmorphism UI)
- Vanilla JavaScript (ES6 Canvas API)
- [SheetJS (xlsx)](https://sheetjs.com/) for Excel parsing
- [JSZip](https://stuk.github.io/jszip/) & [FileSaver.js](https://github.com/eligrey/FileSaver.js/) for ZIP packaging
- Google Fonts (Khmer Typography)
