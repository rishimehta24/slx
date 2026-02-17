# Incident PDF → Excel Converter

Convert incident analysis PDF reports into formatted Excel (XLSX) spreadsheets. Upload one or more PDFs via the web UI and download individual files or a ZIP of all converted reports.

## Requirements

- **Node.js** ≥ 18

## Setup

```bash
npm install
```

## Run locally

**Development** (with hot reload):

```bash
npm run dev
```

**Production** (builds then starts):

```bash
npm start
```

By default the server listens on **http://localhost:5000**. Set the `PORT` environment variable to use another port.

## Usage

1. Open the app in your browser.
2. Select one or more incident PDF files (`.pdf`).
3. Click **Convert to Excel**.
4. Download individual XLSX files or **Download all as ZIP**.

## Deploy

- **Railway / Nixpacks**: Uses `nixpacks.toml`; build runs `npm run build`, start runs `npm start`.
- **Heroku**: Use the included `Procfile` (`web: npm start`).

Set `PORT` in the environment; the app binds to `0.0.0.0`.

## Tech

- **Express** – web server and upload handling  
- **pdf-parse** – PDF text extraction  
- **ExcelJS** – XLSX generation  
- **TypeScript** – implementation

## License

Private / unlicensed.
