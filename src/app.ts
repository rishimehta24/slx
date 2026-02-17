import express, { Request, Response } from "express";
import multer from "multer";
import path from "path";
import fs from "fs";
import { convert } from "./convertPdfToXls";
import archiver from "archiver";
import os from "os";
import crypto from "crypto";

const app = express();
const PORT = parseInt(process.env.PORT ?? "5000", 10);
const SECRET_KEY = process.env.FLASK_SECRET_KEY ?? "dev-secret-key";

app.use(express.urlencoded({ extended: true }));

const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    if (!file.originalname.toLowerCase().endsWith(".pdf")) {
      cb(new Error("Only PDF files are allowed"));
      return;
    }
    cb(null, true);
  },
});

interface ConvertedFile {
  filename: string;
  content: Buffer;
}

interface JobResult {
  files: ConvertedFile[];
  zipContent: Buffer;
}

const jobStore = new Map<string, JobResult>();
const MAX_JOBS = 20;

function uniqueOutputName(name: string, used: Set<string>): string {
  if (!used.has(name)) {
    used.add(name);
    return name;
  }
  const ext = path.extname(name);
  const stem = path.basename(name, ext);
  let counter = 1;
  let candidate = `${stem}_${counter}${ext}`;
  while (used.has(candidate)) {
    counter++;
    candidate = `${stem}_${counter}${ext}`;
  }
  used.add(candidate);
  return candidate;
}

function buildZip(files: ConvertedFile[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const archive = archiver("zip", { zlib: { level: 9 } });
    const chunks: Buffer[] = [];
    archive.on("data", (chunk: Buffer) => chunks.push(chunk));
    archive.on("end", () => resolve(Buffer.concat(chunks)));
    archive.on("error", reject);
    for (const f of files) {
      archive.append(f.content, { name: f.filename });
    }
    archive.finalize();
  });
}

function renderIndex(options: {
  messages?: Array<{ category: string; message: string }>;
  convertedFiles?: ConvertedFile[];
  jobId?: string;
}): string {
  const { messages = [], convertedFiles = null, jobId = "" } = options;
  const messagesHtml =
    messages.length > 0
      ? `<ul class="messages">${messages.map((m) => `<li class="${m.category}">${escapeHtml(m.message)}</li>`).join("")}</ul>`
      : "";
  const resultsHtml =
    convertedFiles && convertedFiles.length > 0 && jobId
      ? `
    <section class="results-card">
      <h2>Converted files</h2>
      <ul class="file-list">
        ${convertedFiles.map((f, i) => `<li><span class="name">${escapeHtml(f.filename)}</span><span class="file-actions"><a href="/download/${jobId}/${i}">Download</a></span></li>`).join("")}
      </ul>
      <div class="zip-row">
        <a class="zip-btn" href="/download/${jobId}/zip">📦 Download all as ZIP</a>
      </div>
    </section>`
      : "";
  return indexTemplate.replace("{{MESSAGES}}", messagesHtml).replace("{{RESULTS}}", resultsHtml);
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

let indexTemplate: string;
try {
  indexTemplate = fs.readFileSync(path.join(__dirname, "..", "views", "index.html"), "utf-8");
} catch (e) {
  console.error("Failed to load views/index.html:", e);
  indexTemplate = "<html><body><h1>Template not found</h1></body></html>";
}

app.get("/health", (_req: Request, res: Response) => {
  res.type("text/plain").status(200).send("ok");
});

app.get("/", (_req: Request, res: Response) => {
  res.send(renderIndex({}));
});

app.post("/", (req: Request, res: Response, next: express.NextFunction) => {
  upload.array("pdfs", 20)(req, res, (err: unknown) => {
    if (err) {
      const msg = err instanceof Error ? err.message : "Upload error";
      return res.send(renderIndex({ messages: [{ category: "error", message: msg }] }));
    }
    next();
  });
}, async (req: Request, res: Response) => {
  const files = (req.files as Express.Multer.File[]) ?? [];
  const uploads = files.filter((f) => f && f.originalname);
  const messages: Array<{ category: string; message: string }> = [];

  if (uploads.length === 0) {
    messages.push({ category: "error", message: "Please choose at least one PDF file to convert." });
    return res.send(renderIndex({ messages }));
  }

  const processedFiles: ConvertedFile[] = [];
  const usedNames = new Set<string>();
  const tmpDir = path.join(os.tmpdir(), `htmlpdf_${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });

  try {
    for (const file of uploads) {
      const originalName = file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_") || "upload.pdf";
      if (!originalName.toLowerCase().endsWith(".pdf")) {
        messages.push({ category: "warning", message: `${originalName} is not a PDF and was skipped.` });
        continue;
      }
      const pdfPath = path.join(tmpDir, originalName);
      fs.writeFileSync(pdfPath, file.buffer);
      const outputName = uniqueOutputName(path.basename(originalName, path.extname(originalName)) + ".xlsx", usedNames);
      const outputPath = path.join(tmpDir, outputName);
      try {
        await convert(pdfPath, outputPath);
        processedFiles.push({ filename: outputName, content: fs.readFileSync(outputPath) });
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        messages.push({ category: "warning", message: `Failed to convert ${originalName}: ${msg}` });
      }
    }
  } finally {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  }

  if (processedFiles.length === 0) {
    messages.push({ category: "error", message: "No files were converted. Please check the uploaded PDFs and try again." });
    return res.send(renderIndex({ messages }));
  }

  const zipContent = await buildZip(processedFiles);
  const jobId = crypto.randomUUID();
  jobStore.set(jobId, { files: processedFiles, zipContent });
  while (jobStore.size > MAX_JOBS) {
    const firstKey = jobStore.keys().next().value;
    if (firstKey) jobStore.delete(firstKey);
  }

  return res.send(renderIndex({ messages, convertedFiles: processedFiles, jobId }));
});

app.get("/download/:jobId/:fileIndex", (req: Request, res: Response) => {
  const job = jobStore.get(req.params.jobId);
  const fileIndex = parseInt(req.params.fileIndex, 10);
  if (!job || fileIndex < 0 || fileIndex >= job.files.length) {
    return res.status(404).send("Not found");
  }
  const file = job.files[fileIndex];
  res.attachment(file.filename).type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").send(file.content);
});

app.get("/download/:jobId/zip", (req: Request, res: Response) => {
  const job = jobStore.get(req.params.jobId);
  if (!job) return res.status(404).send("Not found");
  res.attachment("converted_reports.zip").type("application/zip").send(job.zipContent);
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Server listening on http://0.0.0.0:${PORT}`);
});
