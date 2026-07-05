const express = require('express');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const { execFile } = require('child_process');
const util = require('util');
const path = require('path');
const fs = require('fs');
const os = require('os');

const execFileAsync = util.promisify(execFile);

const app = express();
const PORT = process.env.PORT || 3000;

// Path to the LibreOffice CLI binary. Override with an env var if the
// binary isn't on PATH (common on Windows / some Mac installs), e.g.:
//   Windows: set SOFFICE_PATH="C:\Program Files\LibreOffice\program\soffice.exe"
//   macOS:   export SOFFICE_PATH="/Applications/LibreOffice.app/Contents/MacOS/soffice"
const SOFFICE_BIN = process.env.SOFFICE_PATH || 'soffice';

// Path to Python 3. Override with PYTHON_PATH if needed.
const PYTHON_BIN = process.env.PYTHON_PATH || 'python3';

const MAX_FILE_SIZE = 150 * 1024 * 1024; // 150 MB
const ACCEPTED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm', '.csv', '.ods'];
const OPENPYXL_NATIVE = ['.xlsx', '.xlsm']; // formats openpyxl can open directly
const PREPARE_SCRIPT = path.join(__dirname, 'prepare.py');

app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({
  dest: os.tmpdir(),
  limits: { fileSize: MAX_FILE_SIZE }
});

app.post('/convert', (req, res) => {
  upload.single('file')(req, res, (err) => {
    if (err) {
      const message = err.code === 'LIMIT_FILE_SIZE'
        ? 'File is too large. Max size is 25 MB.'
        : 'Upload failed.';
      return res.status(400).json({ error: message });
    }
    handleConvert(req, res).catch((error) => {
      console.error('Unhandled conversion error:', error);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Something went wrong during conversion.' });
      }
    });
  });
});

async function handleConvert(req, res) {
  if (!req.file) {
    return res.status(400).json({ error: 'No file was received.' });
  }

  const originalName = req.file.originalname || 'spreadsheet.xlsx';
  const ext = path.extname(originalName).toLowerCase();

  if (!ACCEPTED_EXTENSIONS.includes(ext)) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Please upload a spreadsheet file (.xlsx, .xls, .xlsm, .csv, or .ods).' });
  }

  const workDir = path.join(os.tmpdir(), 'presswork-xl-' + uuidv4());
  fs.mkdirSync(workDir, { recursive: true });

  const inputPath = path.join(workDir, 'input' + ext);
  fs.renameSync(req.file.path, inputPath);

  try {
    // Step 1: get a plain .xlsx to work with.
    let rawXlsxPath;
    if (ext === '.csv') {
      rawXlsxPath = path.join(workDir, 'raw.xlsx');
      await execFileAsync(PYTHON_BIN, [PREPARE_SCRIPT, 'csv', inputPath, rawXlsxPath], { timeout: 60000 });
    } else if (OPENPYXL_NATIVE.includes(ext)) {
      rawXlsxPath = inputPath;
    } else {
      // .xls / .ods — convert to .xlsx first so openpyxl can style it.
      await execFileAsync(
        SOFFICE_BIN,
        ['--headless', '--norestore', '--convert-to', 'xlsx', '--outdir', workDir, inputPath],
        { timeout: 60000 }
      );
      rawXlsxPath = path.join(workDir, 'input.xlsx');
      if (!fs.existsSync(rawXlsxPath)) {
        throw new Error('Could not convert the legacy file to .xlsx for styling.');
      }
    }

    // Step 2: apply table styling — borders, header row, fit-to-width, repeated header.
    const styledXlsxPath = path.join(workDir, 'styled.xlsx');
    try {
      await execFileAsync(PYTHON_BIN, [PREPARE_SCRIPT, 'style', rawXlsxPath, styledXlsxPath], { timeout: 60000 });
    } catch (styleErr) {
      console.error('Styling step failed, falling back to unstyled conversion:', styleErr.message);
      fs.copyFileSync(rawXlsxPath, styledXlsxPath);
    }

    // Step 3: convert the styled workbook to PDF.
    await execFileAsync(
      SOFFICE_BIN,
      ['--headless', '--norestore', '--convert-to', 'pdf:calc_pdf_Export', '--outdir', workDir, styledXlsxPath],
      { timeout: 60000 }
    );

    const outputPath = path.join(workDir, 'styled.pdf');
    if (!fs.existsSync(outputPath)) {
      throw new Error('Conversion finished but no PDF was produced.');
    }

    const downloadName = originalName.replace(/\.[^.]+$/i, '.pdf');
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${downloadName.replace(/"/g, '')}"`);

    const stream = fs.createReadStream(outputPath);
    stream.pipe(res);
    stream.on('close', () => cleanup(workDir));
    stream.on('error', () => cleanup(workDir));
  } catch (error) {
    console.error('Conversion pipeline failed:', error.message);
    cleanup(workDir);
    if (!res.headersSent) {
      res.status(500).json({
        error: 'Conversion failed on the server. Make sure LibreOffice and Python (with openpyxl) are installed — see README.'
      });
    }
  }
}

function cleanup(dir) {
  fs.rm(dir, { recursive: true, force: true }, () => {});
}

app.listen(PORT, () => {
  console.log(`Presswork (Excel) server running at http://localhost:${PORT}`);
  console.log(`Using LibreOffice binary: ${SOFFICE_BIN}`);
  console.log(`Using Python binary: ${PYTHON_BIN}`);
});
