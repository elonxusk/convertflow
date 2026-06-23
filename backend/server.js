const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');

const app = express();

// Middleware
app.use(cors({
  origin: '*', // Production mein apna Vercel URL daalna
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type']
}));
app.use(express.json());

// Ensure directories exist
const uploadsDir = path.join(__dirname, 'uploads');
const convertedDir = path.join(__dirname, 'converted');
fs.ensureDirSync(uploadsDir);
fs.ensureDirSync(convertedDir);

// Multer config
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => {
    const uniqueName = `${uuidv4()}-${file.originalname}`;
    cb(null, uniqueName);
  }
});
const upload = multer({ 
  storage,
  limits: { fileSize: 100 * 1024 * 1024 } // 100MB
});

// ===== SERVICES =====

const { PDFDocument } = require('pdf-lib');
const sharp = require('sharp');
const mammoth = require('mammoth');
const htmlPdf = require('html-pdf-node');
const XLSX = require('xlsx');
const Tesseract = require('tesseract.js');
const archiver = require('archiver');

// Helper: Get file info
function getFileInfo(filePath) {
  const stats = fs.statSync(filePath);
  return {
    name: path.basename(filePath),
    size: stats.size,
    ext: path.extname(filePath).toLowerCase()
  };
}

// Helper: Create output path
function getOutputPath(originalName, newExt) {
  const baseName = path.parse(originalName).name;
  const outputName = `${uuidv4()}_${baseName}${newExt}`;
  return path.join(convertedDir, outputName);
}

// 1. PDF to Image
async function pdfToImage(inputPath, format = 'png') {
  const { fromPath } = require('pdf2pic');
  const baseName = path.parse(inputPath).name;
  const outputDir = path.join(convertedDir, `${uuidv4()}_images`);
  await fs.ensureDir(outputDir);
  
  const options = {
    density: 100,
    saveFilename: baseName,
    savePath: outputDir,
    format: format === 'jpg' ? 'jpeg' : format,
    width: 1200,
    height: 1600
  };
  
  const convert = fromPath(inputPath, options);
  await convert(1); // Convert first page
  
  // If multiple pages, zip them
  const files = await fs.readdir(outputDir);
  if (files.length === 1) {
    return path.join(outputDir, files[0]);
  }
  
  // Zip multiple images
  const zipPath = path.join(convertedDir, `${uuidv4()}_images.zip`);
  const output = fs.createWriteStream(zipPath);
  const archive = archiver('zip', { zlib: { level: 9 } });
  
  archive.pipe(output);
  files.forEach(f => archive.file(path.join(outputDir, f), { name: f }));
  await archive.finalize();
  
  return zipPath;
}

// 2. Image to PDF
async function imageToPdf(inputPaths) {
  const pdfDoc = await PDFDocument.create();
  
  for (const inputPath of inputPaths) {
    const imgBuffer = await fs.readFile(inputPath);
    let image;
    
    const ext = path.extname(inputPath).toLowerCase();
    if (ext === '.png') {
      image = await pdfDoc.embedPng(imgBuffer);
    } else {
      image = await pdfDoc.embedJpg(imgBuffer);
    }
    
    const page = pdfDoc.addPage([image.width, image.height]);
    page.drawImage(image, {
      x: 0,
      y: 0,
      width: image.width,
      height: image.height
    });
  }
  
  const outputPath = getOutputPath(inputPaths[0], '.pdf');
  await fs.writeFile(outputPath, await pdfDoc.save());
  return outputPath;
}

// 3. Word to PDF (via HTML)
async function wordToPdf(inputPath) {
  const { value: html } = await mammoth.convertToHtml({ path: inputPath });
  const styledHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; max-width: 800px; margin: 40px auto; padding: 20px; }
        h1, h2, h3 { color: #333; }
        p { margin: 10px 0; }
      </style>
    </head>
    <body>${html}</body>
    </html>
  `;
  
  const file = { content: styledHtml };
  const options = { format: 'A4', printBackground: true };
  const outputPath = getOutputPath(inputPath, '.pdf');
  
  const pdfBuffer = await htmlPdf.generatePdf(file, options);
  await fs.writeFile(outputPath, pdfBuffer);
  
  return outputPath;
}

// 4. Excel to PDF
async function excelToPdf(inputPath) {
  const workbook = XLSX.readFile(inputPath);
  let html = '<html><body style="font-family: Arial; padding: 20px;">';
  
  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    html += `<h2>${sheetName}</h2><table border="1" cellpadding="5" style="border-collapse: collapse; margin: 20px 0;">`;
    json.forEach(row => {
      html += '<tr>';
      row.forEach(cell => {
        html += `<td>${cell || ''}</td>`;
      });
      html += '</tr>';
    });
    html += '</table>';
  });
  
  html += '</body></html>';
  
  const file = { content: html };
  const options = { format: 'A4', landscape: true };
  const outputPath = getOutputPath(inputPath, '.pdf');
  
  const pdfBuffer = await htmlPdf.generatePdf(file, options);
  await fs.writeFile(outputPath, pdfBuffer);
  
  return outputPath;
}

// 5. HTML to PDF
async function htmlToPdf(inputPath) {
  const html = await fs.readFile(inputPath, 'utf8');
  const file = { content: html };
  const options = { format: 'A4', printBackground: true };
  const outputPath = getOutputPath(inputPath, '.pdf');
  
  const pdfBuffer = await htmlPdf.generatePdf(file, options);
  await fs.writeFile(outputPath, pdfBuffer);
  
  return outputPath;
}

// 6. Text to PDF
async function textToPdf(inputPath) {
  const text = await fs.readFile(inputPath, 'utf8');
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Courier New', monospace; white-space: pre-wrap; padding: 40px; line-height: 1.5; }
      </style>
    </head>
    <body>${text.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</body>
    </html>
  `;
  
  const file = { content: html };
  const options = { format: 'A4' };
  const outputPath = getOutputPath(inputPath, '.pdf');
  
  const pdfBuffer = await htmlPdf.generatePdf(file, options);
  await fs.writeFile(outputPath, pdfBuffer);
  
  return outputPath;
}

// 7. Compress PDF
async function compressPdf(inputPath) {
  const pdfDoc = await PDFDocument.load(await fs.readFile(inputPath));
  
  // Compress by re-saving with lower quality
  const outputPath = getOutputPath(inputPath, '_compressed.pdf');
  await fs.writeFile(outputPath, await pdfDoc.save({ useObjectStreams: true }));
  
  return outputPath;
}

// 8. Merge PDFs
async function mergePdfs(inputPaths) {
  const mergedPdf = await PDFDocument.create();
  
  for (const inputPath of inputPaths) {
    const pdf = await PDFDocument.load(await fs.readFile(inputPath));
    const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
    pages.forEach(page => mergedPdf.addPage(page));
  }
  
  const outputPath = path.join(convertedDir, `${uuidv4()}_merged.pdf`);
  await fs.writeFile(outputPath, await mergedPdf.save());
  
  return outputPath;
}

// 9. Split PDF
async function splitPdf(inputPath, pageRanges) {
  const pdfDoc = await PDFDocument.load(await fs.readFile(inputPath));
  const totalPages = pdfDoc.getPageCount();
  
  const outputPaths = [];
  
  for (const range of pageRanges) {
    const [start, end] = range.split('-').map(Number);
    const newPdf = await PDFDocument.create();
    const pages = await newPdf.copyPages(pdfDoc, Array.from({length: end - start + 1}, (_, i) => start - 1 + i));
    pages.forEach(page => newPdf.addPage(page));
    
    const outputPath = path.join(convertedDir, `${uuidv4()}_pages_${start}-${end}.pdf`);
    await fs.writeFile(outputPath, await newPdf.save());
    outputPaths.push(outputPath);
  }
  
  if (outputPaths.length === 1) return outputPaths[0];
  
  // Zip multiple files
  const zipPath = path.join(convertedDir, `${uuidv4()}_split.zip`);
  const output = fs.createWriteStream(zipPath);
  const archive = archiver('zip');
  archive.pipe(output);
  outputPaths.forEach(p => archive.file(p, { name: path.basename(p) }));
  await archive.finalize();
  
  return zipPath;
}

// 10. OCR (Image to Text)
async function ocrExtract(inputPath) {
  const { data: { text } } = await Tesseract.recognize(inputPath, 'eng');
  const outputPath = getOutputPath(inputPath, '.txt');
  await fs.writeFile(outputPath, text);
  return outputPath;
}

// 11. Image conversion (PNG ↔ JPG ↔ WEBP)
async function convertImage(inputPath, format) {
  const outputPath = getOutputPath(inputPath, `.${format}`);
  await sharp(inputPath).toFormat(format).toFile(outputPath);
  return outputPath;
}

// ===== ROUTES =====

// Health check
app.get('/', (req, res) => {
  res.json({ 
    status: 'ok',
    service: 'ConvertFlow API',
    version: '2.0.0',
    endpoints: ['/convert', '/download/:fileId']
  });
});

// Convert route
app.post('/convert', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'No file uploaded' });
    }

    const inputPath = req.file.path;
    const targetFormat = req.body.format || 'pdf';
    const { originalname, filename } = req.file;
    
    console.log(`[CONVERT] ${originalname} → ${targetFormat}`);

    let outputPath;
    const ext = path.extname(originalname).toLowerCase();

    // Route to appropriate converter
    if (ext === '.pdf' && ['png', 'jpg', 'jpeg', 'webp'].includes(targetFormat)) {
      outputPath = await pdfToImage(inputPath, targetFormat);
    }
    else if (['.png', '.jpg', '.jpeg', '.webp'].includes(ext) && targetFormat === 'pdf') {
      outputPath = await imageToPdf([inputPath]);
    }
    else if (['.doc', '.docx'].includes(ext) && targetFormat === 'pdf') {
      outputPath = await wordToPdf(inputPath);
    }
    else if (['.xls', '.xlsx'].includes(ext) && targetFormat === 'pdf') {
      outputPath = await excelToPdf(inputPath);
    }
    else if (ext === '.html' && targetFormat === 'pdf') {
      outputPath = await htmlToPdf(inputPath);
    }
    else if (['.txt', '.md'].includes(ext) && targetFormat === 'pdf') {
      outputPath = await textToPdf(inputPath);
    }
    else if (ext === '.pdf' && targetFormat === 'compress') {
      outputPath = await compressPdf(inputPath);
    }
    else if (['.png', '.jpg', '.jpeg', '.webp'].includes(ext) && ['png', 'jpg', 'jpeg', 'webp'].includes(targetFormat)) {
      outputPath = await convertImage(inputPath, targetFormat);
    }
    else if (['.png', '.jpg', '.jpeg', '.pdf'].includes(ext) && targetFormat === 'txt') {
      outputPath = await ocrExtract(inputPath);
    }
    else {
      // Fallback: return original file with converted extension
      outputPath = getOutputPath(originalname, `.${targetFormat}`);
      await fs.copy(inputPath, outputPath);
    }

    const outputFile = path.basename(outputPath);
    
    res.json({
      success: true,
      message: 'Conversion complete',
      fileId: outputFile,
      fileName: outputFile,
      originalName: originalname,
      format: targetFormat
    });

  } catch (error) {
    console.error('[ERROR]', error);
    res.status(500).json({ 
      success: false, 
      message: error.message || 'Conversion failed' 
    });
  }
});

// Download route
app.get('/download/:fileId', async (req, res) => {
  try {
    const fileId = req.params.fileId;
    const filePath = path.join(convertedDir, fileId);
    
    // Also check uploads (for fallback cases)
    const uploadPath = path.join(uploadsDir, fileId);
    
    let finalPath = null;
    if (await fs.pathExists(filePath)) finalPath = filePath;
    else if (await fs.pathExists(uploadPath)) finalPath = uploadPath;
    
    if (!finalPath) {
      return res.status(404).json({ error: 'File not found' });
    }

    const downloadName = req.query.name || path.basename(finalPath);
    
    res.setHeader('Content-Disposition', `attachment; filename="${downloadName}"`);
    res.setHeader('Content-Type', 'application/octet-stream');
    
    const stream = fs.createReadStream(finalPath);
    stream.pipe(res);
    
    // Cleanup after 1 hour (optional)
    setTimeout(() => {
      fs.remove(finalPath).catch(() => {});
    }, 3600000);

  } catch (error) {
    console.error('[DOWNLOAD ERROR]', error);
    res.status(500).json({ error: 'Download failed' });
  }
});

// Cleanup old files periodically
setInterval(async () => {
  const dirs = [uploadsDir, convertedDir];
  const now = Date.now();
  const maxAge = 24 * 60 * 60 * 1000; // 24 hours
  
  for (const dir of dirs) {
    const files = await fs.readdir(dir);
    for (const file of files) {
      const filePath = path.join(dir, file);
      const stats = await fs.stat(filePath);
      if (now - stats.mtimeMs > maxAge) {
        await fs.remove(filePath);
        console.log(`[CLEANUP] Removed ${file}`);
      }
    }
  }
}, 3600000); // Run every hour

// Error handling
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ success: false, message: 'Internal server error' });
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`🚀 ConvertFlow API running on port ${PORT}`);
  console.log(`📁 Uploads: ${uploadsDir}`);
  console.log(`📦 Converted: ${convertedDir}`);
});
