const express = require('express');
const multer = require('multer');
const libre = require('libreoffice-convert');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static('public'));

const upload = multer({
  dest: 'uploads/',
  fileFilter: (req, file, cb) => {
    // ファイル名の文字エンコーディングを修正
    if (file.originalname) {
      try {
        // Buffer.from()でエンコーディングを適切に処理
        file.originalname = Buffer.from(file.originalname, 'latin1').toString('utf8');
      } catch (err) {
        // フォールバック: そのまま使用
      }
    }
    
    const allowedTypes = ['.doc', '.docx', '.ppt', '.pptx'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Unsupported file type'), false);
    }
  }
});

if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}
if (!fs.existsSync('output')) {
  fs.mkdirSync('output');
}

app.post('/convert', upload.array('files'), async (req, res) => {
  try {
    const convertedFiles = [];

    for (const file of req.files) {
      const inputPath = file.path;
      
      // ファイル名を安全に処理
      const baseName = path.parse(file.originalname).name;
      const outputFileName = baseName + '.pdf';
      const outputPath = path.join('output', outputFileName);

      const docxBuf = fs.readFileSync(inputPath);
      
      const pdfBuf = await new Promise((resolve, reject) => {
        libre.convert(docxBuf, '.pdf', undefined, (err, done) => {
          if (err) {
            reject(err);
          } else {
            resolve(done);
          }
        });
      });

      fs.writeFileSync(outputPath, pdfBuf);
      
      convertedFiles.push({
        originalName: file.originalname,
        pdfFileName: outputFileName,
        downloadUrl: `/download/${encodeURIComponent(outputFileName)}`
      });

      fs.unlinkSync(inputPath);
    }

    res.json({ 
      success: true, 
      message: `${convertedFiles.length} files converted successfully`,
      files: convertedFiles 
    });

  } catch (error) {
    console.error('Conversion error:', error);
    res.status(500).json({ 
      success: false, 
      message: 'Conversion failed: ' + error.message 
    });
  }
});

app.get('/download/:filename', (req, res) => {
  const filename = decodeURIComponent(req.params.filename);
  const filePath = path.join('output', filename);
  
  if (fs.existsSync(filePath)) {
    // UTF-8エンコードでファイル名を設定
    const encodedFilename = encodeURIComponent(filename);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodedFilename}`);
    res.download(filePath, filename, (err) => {
      if (err) {
        console.error('Download error:', err);
        res.status(500).send('Download failed');
      }
    });
  } else {
    res.status(404).send('File not found');
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});