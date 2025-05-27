const express = require('express');
const multer = require('multer');
const libre = require('libreoffice-convert');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const XLSX = require('xlsx');
const PDFMerger = require('pdf-merger-js');
const archiver = require('archiver');

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
    
    const allowedTypes = ['.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx'];
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

// Excel処理用のヘルパー関数
async function processExcelFile(file, convertedFiles) {
  const inputPath = file.path;
  const baseName = path.parse(file.originalname).name;
  const ext = path.extname(file.originalname).toLowerCase();

  if (ext === '.xls' || ext === '.xlsx') {
    // Excelファイルを読み込んでシート情報を取得
    const workbook = XLSX.readFile(inputPath);
    const sheetNames = workbook.SheetNames;

    // 複数シートがある場合は各シートを処理して結合
    if (sheetNames.length > 1) {
      const outputFileName = baseName + '.pdf';
      const outputPath = path.join('output', outputFileName);
      const tempPdfPaths = [];

      // 各シートを個別のPDFに変換
      for (let i = 0; i < sheetNames.length; i++) {
        const sheetName = sheetNames[i];
        
        // 見出しシートを作成
        const headerWorkbook = XLSX.utils.book_new();
        const headerData = [
          [`--- ${sheetName} ---`], // シート名を見出しとして表示
          [''], // 空行
        ];
        const headerSheet = XLSX.utils.aoa_to_sheet(headerData);
        
        // 見出しのスタイルを設定（可能な範囲で）
        headerSheet['A1'].s = {
          font: { bold: true, sz: 16 },
          alignment: { horizontal: 'center' }
        };
        
        XLSX.utils.book_append_sheet(headerWorkbook, headerSheet, 'Header');

        // 実際のシートデータを追加（書式も保持）
        const worksheet = workbook.Sheets[sheetName];
        
        // シートのセル範囲を取得
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        // セルスタイルを保持してコピー
        const styledWorksheet = {};
        for (let row = range.s.r; row <= range.e.r; row++) {
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            if (worksheet[cellAddress]) {
              styledWorksheet[cellAddress] = { ...worksheet[cellAddress] };
              
              // セルに枠線スタイルを追加（可能な範囲で）
              if (!styledWorksheet[cellAddress].s) {
                styledWorksheet[cellAddress].s = {};
              }
              if (!styledWorksheet[cellAddress].s.border) {
                styledWorksheet[cellAddress].s.border = {
                  top: { style: 'thin', color: { rgb: '000000' } },
                  bottom: { style: 'thin', color: { rgb: '000000' } },
                  left: { style: 'thin', color: { rgb: '000000' } },
                  right: { style: 'thin', color: { rgb: '000000' } }
                };
              }
            }
          }
        }
        
        // レンジ情報をコピー
        if (worksheet['!ref']) {
          styledWorksheet['!ref'] = worksheet['!ref'];
        }
        if (worksheet['!merges']) {
          styledWorksheet['!merges'] = worksheet['!merges'];
        }
        if (worksheet['!cols']) {
          styledWorksheet['!cols'] = worksheet['!cols'];
        }
        if (worksheet['!rows']) {
          styledWorksheet['!rows'] = worksheet['!rows'];
        }
        
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, headerSheet, 'Header');
        XLSX.utils.book_append_sheet(newWorkbook, styledWorksheet, sheetName);

        // 一時的なExcelファイルを作成
        const tempExcelPath = inputPath + `_sheet_${i}.xlsx`;
        XLSX.writeFile(newWorkbook, tempExcelPath);

        // LibreOfficeでPDFに変換
        const excelBuf = fs.readFileSync(tempExcelPath);
        const pdfBuf = await new Promise((resolve, reject) => {
          libre.convert(excelBuf, '.pdf', undefined, (err, done) => {
            if (err) {
              reject(err);
            } else {
              resolve(done);
            }
          });
        });

        // 一時PDFファイルを保存
        const tempPdfPath = path.join('output', `temp_${baseName}_${i}.pdf`);
        fs.writeFileSync(tempPdfPath, pdfBuf);
        tempPdfPaths.push(tempPdfPath);
        
        fs.unlinkSync(tempExcelPath); // 一時Excelファイルを削除
      }

      // 複数のPDFを結合
      const merger = new PDFMerger();
      for (const pdfPath of tempPdfPaths) {
        await merger.add(pdfPath);
      }
      
      await merger.save(outputPath);

      // 一時PDFファイルを削除
      tempPdfPaths.forEach(tempPath => {
        if (fs.existsSync(tempPath)) {
          fs.unlinkSync(tempPath);
        }
      });

      convertedFiles.push({
        originalName: file.originalname,
        pdfFileName: outputFileName,
        downloadUrl: `/download/${encodeURIComponent(outputFileName)}`,
        sheetCount: sheetNames.length
      });
    } else {
      // シートが1つの場合は通常処理
      const outputFileName = baseName + '.pdf';
      const outputPath = path.join('output', outputFileName);

      const excelBuf = fs.readFileSync(inputPath);
      const pdfBuf = await new Promise((resolve, reject) => {
        libre.convert(excelBuf, '.pdf', undefined, (err, done) => {
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
    }
  } else {
    // Word/PowerPointファイルの処理
    const outputFileName = baseName + '.pdf';
    const outputPath = path.join('output', outputFileName);

    const docBuf = fs.readFileSync(inputPath);
    const pdfBuf = await new Promise((resolve, reject) => {
      libre.convert(docBuf, '.pdf', undefined, (err, done) => {
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
  }
}

app.post('/convert', upload.array('files'), async (req, res) => {
  try {
    const convertedFiles = [];

    for (const file of req.files) {
      await processExcelFile(file, convertedFiles);
      fs.unlinkSync(file.path); // アップロードファイルを削除
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

// ZIPダウンロード用エンドポイント
app.post('/download-zip', express.json(), (req, res) => {
  const { fileNames } = req.body;
  
  if (!fileNames || !Array.isArray(fileNames) || fileNames.length === 0) {
    return res.status(400).json({ error: 'ファイル名のリストが必要です' });
  }

  const zipFileName = `converted_files_${Date.now()}.zip`;
  const zipPath = path.join('output', zipFileName);

  // ZIPファイルを作成
  const output = fs.createWriteStream(zipPath);
  const archive = archiver('zip', {
    zlib: { level: 9 } // 最高圧縮レベル
  });

  output.on('close', () => {
    console.log(`ZIP作成完了: ${archive.pointer()} bytes`);
    
    // ZIPファイルをダウンロード用に送信
    const encodedZipName = encodeURIComponent(zipFileName);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodedZipName}`);
    res.download(zipPath, zipFileName, (err) => {
      if (err) {
        console.error('ZIP download error:', err);
        res.status(500).send('ZIP download failed');
      } else {
        // ダウンロード完了後にZIPファイルを削除
        setTimeout(() => {
          if (fs.existsSync(zipPath)) {
            fs.unlinkSync(zipPath);
          }
        }, 5000);
      }
    });
  });

  archive.on('error', (err) => {
    console.error('ZIP creation error:', err);
    res.status(500).json({ error: 'ZIP作成に失敗しました' });
  });

  archive.pipe(output);

  // 指定されたファイルをZIPに追加
  let filesAdded = 0;
  for (const fileName of fileNames) {
    const filePath = path.join('output', fileName);
    if (fs.existsSync(filePath)) {
      archive.file(filePath, { name: fileName });
      filesAdded++;
    }
  }

  if (filesAdded === 0) {
    return res.status(404).json({ error: '指定されたファイルが見つかりません' });
  }

  archive.finalize();
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});