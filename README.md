# Office to PDF Converter

Microsoft Office ファイル（.doc, .docx, .ppt, .pptx）をPDFに変換するWebアプリケーションです。

## 機能

- 複数ファイルの同時変換
- ドラッグ&ドロップでのファイル選択
- フォルダ一括選択機能
- 変換後PDFファイルのダウンロード
- レスポンシブデザイン

## セットアップ

### Docker Compose を使用（推奨）

1. Docker Compose でアプリケーションを起動:
```bash
docker-compose up -d
```

初回起動時はイメージのビルドに時間がかかります（LibreOfficeと日本語フォントのインストールのため）。

2. ブラウザで `http://localhost:3000` にアクセス

3. 停止する場合:
```bash
docker-compose down
```

#### トラブルシューティング

ビルドエラーが発生した場合：
```bash
# キャッシュをクリアして再ビルド
docker-compose build --no-cache
docker-compose up -d
```

### ローカル環境での実行

#### 前提条件

- Node.js (v14以上)
- LibreOffice（PDF変換に使用）

#### LibreOfficeのインストール

```bash
# Ubuntu/Debian
sudo apt-get update
sudo apt-get install libreoffice

# 日本語フォントもインストール（文字化け防止）
sudo apt-get install fonts-noto-cjk fonts-noto-cjk-extra

# macOS (Homebrew)
brew install --cask libreoffice

# Windows
# https://www.libreoffice.org/download/download/ からダウンロード
# 日本語フォント（游ゴシック、游明朝、MS ゴシック等）がプリインストール済み
```

#### アプリケーションのセットアップ

1. 依存関係のインストール:
```bash
npm install
```

2. サーバーの起動:
```bash
npm start
```

または開発モード:
```bash
npm run dev
```

3. ブラウザで `http://localhost:3000` にアクセス

## 使用方法

1. Webブラウザでアプリケーションにアクセス
2. ファイルをドラッグ&ドロップするか、「ファイルを選択」「フォルダを選択」ボタンでファイルを選択
3. 「PDFに変換」ボタンをクリック
4. 変換完了後、各ファイルのダウンロードボタンでPDFをダウンロード

## サポートファイル形式

- Microsoft Word: .doc, .docx
- Microsoft PowerPoint: .ppt, .pptx

## 技術スタック

- **バックエンド**: Node.js, Express.js
- **PDF変換**: LibreOffice (libreoffice-convert)
- **フロントエンド**: HTML, CSS, JavaScript
- **ファイルアップロード**: Multer

## ディレクトリ構造

```
document-pdf-converter/
├── server.js              # メインサーバーファイル
├── package.json            # Node.js依存関係
├── public/                 # フロントエンドファイル
│   ├── index.html         # メインHTML
│   ├── style.css          # スタイルシート
│   └── script.js          # JavaScript
├── uploads/               # 一時アップロードフォルダ（自動作成）
└── output/                # 変換済みPDFフォルダ（自動作成）
```