# 文字化け対策

## PDF内の日本語文字化けの解決方法

### 1. 日本語フォントのインストール

```bash
# Ubuntu/Debian
sudo apt-get install fonts-noto-cjk fonts-noto-cjk-extra fonts-takao-gothic fonts-takao-mincho

# CentOS/RHEL
sudo yum install google-noto-cjk-fonts ipa-gothic-fonts ipa-mincho-fonts

# macOS
# 通常は日本語フォントがプリインストール済み
```

### 2. LibreOfficeフォント設定の確認

LibreOfficeを起動して以下を確認：
1. ツール → オプション → フォント
2. 「常に次のフォントを使用」にチェック
3. 標準フォントを日本語フォント（Noto Sans CJK JP等）に設定

### 3. システムフォントキャッシュの更新

```bash
sudo fc-cache -fv
```

### 4. 環境変数の設定（必要に応じて）

```bash
export LANG=ja_JP.UTF-8
export LC_ALL=ja_JP.UTF-8
```

### 5. 文書側での対策

- 元のOfficeファイルでフォントを明示的に日本語フォント（MS Gothic、MS Mincho、游ゴシック等）に設定
- PDF変換前にフォントが正しく設定されていることを確認

## Excel書式保持の改善

### テーブル枠線や図表が欠落する場合

1. **元ファイルの確認**
   - セルの枠線が明示的に設定されているか確認
   - 図表やグラフが埋め込まれているか確認

2. **LibreOffice設定の最適化**
   アプリケーションでは以下のオプションを使用しています：
   ```
   --calc --export-to-pdf
   --export-pdf-form-fields=true
   --export-pdf-links=true
   --export-pdf-bookmarks=true
   --export-pdf-notes=true
   --export-pdf-tagged=true
   ```

3. **代替手法**
   - 元のExcelファイルで「印刷プレビュー」を確認
   - セルの書式→枠線を明示的に設定
   - 図表は「図として貼り付け」を使用

## 確認方法

変換後のPDFファイルをPDFビューアで開き、以下を確認してください：
- 日本語テキストが正しく表示される
- テーブルの枠線が保持されている
- 図表やグラフが表示されている