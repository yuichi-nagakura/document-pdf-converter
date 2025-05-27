FROM node:18-slim

# 作業ディレクトリを設定
WORKDIR /app

# 必要なパッケージを段階的にインストール
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    wget \
    gnupg \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# LibreOfficeインストール
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

# 日本語フォントをインストール
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    fonts-noto-cjk \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# フォントキャッシュを更新
RUN fc-cache -fv

# package.jsonとpackage-lock.jsonをコピー
COPY package*.json ./

# 依存関係をインストール
RUN npm ci --omit=dev

# アプリケーションファイルをコピー
COPY . .

# 必要なディレクトリを作成
RUN mkdir -p uploads output

# ポートを公開
EXPOSE 3000

# 環境変数を設定
ENV NODE_ENV=production
ENV LANG=ja_JP.UTF-8
ENV LC_ALL=ja_JP.UTF-8

# アプリケーションを実行
CMD ["npm", "start"]