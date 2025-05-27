# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

**Docker Compose (Recommended)**:
- `docker-compose up -d` - Start application in background
- `docker-compose down` - Stop application
- `docker-compose logs -f` - View logs
- `docker-compose exec pdf-converter bash` - Access container shell
- `docker-compose build --no-cache` - Rebuild image without cache (if build fails)

**Local Development**:
- `npm install` - Install dependencies
- `npm start` - Start production server (http://localhost:3000)
- `npm run dev` - Start development server with auto-reload using nodemon

## Application Architecture

This is a web-based Office to PDF converter with a simple client-server architecture:

**Backend (server.js)**:
- Express.js server serving static files from `public/` directory
- Uses `multer` for handling file uploads to `uploads/` directory with Japanese filename encoding support
- Converts Office files (.doc, .docx, .ppt, .pptx) to PDF using `libreoffice-convert` library
- Serves converted PDFs from `output/` directory via download endpoint with proper UTF-8 filename handling
- Automatically creates required directories (`uploads/`, `output/`) on startup

**Frontend (public/)**:
- Single-page application with drag-and-drop file selection
- Supports multiple file selection and folder selection (webkitdirectory)
- Real-time progress indication during conversion
- Direct download links for converted PDFs
- HTML escaping for Japanese filenames to prevent display corruption

**File Processing Flow**:
1. Files uploaded via `/convert` endpoint (POST with multipart/form-data)
2. Server validates file types (only .doc, .docx, .ppt, .pptx allowed)
3. LibreOffice converts each file to PDF format
4. Temporary upload files are cleaned up after conversion
5. Converted PDFs available via `/download/:filename` endpoint

## System Requirements

**Docker Deployment (Recommended)**:
- Docker and Docker Compose installed
- All dependencies (LibreOffice, Japanese fonts, Node.js) are included in the container
- Uses node:18-slim base image with LibreOffice and fonts-noto-cjk for Japanese support

**Local Development**:
- LibreOffice must be installed on the system (used by libreoffice-convert for PDF conversion)
- Japanese fonts required for proper PDF text rendering (fonts-noto-cjk, fonts-noto-cjk-extra on Ubuntu/Debian)
- Node.js v14+ required

## Japanese Language Support

This application includes specific handling for Japanese filenames and content:

**Character Encoding**:
- File uploads use latin1 to UTF-8 conversion for proper Japanese filename handling
- Download responses include UTF-8 Content-Disposition headers
- Frontend uses HTML escaping to prevent filename display corruption

**PDF Font Support**:
- Install Japanese fonts: `sudo apt-get install fonts-noto-cjk fonts-noto-cjk-extra`
- Update font cache: `sudo fc-cache -fv`
- LibreOffice will use system fonts for proper Japanese text rendering in PDFs

**Troubleshooting Japanese Text**:
- If PDF content shows garbled text, ensure Japanese fonts are installed on the system
- Original Office documents should use standard Japanese fonts (MS Gothic, MS Mincho, etc.)
- File path in URLs are URL-encoded/decoded to handle Japanese characters safely