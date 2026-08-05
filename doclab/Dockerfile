FROM debian:13-slim

# ═══════════════════════════════════════════════════════════════
# Atlas Document Lab — Full Excel Development Toolkit
# Supports: VBA (text), Python, TypeScript/Office Scripts, C#/.NET,
#           DAX, Power Query M, SQL, LAMBDA, JavaScript Add-ins
# ═══════════════════════════════════════════════════════════════

RUN apt-get update && apt-get install -y --no-install-recommends \
    # ── Office Suite ──
    libreoffice-core \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    libreoffice-common \
    # ── Document Conversion ──
    pandoc \
    # ── LaTeX ──
    texlive-latex-base \
    texlive-latex-extra \
    texlive-latex-recommended \
    texlive-fonts-recommended \
    texlive-extra-utils \
    texlive-lang-european \
    # ── PDF Tools ──
    poppler-utils \
    qpdf \
    pdftk \
    # ── Editors ──
    vim \
    notepadqq \
    # ── Database ──
    sqlite3 \
    mdbtools \
    # ── TypeScript / Node.js (Office Scripts) ──
    nodejs \
    npm \
    # ── C# / .NET (Excel Add-in development) ──
    mono-complete \
    # ── XML inspection (xlsx internals) ──
    xmlstarlet \
    jq \
    # ── Python ──
    python3 \
    python3-pip \
    python3-venv \
    # ── Network / Build ──
    curl \
    ca-certificates \
    git \
    unzip \
    && rm -rf /var/lib/apt/lists/*

# ── Python Libraries for Excel & Data Science ──
RUN pip3 install --no-cache-dir --break-system-packages \
    openpyxl \
    xlsxwriter \
    xlrd \
    pandas \
    matplotlib \
    seaborn \
    python-docx \
    python-pptx \
    reportlab \
    psycopg2-binary \
    pymysql \
    xlwings

# ── TypeScript / Office Scripts support ──
RUN npm install -g typescript \
    && npm install -g @types/node

WORKDIR /workspace

ENTRYPOINT ["sleep", "infinity"]
