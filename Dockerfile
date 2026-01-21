# =============================================================================
# Sales Pivot Report Generator - Dockerfile (Portainer URL build from GitHub tar.gz)
# =============================================================================
FROM python:3.11-slim

LABEL maintainer="Sales Report Team"
LABEL description="Sales Pivot Report Generator - Streamlit Web App"
LABEL version="1.0.0"

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONFAULTHANDLER=1 \
    STREAMLIT_SERVER_PORT=8501 \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
    STREAMLIT_SERVER_ENABLECORS=false \
    STREAMLIT_SERVER_ENABLEXSRSFPROTECTION=false

# Portainer URL build drops the extracted archive into the build context.
# We copy everything to /build, then "cd" into the extracted repo folder at runtime.
WORKDIR /build

# System deps (curl used by HEALTHCHECK)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy the entire build context (GitHub tar.gz extracted contents)
COPY . /build

# Install Python dependencies from the extracted repo folder.
# This finds the first directory that contains requirements-docker.txt.
RUN set -eux; \
    REPO_DIR="$(find /build -maxdepth 2 -type f -name requirements-docker.txt -printf '%h\n' | head -n 1)"; \
    echo "Detected repo dir: ${REPO_DIR}"; \
    python -m pip install --no-cache-dir --upgrade pip; \
    python -m pip install --no-cache-dir -r "${REPO_DIR}/requirements-docker.txt"; \
    echo "${REPO_DIR}" > /build/_repo_dir.txt

EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=10s --start-period=20s --retries=3 \
  CMD curl --fail http://127.0.0.1:8501/_stcore/health || exit 1

# Start Streamlit from the detected repo folder (where app.py lives)
CMD ["sh", "-c", "cd \"$(cat /build/_repo_dir.txt)\" && exec streamlit run app.py --server.address=0.0.0.0 --server.port=8501"]
