# =============================================================================
# Sales Pivot Report Generator - Dockerfile (Single-stage for Portainer)
# Technology: Python 3.11 + Streamlit
# =============================================================================

FROM python:3.11-slim

# Labels
LABEL maintainer="Sales Report Team"
LABEL description="Sales Pivot Report Generator - Streamlit Web App"
LABEL version="1.0.0"

# Environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONFAULTHANDLER=1 \
    STREAMLIT_SERVER_PORT=8501 \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
    # Make Streamlit work nicely in containers
    STREAMLIT_SERVER_ENABLECORS=false \
    STREAMLIT_SERVER_ENABLEXSRSFPROTECTION=false

WORKDIR /app

# Install system dependencies (curl used by HEALTHCHECK)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

# Copy and install Python dependencies
COPY requirements-docker.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements-docker.txt

# Copy application code
COPY app.py .
COPY process.py .
COPY cli.py .

# Create non-root user
RUN useradd --create-home appuser && chown -R appuser:appuser /app
USER appuser

# Expose port
EXPOSE 8501

# Health check (Streamlit built-in health endpoint)
HEALTHCHECK --interval=30s --timeout=10s --start-period=20s --retries=3 \
    CMD curl --fail http://127.0.0.1:8501/_stcore/health || exit 1

# IMPORTANT:
# Use CMD (not ENTRYPOINT) so Portainer "Command override" won't accidentally break startup.
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
