FROM python:3.11-slim

LABEL org.opencontainers.image.title="EQEmu GM Dashboard"
LABEL org.opencontainers.image.description="GM activity dashboard and reporting server for EQEmu servers"
LABEL org.opencontainers.image.source="https://github.com/ObiwonKenobi/eqemu-gm-dashboard"

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY reports_server.py .
COPY gm_excel.py .

# Reports are served from /reports (mount your host path here)
RUN mkdir -p /reports

EXPOSE 8765

CMD ["python", "/app/reports_server.py"]
