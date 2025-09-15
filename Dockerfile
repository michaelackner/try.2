FROM python:3.11-slim

WORKDIR /app

# System deps (optional: for faster openpyxl)
RUN pip install --no-cache-dir --upgrade pip

# Install Python deps first for better layer caching
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy app source
COPY server.py index.html app.js styles.css ./

# Runtime configuration
ENV PORT=8000
EXPOSE 8000

# Start the FastAPI app with Uvicorn
CMD ["sh", "-c", "uvicorn server:app --host 0.0.0.0 --port ${PORT:-8000}"]

