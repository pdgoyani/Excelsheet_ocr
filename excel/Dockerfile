FROM python:3.10-slim

# Install system dependencies (Tesseract OCR & Poppler for PDF-to-image)
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all project files into the container
COPY . .

# Expose Streamlit's default port
EXPOSE 8501

# Run Streamlit, utilizing Render's dynamic PORT environment variable
CMD ["sh", "-c", "streamlit run app.py --server.port=${PORT:-8501} --server.address=0.0.0.0"]
