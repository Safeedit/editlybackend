# Use an official Python base image
FROM python:3.11-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    tesseract-ocr \
    ghostscript \
    && apt-get clean

# Set working directory
WORKDIR /app

# Copy all project files into the container
COPY . .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose port (used by gunicorn and Flask)
EXPOSE 10000

# Start the Flask server with gunicorn
CMD ["gunicorn", "server.main:app", "--bind", "0.0.0.0:10000"]
