# Use fixed python version
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Copy project files
COPY . /app

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Create required folders
RUN mkdir uploads output

# Expose port
EXPOSE 10000

# Start application
CMD ["gunicorn", "-b", "0.0.0.0:10000", "app:app"]
