# Use Python base image
FROM python:3

# Set working directory
WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY . .

# Expose port (Cloud Run expects $PORT)
EXPOSE 8080

# Run Streamlit
CMD streamlit run app.py --server.port=8080 --server.address=0.0.0.0
