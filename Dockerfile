FROM python:3.11-slim

# Tối ưu build
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Cài deps hệ thống đủ cho lxml/python-docx/openpyxl
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libxml2-dev \
    libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements trước để layer cache
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy mã nguồn và data
COPY app.py ./
COPY .streamlit ./\.streamlit
COPY data ./data

# Vercel sẽ cung cấp PORT, ta run Streamlit bind 0.0.0.0:$PORT
CMD ["bash", "-lc", "streamlit run app.py --server.address 0.0.0.0 --server.port ${PORT:-8501} --browser.gatherUsageStats false"]
