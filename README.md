# MIS HTML to Excel Service (FastAPI)

A stateless, high-performance FastAPI service that transforms MIS HTML reports into styled multi-sheet Excel workbooks with VLOOKUP integration.

---

## ⚡ Zero-Disk Storage Architecture

Cloud storage (EBS volumes, persistent disks, S3 egress) can be expensive. This service operates with **zero persistent disk storage**:

```
[Client / Browser]
       │
       │  1. Multipart Form Upload (HTML files)
       ▼
 [FastAPI App]
       │
       │  2. Read directly into RAM bytes (UploadFile.read())
       ▼
[BeautifulSoup & Pandas]
       │
       │  3. Parse HTML tables & perform VLOOKUP in memory
       ▼
  [OpenPyXL]
       │
       │  4. Write workbook directly into io.BytesIO() buffer
       ▼
[StreamingResponse]
       │
       │  5. Stream .xlsx bytes back as HTTP download attachment
       ▼
[Client / Browser saves file]
```

- **0 bytes** written to server disk or `/tmp`.
- Memory is released immediately by Python's garbage collector.
- Runs smoothly even on the smallest **512 MB RAM** free-tier instances.

---

## 🚀 Running Locally

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Start the Server
```bash
python main.py
# Or with uvicorn directly:
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

### 3. Open in Browser
Visit **`http://localhost:8000`** to access the web interface:
- Drag and drop **HTML 1** (Ranking list)
- Drag and drop **HTML 2** (Period stats)
- Click **"Convert & Download Excel"** to instantly receive `report_with_lookup.xlsx`.

---

## 📡 API Endpoints

### 1. Web Portal
- **`GET /`**: Interactive UI for drag-and-drop file processing and direct browser download.

### 2. Convert and Stream Excel
- **`POST /api/convert`**
  - **Parameters (multipart/form-data)**:
    - `file1`: Ranking list HTML file (`html1.htm`)
    - `file2`: Period stats HTML file (`html2.htm`)
    - `table_index` *(optional, default: 4)*: Index of data table in the HTML.
  - **Response**: Streams binary `.xlsx` with `Content-Disposition: attachment`.
  - **Custom Headers**:
    - `X-Stats-Matched`: Number of agents matched.
    - `X-Stats-Unmatched`: Number of agents unmatched.
    - `X-Stats-Total-Agents`: Total agents in file 1.

**cURL Example**:
```bash
curl -X POST "http://localhost:8000/api/convert" \
  -F "file1=@html1.htm" \
  -F "file2=@html2.htm" \
  -F "table_index=4" \
  --output report_with_lookup.xlsx
```

### 3. Summary Preview (JSON)
- **`POST /api/summary`**: Returns metrics (matched/unmatched/policy totals) as JSON without generating the full download.

### 4. Health Check
- **`GET /health`**: Returns `{"status": "healthy"}` for container health monitoring.

---

## ☁️ Cloud Deployment Options (Zero Storage Cost)

### Option A: Free / Cheap PaaS (Render, Railway, Fly.io)
Because the app is stateless and writes nothing to disk, you can deploy it on free or $5/mo tiers without needing any attached storage volume.

1. **Render.com**:
   - Create a new **Web Service**.
   - Connect your GitHub repository.
   - Set **Build Command**: `pip install -r requirements.txt`
   - Set **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT --workers 2`

2. **Railway.app**:
   - Push code to GitHub.
   - Deploy new project directly from repo (detected via `Procfile` / `Dockerfile`).

3. **Google Cloud Run / AWS App Runner**:
   - Completely serverless (scales to zero when not in use).
   - Zero storage bill.

---

### Option B: Docker / Docker Compose

Build and run the containerized service:
```bash
docker compose up -d --build
```
Access at `http://<your-server-ip>:8000`.

---

### Option C: Linux VPS (Ubuntu/Debian) with Systemd

1. Copy the project files to `/opt/mis-excel-service`.
2. Create a virtual environment:
   ```bash
   python3 -m venv /opt/mis-excel-service/venv
   /opt/mis-excel-service/venv/bin/pip install -r /opt/mis-excel-service/requirements.txt
   ```
3. Create a systemd service file `/etc/systemd/system/mis-excel.service`:
   ```ini
   [Unit]
   Description=MIS HTML to Excel Service
   After=network.target

   [Service]
   User=www-data
   WorkingDirectory=/opt/mis-excel-service
   ExecStart=/opt/mis-excel-service/venv/bin/uvicorn main:app --host 0.0.0.0 --port 8000 --workers 2
   Restart=always

   [Install]
   WantedBy=multi-user.target
   ```
4. Start and enable service:
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable --now mis-excel
   ```
# misData
