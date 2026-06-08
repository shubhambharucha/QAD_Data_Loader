# QAD Data Loader - Full Setup Guide

## 🚀 Quick Start

### Prerequisites
- Python 3.8+ (for backend)
- Node.js 16+ & npm (for frontend)

---

## 📋 Backend Setup (FastAPI)

### Step 1: Install Python Dependencies

Open a terminal in the project root (`c:\Users\shubham.bharucha\Desktop\Bulk_upload\`) and run:

```bash
pip install fastapi uvicorn python-multipart
```

**Optional:** If you need specific data validation packages, install them:
```bash
pip install openpyxl pandas
```

### Step 2: Start the FastAPI Backend

```bash
python main.py
```

You should see:
```
INFO:     Uvicorn running on http://0.0.0.0:8000
INFO:     Application startup complete
```

✅ **Backend is running at:** `http://localhost:8000`

---

## 🎨 Frontend Setup (React + Vite)

### Step 1: Install Dependencies

In a new terminal, navigate to frontend:
```bash
cd frontend
npm install
```

### Step 2: Start Development Server

```bash
npm run dev
```

You should see:
```
➜  Local:   http://localhost:5173/
➜  Network: use --host to expose
```

✅ **Frontend is running at:** `http://localhost:5173`

---

## 🎯 Using the Application

### Access the App
- Open browser to: **`http://localhost:5173`**

### Pages Available

#### 1. **Data Loader Dashboard** (Default)
- Select entities to validate/load
- Real-time progress updates
- Load summary with results

#### 2. **CSS Awards Showcase**
- Click "View Awards Showcase" button
- Beautiful animations and demonstrations
- Click "Go to Data Loader" to return

---

## 📁 Data Folder Structure

Place your .xlsx files in these folders:
```
Bulk_upload/
├── Supplier/          ← Drop .xlsx files here
├── Customer/          ← Drop .xlsx files here
├── GCM/
├── BR/
├── Customer_Item/
├── ProductionOrder/
├── PurchaseOrder/
├── SalesOrder/
├── Supplier_Item/
└── SupplierPriceList/
```

---

## 🔄 API Endpoints

The backend provides these REST endpoints:

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/api/status` | GET | Get file counts for all entities |
| `/api/validate` | POST | Validate selected entities (SSE stream) |
| `/api/load` | POST | Load selected entities (SSE stream) |

### Example Request:
```bash
curl -X POST http://localhost:8000/api/validate \
  -H "Content-Type: application/json" \
  -d '{"entities": ["Supplier", "Customer"]}'
```

---

## 🛠️ Troubleshooting

### Backend Issues

**Port 8000 already in use:**
```bash
# Find and kill process on port 8000 (Windows)
netstat -ano | findstr :8000
taskkill /PID <PID> /F
```

**Missing validation scripts:**
- Ensure all validate_*.py files exist in `Scripts/` folder
- Check spelling matches entity names

### Frontend Issues

**Port 5173 already in use:**
```bash
npm run dev -- --port 5174
```

**API not connecting:**
- Verify backend is running on `http://localhost:8000`
- Check browser console (F12) for errors
- Ensure CORS is properly configured in main.py

**Dependencies not found:**
```bash
cd frontend
rm -rf node_modules package-lock.json
npm install
```

---

## 📦 Build for Production

### Backend (no build needed)
FastAPI runs directly with Python

### Frontend (create optimized build)
```bash
cd frontend
npm run build
```

Output files in `frontend/dist/` folder

---

## 🔌 Environment Variables (Optional)

Create `.env.local` in frontend folder:
```env
REACT_APP_API_URL=http://localhost:8000
```

This allows easy switching between development and production APIs.

---

## 📊 File Processing Flow

1. **Select** → Choose entities from dashboard
2. **Validate** → Backend validates .xlsx files using Scripts/validate_*.py
3. **Review** → Check validation results (✓ pass or ✗ fail)
4. **Load** → Backend loads validated data using Scripts/*_load.py
5. **Archive** → Successful files moved to Archive/ folder

---

## 🎓 Key Features

✨ **Real-time SSE Streaming** - Live progress updates during validation/loading
🎨 **Beautiful UI** - Tailwind CSS + Framer Motion animations
📊 **Detailed Logging** - See exactly what's happening
🔄 **Error Handling** - Automatic retry prefixing for failed files
📁 **File Archiving** - Successfully loaded files automatically archived

---

## 🤝 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review browser console (F12) for error messages
3. Check backend terminal for server-side errors
4. Verify all required Python packages are installed

Enjoy! 🎉
