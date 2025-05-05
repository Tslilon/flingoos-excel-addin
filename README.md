# Flingoos Excel Logger Add-in

This Excel Add-in captures detailed semantic events from Excel and sends them to a local listener server for workflow tracing.

## Deployment Steps

### 1. Push to GitHub

```bash
cd excel-addin
git init
git add .
git commit -m "Excel add-in"
gh repo create flingoos-excel --public --source=. --push
```

### 2. Deploy with Vercel

1. Go to https://vercel.com
2. Sign in (GitHub recommended)
3. Import your repo
4. Deploy (it'll give you a URL like https://flingoos-excel.vercel.app)

### 3. Update manifest.xml (if needed)

If your Vercel URL is different from `https://flingoos-excel.vercel.app`, update all URLs in the manifest.xml file.

### 4. Upload to Microsoft for Sideloading

1. Go to https://appsource.microsoft.com
2. Top right → Profile icon → My Add-ins
3. Click Upload My Add-in, choose your updated manifest.xml file

### 5. Use the Add-in

1. Start the Python listener server on your local machine:
   ```bash
   cd ~/code/flingoos-v0
   source .venv/bin/activate
   python src/excel_logger.py
   ```

2. Open Excel and go to the Insert tab > My Add-ins > find the Flingoos Excel Logger

3. Use the add-in to log Excel events:
   - Click "Start Logging" to begin capturing Excel events
   - Perform your Excel tasks
   - Click "Stop Logging" when done

4. Events are logged to your local machine at `data/raw/excel_*.jsonl`

## Hosted Add-in Benefits

This approach provides:
- HTTPS support required by Excel
- Easy access for managed Excel installations
- Simple deployment without self-signed certificates
- Works with admin-managed Excel installations 