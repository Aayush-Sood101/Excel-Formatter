# Excel Duplicate Remover Web App

A Flask web application that removes consecutive duplicate rows from Excel files based on Column B (Author) and Column C (Title).

## Features

- 🚀 Web-based interface - no installation needed for users
- 📊 Processes Excel files (.xlsx, .xls)
- 🔍 Identifies duplicates based on Column B and Column C only
- ✅ Preserves all other columns and data
- 📥 Automatic download of cleaned file
- 🎨 Beautiful, responsive UI with drag-and-drop support
- 📈 Shows statistics (rows removed, before/after counts)

## Local Development

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python app.py
```

3. Open your browser and go to:
```
http://localhost:5000
```

## How to Deploy on Render (FREE)

### Step 1: Prepare Your Code

1. Make sure all files are in your project folder:
   - `app.py`
   - `requirements.txt`
   - `render.yaml`
   - `templates/index.html`

### Step 2: Create a GitHub Repository

1. Go to [GitHub](https://github.com) and create a new repository
2. Initialize git in your project folder:
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git push -u origin main
```

### Step 3: Deploy on Render

1. Go to [Render](https://render.com) and sign up/login (you can use GitHub to sign in)
2. Click **"New +"** → **"Web Service"**
3. Connect your GitHub repository
4. Render will auto-detect the `render.yaml` file
5. Click **"Create Web Service"**
6. Wait for deployment (usually takes 2-3 minutes)
7. Your app will be live at: `https://your-app-name.onrender.com`

### Important Notes for Render:

- ✅ **FREE tier** includes:
  - 750 hours/month (enough for continuous running)
  - Automatic HTTPS
  - Custom domains supported
  
- ⚠️ **Free tier limitations**:
  - App spins down after 15 minutes of inactivity
  - First request after inactivity may take 30-60 seconds
  - 512 MB RAM limit

### Step 4: Configure (Optional)

You can customize in the Render dashboard:
- Environment variables
- Custom domain
- Auto-deploy on git push
- Health check endpoints

## Usage

1. Visit your deployed website
2. Click "Choose a file" or drag-and-drop an Excel file
3. Click "Process & Download"
4. The cleaned file will automatically download

## File Structure

```
NEW PROJ/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── render.yaml           # Render deployment configuration
├── README.md             # This file
├── templates/
│   └── index.html        # Web interface
└── remove_duplicates.py  # Original standalone script (optional)
```

## Technical Details

- **Framework**: Flask 3.0
- **Excel Processing**: pandas + openpyxl
- **Server**: Gunicorn (production)
- **Max File Size**: 16MB
- **Supported Formats**: .xlsx, .xls

## Security Notes

- Change the `secret_key` in `app.py` before deployment
- The app doesn't store any uploaded files
- All processing is done in memory

## Support

If you encounter issues:
- Check Render logs in the dashboard
- Ensure your Excel file has at least 3 columns
- File size must be under 16MB

## License

Free to use and modify.
