# Law School Scholarship Visualizer

## Local Development

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Place your data file at:
   ```
   E:\Law School Application\Scholarship data\ALL_SCHOOLS_COMBINED.csv
   ```
   (or update the path in `law_school_dashboard.py`)

3. Run:
   ```
   python law_school_dashboard.py
   ```

## Deploying to Railway

1. Create a GitHub repository
2. Add these files to the repo:
   - `law_school_dashboard.py`
   - `requirements.txt`
   - `Procfile`
   - `data/ALL_SCHOOLS_COMBINED.csv`  ← put your CSV here
3. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub
4. Select your repo — Railway auto-detects and deploys
5. Your app will be live at a `.railway.app` URL

## Deploying to Render

1. Same GitHub setup as above
2. Go to [render.com](https://render.com) → New → Web Service
3. Connect your GitHub repo
4. Set:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn law_school_dashboard:server --bind 0.0.0.0:$PORT --workers 1 --timeout 120`
5. Deploy

## File Structure

```
your-repo/
├── law_school_dashboard.py
├── requirements.txt
├── Procfile
├── .gitignore
├── README.md
└── data/
    └── ALL_SCHOOLS_COMBINED.csv   ← your data file goes here
```

## Notes

- The `data/` folder is gitignored for `.xlsx` files but not `.csv`
- The app automatically finds the CSV in `data/` when deployed
- Falls back to your local path when running locally
