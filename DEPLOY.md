# HR Guru ATS — Deployment Guide
# Deploy to Render.com (free) in ~15 minutes, no command line needed

=====================================================================
WHAT YOU NEED (all free)
=====================================================================
1. GitHub account       → github.com
2. Render account       → render.com  (sign in with GitHub)
3. Cloudinary account   → cloudinary.com  (for CV file storage)

=====================================================================
STEP 1 — Put the code on GitHub (5 min)
=====================================================================
1. Go to github.com → click "+" → "New repository"
2. Name it: hrguru-ats
3. Keep it Private
4. Click "Create repository"
5. On the next page, click "uploading an existing file"
6. Drag and drop ALL files from this ats4 folder:
      app.py
      requirements.txt
      Procfile
      render.yaml
      templates/  (the whole folder)
7. Click "Commit changes"

=====================================================================
STEP 2 — Get your Cloudinary URL (3 min)
=====================================================================
(Cloudinary stores CV files since Render's free disk resets)

1. Go to cloudinary.com → Sign Up Free
2. After login, go to Dashboard
3. Copy the "API Environment variable" — it looks like:
      CLOUDINARY_URL=cloudinary://123456789:abc123@yourcloudname
4. Save this somewhere — you'll need it in Step 3

=====================================================================
STEP 3 — Deploy on Render (7 min)
=====================================================================
1. Go to render.com → Log in with GitHub
2. Click "New +" → "Blueprint"
3. Select your hrguru-ats repository
4. Render will read render.yaml and auto-create:
      • A web service (your app)
      • A free PostgreSQL database
5. You'll see an environment variable "CLOUDINARY_URL" needs a value
   → Paste the URL you copied in Step 2
6. Click "Apply"
7. Wait 3-4 minutes for deploy to finish

=====================================================================
STEP 4 — Initialize the database (1 min)
=====================================================================
After deploy completes:
1. In Render dashboard → click your "hrguru-ats" web service
2. Click "Shell" tab (top right)
3. Type: python -c "from app import init_db; init_db()"
4. Press Enter — you'll see no errors if it worked

=====================================================================
DONE — Your URLs
=====================================================================
Admin dashboard:  https://hrguru-ats.onrender.com
Recruiter upload: https://hrguru-ats.onrender.com/upload

Share the /upload URL with your team.
Bookmark the main URL for yourself.

=====================================================================
IMPORTANT — Free tier behaviour
=====================================================================
Render's free web service "sleeps" after 15 minutes of inactivity.
The first visit after sleep takes ~30 seconds to wake up.
The database and all data are permanent — nothing is lost.

To avoid the sleep delay, you can:
• Open the dashboard once in the morning before your team starts
• OR upgrade to Render's $7/month "Starter" plan (always on)

=====================================================================
UPDATING THE APP IN FUTURE
=====================================================================
To deploy any changes:
1. Go to your GitHub repo
2. Click the file you want to change → click the pencil icon
3. Edit and commit
4. Render auto-deploys within 2 minutes

=====================================================================
EXPORT (NEW FEATURE)
=====================================================================
On the Candidates page, two buttons appear at the top right:
• "↓ CSV"   — exports current filtered view as CSV
• "↓ Excel" — exports current filtered view as formatted Excel

The export respects all active filters (role, location, notice, search).
So to export "all Python developers in Bangalore" — just set those
filters and click export.
