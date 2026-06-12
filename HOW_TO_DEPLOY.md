# How to Deploy GEMS PE Lesson Generator
## Live link for everyone — free, no login needed

---

## Step 1 — Create a GitHub account (if you don't have one)
Go to https://github.com and sign up. It's free.

---

## Step 2 — Upload the project to GitHub

1. Go to https://github.com/new
2. Repository name: `gems-lesson-generator`
3. Set to **Public**
4. Click **Create repository**
5. Click **uploading an existing file**
6. Upload ALL files from this folder:
   - `app.py`
   - `requirements.txt`
   - `render.yaml`
   - `static/index.html` (create a `static` folder first)
7. Click **Commit changes**

---

## Step 3 — Create a free Render account

1. Go to https://render.com
2. Sign up with your GitHub account

---

## Step 4 — Deploy on Render

1. Click **New** → **Web Service**
2. Connect your GitHub repo `gems-lesson-generator`
3. Render will auto-detect the settings from `render.yaml`
4. Scroll down to **Environment Variables**
5. Add:
   - Key: `ANTHROPIC_API_KEY`
   - Value: your API key from https://platform.anthropic.com
6. Click **Create Web Service**

Wait ~2 minutes. Render will give you a link like:
**https://gems-lesson-generator.onrender.com**

---

## Step 5 — Share the link!

Send that link to anyone. They open it in any browser, no login, no setup.

---

## Cost
- Render free tier: $0/month
- Anthropic API: ~$0.003 per lesson generated
- 100 lessons ≈ $0.30

---

## Notes
- Free Render services sleep after 15 min of inactivity — first load may take 30 seconds to wake up
- Upgrade to Render Starter ($7/month) for always-on if needed
