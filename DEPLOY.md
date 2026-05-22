# Deployment Instructions

## Step 1 — Push to GitHub

1. Create a new repository at https://github.com/new (public or private).
2. In your terminal, from the project folder:

```bash
git init
git add agent.py requirements.txt
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

## Step 2 — Deploy on Streamlit Cloud

1. Go to https://streamlit.io/cloud and sign in with your GitHub account.
2. Click **"New app"**.
3. Select your repository, branch (`main`), and set the main file path to `agent.py`.
4. Click **"Deploy"**.

## Step 3 — Add your Groq API key as a secret

1. Once deployed, open your app's **Settings** (three-dot menu → Settings).
2. Go to the **Secrets** tab.
3. Add the following:

```toml
GROQ_API_KEY = "your_groq_api_key_here"
```

4. Click **Save**. The app will automatically restart.

> Get your free Groq API key at https://console.groq.com

## Step 4 — Share

Copy the app URL (e.g. `https://your-app-name.streamlit.app`) and share it.
