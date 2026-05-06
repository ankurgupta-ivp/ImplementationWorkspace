# PriceMaster Implementation Hub

A React + Supabase web application for managing PriceMaster implementation projects — tracking tasks, questionnaires, estimates, and RAID logs across the full project lifecycle.

## Tech Stack
- **Frontend:** React 18 (Create React App)
- **Database:** Supabase (PostgreSQL)
- **Hosting:** GitHub Pages / Netlify / Vercel (your choice)

## Project Structure
```
src/
  lib/
    supabase.js     # Supabase client
    db.js           # All database operations
    defaults.js     # Default templates & task data
  hooks/
    useApp.js       # Global state context
  components/
    AppShell.jsx    # Layout: topbar + sidebar
    UI.jsx          # Shared UI components
  pages/
    Overview.jsx    # Project overview & KPIs
    Kickoff.jsx     # Kickoff questionnaire
    Tasks.jsx       # Tasks & checklist
    RaidLog.jsx     # RAID log
    Estimator.jsx   # Effort estimator
    OtherPages.jsx  # Dashboard, Templates, Docs, DataSources
  App.js            # Router
  index.js          # Entry point
supabase_schema.sql # Run this in Supabase SQL Editor first
```

## Quick Start

### 1. Supabase Setup
1. Create a free account at [supabase.com](https://supabase.com)
2. Create a new project
3. Go to **SQL Editor** and run the entire contents of `supabase_schema.sql`
4. Go to **Settings → API** and copy your Project URL and anon key

### 2. Local Development
```bash
# Clone the repo
git clone https://github.com/YOUR_ORG/pricemaster-hub.git
cd pricemaster-hub

# Install dependencies
npm install

# Create environment file
cp .env.example .env.local
# Edit .env.local with your Supabase credentials

# Start dev server
npm start
```

### 3. Deploy to GitHub Pages
```bash
npm install --save-dev gh-pages

# Add to package.json:
# "homepage": "https://YOUR_ORG.github.io/pricemaster-hub"
# "scripts": { "predeploy": "npm run build", "deploy": "gh-pages -d build" }

npm run deploy
```

> **Note:** For GitHub Pages, set environment variables via a `.env.production` file or use Netlify/Vercel for easier env var management.

## Environment Variables
```
REACT_APP_SUPABASE_URL=https://YOUR_PROJECT_ID.supabase.co
REACT_APP_SUPABASE_ANON_KEY=YOUR_ANON_KEY
```

## Features
- **Multi-project portfolio** — create and switch between client implementations
- **Kickoff Questionnaire** — 14 sections, 60+ questions, answers saved per project
- **Effort Estimator** — step-by-step model with risk adjustments
- **Tasks & Checklist** — 80+ pre-loaded tasks, filterable, exportable to Excel
- **RAID Log** — Risks, Actions, Issues, Dependencies, Decisions, Assumptions
- **Real-time persistence** — all data saved to Supabase Postgres
- **Excel export** — tasks and RAID log exportable to .xlsx
