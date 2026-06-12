# Equipment Log — truck asset tracker

Log equipment (name / serial / photo / assigned truck), import the station
asset register (.xlsx) for category + name dropdowns, and export the log
back to Excel.

Built with Vite + React + Tailwind CSS. Data is saved in the browser's
localStorage, so each device keeps its own log.

## Run locally
```bash
npm install
npm run dev
```

## Deploy to Vercel (free)

Option A — easiest, no terminal:
1. Push this folder to a GitHub repository.
2. Go to vercel.com, sign up free with your GitHub account.
3. Click "Add New… > Project", pick the repo, and click Deploy.
   Vercel auto-detects Vite — no settings needed.
4. You get a free https://your-app.vercel.app URL. Every git push redeploys.

Option B — from the terminal:
```bash
npm install -g vercel
vercel
```
and follow the prompts.

## Notes
- Photos are compressed before saving, but localStorage holds roughly
  5MB per browser — fine for dozens of entries, not hundreds.
- Each phone/browser has its own separate data. A shared multi-user log
  would need a backend (e.g. Supabase free tier).
