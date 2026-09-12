Vercel static-hosting fix

Upload this structure to the ROOT of the GitHub repository:

public/
  index.html
  taskpane.html
  taskpane.js
  taskpane.css
  icon-32.png
vercel.json

Then in Vercel:
- Framework Preset: Other
- Build Command: leave empty
- Output Directory: leave the dashboard override OFF/empty, because vercel.json sets it to public
- Root Directory: repository root
- Redeploy

Test:
https://<your-production-domain>/
https://<your-production-domain>/taskpane.html
