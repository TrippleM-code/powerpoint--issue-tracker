# P1 Development Deployment

Keep the existing alpha deployment untouched.

Recommended production-development deployment:
1. Use the `production/` directory as a separate Vercel project root.
2. Vercel build command: `npm run build`
3. Vercel output directory: `dist`
4. After Vercel gives the HTTPS URL, update:
   `manifests/issueflow-production-dev.xml`
5. Replace `https://YOUR-PRODUCTION-DEV-URL.example.com/` with the real HTTPS URL.
6. Sideload the production-development manifest separately from Alpha.

This keeps:
- Alpha = stable reference
- Production Dev = new modular application
