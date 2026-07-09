# ZenkyBox Web - Deployment Guide

## Step 1: Prepare Your Local Environment

### Install Node.js (if not already installed)
- Download from: https://nodejs.org/
- Install LTS version
- Verify: `node --version` and `npm --version`

### Install Dependencies
```bash
cd zenkybox-web
npm install
```

### Test Locally
```bash
npm run dev
# Open http://localhost:3000 in your browser
```

## Step 2: Create GitHub Account & Repository

### Setup GitHub
1. Go to: https://github.com
2. Sign up or login
3. Create new repository: zenkybox-web
4. Make it Public

### Push Code to GitHub
```bash
# Initialize git (if not already done)
git init

# Add all files
git add .

# Commit
git commit -m "Initial commit - ZenkyBox Web"

# Add remote
git remote add origin https://github.com/YOUR_USERNAME/zenkybox-web.git

# Push
git branch -M main
git push -u origin main
```

## Step 3: Deploy to Vercel

### Create Vercel Account
1. Go to: https://vercel.com
2. Click: Sign Up
3. Choose: GitHub
4. Authorize Vercel

### Import Project
1. Click: New Project
2. Click: Import Git Repository
3. Select: zenkybox-web repository
4. Click: Import
5. Click: Deploy
6. Wait 2-3 minutes for build

### Access Your App
- Vercel provides: `https://zenkybox-web.vercel.app`
- Share this URL with your team!

## Step 4: Configure Custom Domain (Optional)

### Add Domain to Vercel
1. Go to Project Settings
2. Click: Domains
3. Add your domain (e.g., inventory.zenkybox.in)
4. Update DNS records as instructed

## Step 5: Update Your App

### Make Changes
1. Edit files locally
2. Test with: `npm run dev`
3. Commit: `git commit -am "message"`
4. Push: `git push`
5. Vercel auto-rebuilds and deploys
6. Changes live in 1-2 minutes!

## File Structure Explanation

```
zenkybox-web/
├── components/App.jsx
│   └── Main application component with all features
│
├── pages/
│   ├── _app.js           # Next.js wrapper
│   └── index.js          # Home page (loads App.jsx)
│
├── styles/globals.css    # All styling (1100+ lines)
│
├── public/               # Static files (images, etc.)
│
├── package.json          # Dependencies
├── next.config.js        # Next.js configuration
├── vercel.json           # Vercel deployment config
├── .gitignore            # Git ignore rules
├── .env.example          # Environment variables template
├── README.md             # Project info
└── DEPLOYMENT.md         # This file
```

## Troubleshooting

### Build Failed on Vercel
**Solution:**
1. Check package.json is valid JSON
2. Verify all file paths are correct
3. Check for missing dependencies
4. Try rebuilding: Settings → Redeploy

### App Loads But Shows Blank
**Solution:**
1. Wait 5 seconds (initial load)
2. Refresh page (F5)
3. Check browser console (F12)
4. Clear browser cache

### Cannot Push to GitHub
**Solution:**
1. Verify git is installed: `git --version`
2. Configure git: 
   ```bash
   git config --global user.name "Your Name"
   git config --global user.email "your@email.com"
   ```
3. Check remote: `git remote -v`
4. Re-authenticate on GitHub

### Images Not Showing
**Solution:**
1. Images stored in browser storage
2. Clear browser cache
3. Re-upload images
4. Try different browser

## Performance Tips

1. **Images:** Keep images under 500KB
2. **Code:** Minimize large libraries
3. **Cache:** Let Vercel handle caching
4. **CDN:** Vercel uses global CDN automatically

## Security

- HTTPS enabled by default ✓
- No sensitive data in code ✓
- Environment variables for secrets ✓
- CORS headers configured ✓

## Support

For issues:
1. Check browser console (F12)
2. Review Vercel deployment logs
3. Check GitHub for code issues
4. Verify dependencies in package.json

## Next Steps

1. ✅ Deploy to Vercel (20 minutes)
2. Share URL with team
3. Start using the app
4. Add Firebase for cloud backup (optional)
5. Add user authentication (optional)

---

**You're ready to go live!** 🚀

Questions? Check README.md or ../WEB_VERSION_DEPLOYMENT_GUIDE.txt
