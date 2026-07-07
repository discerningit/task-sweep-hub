# Deploy TaskSweep Hub (B-06)

Get HTTPS so you can **install the PWA on your phone** and use **M365 sign-in** with a real redirect URI.

---

## Option A: Cloudflare Pages (recommended — simpler URL)

**Result:** `https://tasksweep-hub.pages.dev` (or your custom domain)

### Steps

1. Push this project to GitHub (see below if you haven't).
2. Create a free account at [cloudflare.com](https://dash.cloudflare.com/sign-up).
3. Go to **Workers & Pages → Create → Pages → Connect to Git**.
4. Select your `task-sweep-hub` repository.
5. Build settings:

   | Setting | Value |
   |---------|-------|
   | Framework preset | None (or Vite) |
   | Build command | `npm run build` |
   | Build output directory | `dist` |
   | Node version | 22 |

6. Click **Save and Deploy**.
7. When done, open your `*.pages.dev` URL — you should see TaskSweep Hub.
8. **Install PWA:** Chrome menu → **Install TaskSweep Hub** (or the install icon in the address bar).

### M365 redirect URI (after deploy)

In Azure Portal → your app → **Authentication** → add:

```
https://tasksweep-hub.pages.dev
```

(Use your actual Cloudflare URL if different.)

Then add that same URL in TaskSweep **Settings** before signing in.

---

## Option B: GitHub Pages (free, no extra account)

**Result:** `https://<your-username>.github.io/task-sweep-hub/`

A deploy workflow is already included (`.github/workflows/deploy-github-pages.yml`).

### Steps

1. Create a GitHub repository named `task-sweep-hub`.
2. Push your code:

   ```bash
   cd task-sweep-hub
   git remote add origin https://github.com/<your-username>/task-sweep-hub.git
   git push -u origin main
   ```

3. On GitHub: **Settings → Pages**.
4. Under **Build and deployment**, set **Source** to **GitHub Actions**.
5. The workflow runs automatically on every push to `main`.
6. Wait ~2 minutes, then open:

   ```
   https://<your-username>.github.io/task-sweep-hub/
   ```

7. **Install PWA** from Chrome/Edge on your phone or desktop.

### M365 redirect URI (GitHub Pages)

Add this in Azure Portal (replace `<your-username>`):

```
https://<your-username>.github.io/task-sweep-hub/
```

---

## Verify deployment

After either option:

- [ ] App loads over **https://** (lock icon in browser)
- [ ] Paste sweep works
- [ ] **Install** option appears (PWA)
- [ ] Works on phone browser
- [ ] M365 redirect URI added in Azure (if using M365)

---

## Build locally (same as CI)

```bash
# Cloudflare / custom domain at root
npm run build

# GitHub Pages project site
npm run build:gh-pages

# Preview production build
npm run preview
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Blank page on GitHub Pages | Ensure URL ends with `/task-sweep-hub/` and Pages source is **GitHub Actions** |
| M365 sign-in fails | Redirect URI in Azure must **exactly** match your deployed URL (trailing slash matters for GitHub Pages) |
| PWA won't install | Must be HTTPS; try Chrome or Edge |
| Old version after deploy | Hard refresh (`Ctrl+Shift+R`) — PWA auto-updates on next visit |

---

## Custom domain (optional)

**Cloudflare Pages:** Pages project → **Custom domains** → add `tasksweep.yourdomain.com`.

**GitHub Pages:** Repo **Settings → Pages → Custom domain**, then add DNS records at your registrar.

Update your M365 redirect URI to match the custom domain.