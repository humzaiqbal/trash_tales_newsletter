# Trash Tales Newsletter Site

Static newsletter site generated from DOCX files in `~/Downloads/newsletter`, designed for GitHub Pages.

## Rebuild the site

```bash
python3 scripts/build_site.py
```

The generated website is written to:

- `site/index.html`
- `site/posts/*.html`
- `site/assets/styles.css`
- `site/assets/app.js`

## Publish on GitHub Pages (`humzaiqbal`)

1. Create a new repository on GitHub, for example `trash-tales-newsletter`.
2. From this folder, initialize and push:

```bash
git init
git add .
git commit -m "Initial newsletter site"
git branch -M main
git remote add origin git@github.com:humzaiqbal/trash-tales-newsletter.git
git push -u origin main
```

3. In GitHub repo settings:
   - Open **Pages**
   - Set **Source** to **Deploy from a branch**
   - Choose `main` branch and `/site` folder
4. Your site will appear at:
   - `https://humzaiqbal.github.io/trash-tales-newsletter/`

For a root user site (`https://humzaiqbal.github.io/`), use repository name `humzaiqbal.github.io` and move generated files from `site/` to repository root.
