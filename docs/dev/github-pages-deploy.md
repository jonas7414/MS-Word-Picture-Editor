# GitHub Pages Deployment

## What Gets Deployed

GitHub Pages hosts the built task pane app from `dist/`.

The deployment also generates:

```text
dist/manifest.xml
```

Use that generated manifest when sideloading the Word add-in on another device.

## Repository Setup

1. Push this repository to GitHub.
2. In GitHub, open the repository settings.
3. Go to `Pages`.
4. Set `Build and deployment` source to `GitHub Actions`.
5. Push to `master`, or manually run the `Deploy GitHub Pages` workflow.

The app URL will be:

```text
https://<github-owner>.github.io/<repository-name>/
```

The generated add-in manifest URL will be:

```text
https://<github-owner>.github.io/<repository-name>/manifest.xml
```

## Local Pages Build Check

Replace the URL with your actual GitHub Pages URL:

```bash
GITHUB_PAGES_BASE=/word_pic_editor/ \
GITHUB_PAGES_URL=https://<github-owner>.github.io/word_pic_editor \
npm run build:pages
```

Then inspect:

```text
dist/index.html
dist/manifest.xml
```

`dist/manifest.xml` should not contain `localhost`.

## Use on Another Device

1. Open the generated `manifest.xml` URL in a browser and save it, or download it from the GitHub Pages deployment artifact.
2. In Word, sideload that manifest.
3. Open the `圖片標註` add-in from the Home ribbon.

The manifest must point to HTTPS. GitHub Pages satisfies this requirement.
