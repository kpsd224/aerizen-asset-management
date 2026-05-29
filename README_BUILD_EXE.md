# Build Aerizen Windows EXE from GitHub

## Fixed build problem
The previous GitHub Actions build failed at `npm run build` because Tailwind was using `latest`.
`latest` installed Tailwind CSS v4, but the project used a Tailwind CSS v3 PostCSS config.
This caused the Vite CSS error:

`It looks like you're trying to use tailwindcss directly as a PostCSS plugin.`

This version fixes the problem by pinning Tailwind CSS to `3.4.17` and other build packages to stable versions.

## How to build
1. Upload all files in this folder to the root of your GitHub repository.
2. Make sure this path exists:

```txt
.github/workflows/build-windows-exe.yml
```

3. Go to GitHub → Actions → Build Windows EXE → Run workflow.
4. After the workflow finishes, open the workflow result and download:

```txt
Aerizen-Asset-Management-Windows-Installer
```

The installer EXE is generated from:

```txt
release/*.exe
```
