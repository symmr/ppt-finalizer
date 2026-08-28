# Deploy ppt-finalizer to GitHub and enable GitHub Pages (/docs)
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("Path", "User")

Write-Host "Sync web -> docs..."
Copy-Item -Force "$Root\web\replace-fonts.html" "$Root\docs\index.html"

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) not found. Run: winget install GitHub.cli"
}

gh auth status | Out-Null

git add docs/index.html
$dirty = git status --porcelain
if ($dirty) {
    git commit -m "Sync docs/index.html from web/replace-fonts.html"
}

$originUrl = git remote get-url origin 2>$null
if (-not $originUrl) {
    Write-Host "Creating public repo symmr/ppt-finalizer ..."
    gh repo create symmr/ppt-finalizer --public --source=. --remote=origin --push --description "Browser-side PowerPoint (.pptx) finalizer"
} else {
    Write-Host "Pushing to origin..."
    git push -u origin main
}

Write-Host "Enabling GitHub Pages (branch main, folder /docs)..."
gh api repos/symmr/ppt-finalizer/pages -X POST `
    -f build_type=legacy `
    -f "source[branch]=main" `
    -f "source[path]=/docs" 2>$null

if ($LASTEXITCODE -ne 0) {
    Write-Host "Pages POST skipped; trying PUT update..."
    gh api repos/symmr/ppt-finalizer/pages -X PUT `
        -f build_type=legacy `
        -f "source[branch]=main" `
        -f "source[path]=/docs" 2>$null
}

Write-Host ""
Write-Host "Done. Site (may take 1-2 min to build):"
Write-Host "  https://symmr.github.io/ppt-finalizer/"
Write-Host "Repo:"
Write-Host "  https://github.com/symmr/ppt-finalizer"
