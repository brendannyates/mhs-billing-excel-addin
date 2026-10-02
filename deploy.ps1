param([string]$Owner, [string]$Repo = "mhs-billing-excel-addin")
$ErrorActionPreference = "Stop"
if (-not $Owner) { $Owner = gh api user --jq .login }
if (-not $Owner) { throw "Sign in using gh auth login first." }
$url = "https://$Owner.github.io/$Repo/"
$manifest = Get-Content manifest.xml -Raw
$manifest = $manifest.Replace("https://GITHUB_OWNER.github.io/mhs-billing-excel-addin/", $url)
Set-Content manifest.xml $manifest -Encoding utf8
npm test
if ($LASTEXITCODE -ne 0) { throw "Tests failed" }
git init -b main
git add index.html app.js core.js style.css manifest.xml package.json tests office-scripts README.md AUTOMATION.md .github deploy.ps1
git commit -m "Create MHS billing ticket Excel add-in"
gh repo create "$Owner/$Repo" --public --source . --remote origin --push
if ($LASTEXITCODE -ne 0) { throw "Repository creation failed" }
gh api --method POST "repos/$Owner/$Repo/pages" -f build_type=workflow
if ($LASTEXITCODE -ne 0) { throw "Enable Pages using GitHub Actions in repository Settings > Pages, then rerun the workflow." }
gh workflow run pages.yml --repo "$Owner/$Repo"
Write-Host "Pages URL: $url"
Write-Host "Upload manifest.xml in Excel after the deployment completes."
