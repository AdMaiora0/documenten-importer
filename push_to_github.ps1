$gh = "c:\Dev\documenten-importer\tools\bin\gh.exe"
$git = "C:\Program Files\Git\cmd\git.exe"

Write-Host "--- GitHub Setup ---" -ForegroundColor Cyan

# 1. Check Auth
Write-Host "1. Controleren op GitHub login..."
& $gh auth status
if ($LASTEXITCODE -ne 0) {
    Write-Host ">> Je bent nog niet ingelogd. We starten nu de login procedure." -ForegroundColor Yellow
    Write-Host ">> Kies 'GitHub.com', 'HTTPS', en 'Login with a web browser'." -ForegroundColor Yellow
    & $gh auth login
}

# 2. Create Repo
Write-Host "2. Repository aanmaken op GitHub..."
# Check if remote already exists
$remotes = & $git remote
if ($remotes -contains "origin") {
    Write-Host ">> Remote 'origin' bestaat al." -ForegroundColor Gray
} else {
    & $gh repo create documenten-importer --public --source=. --remote=origin
}

# 3. Push
Write-Host "3. Code pushen..."
& $git push -u origin master

Write-Host "--- Klaar! ---" -ForegroundColor Green
