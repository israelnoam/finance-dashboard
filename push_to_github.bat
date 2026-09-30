@echo off
title Push Finance Dashboard to GitHub
echo ========================================================
echo   Pushing Finance Dashboard to GitHub
echo   Repository: https://github.com/israelnoam/finance-dashboard
echo ========================================================
echo.
echo If your repository is private or requires authentication,
echo you will need a GitHub Personal Access Token with 'repo' scope.
echo Generate one at: https://github.com/settings/tokens/new
echo.
set /p GITHUB_TOKEN="Enter your GitHub Personal Access Token (or press Enter to try default): "

if "%GITHUB_TOKEN%"=="" (
    .\mingit\cmd\git.exe push -u origin main
) else (
    .\mingit\cmd\git.exe push -u https://%GITHUB_TOKEN%@github.com/israelnoam/finance-dashboard.git main
)

echo.
echo ========================================================
echo   Done! Check https://github.com/israelnoam/finance-dashboard
echo ========================================================
pause
