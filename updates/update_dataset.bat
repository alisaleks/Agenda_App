@echo off
echo Changing to app directory
cd "C:\Users\aaleksan\OneDrive - Amplifon S.p.A\Documentos\python_alisa\saturation\Saturation\Satapp\agenda_app" || (echo Failed to change directory & exit /b 1)

echo Adding files to Git
git add . || (echo Failed to add shiftslots.xlsx & exit /b 1)

echo Committing changes
git commit -m "Automatic dataset update" || (echo Nothing to commit or commit failed & exit /b 1)

echo Pulling latest changes from remote
git pull origin main --rebase || (echo Failed to pull changes & exit /b 1)

echo Pushing to GitHub
git push origin main || (echo Push failed & exit /b 1)

echo Push completed successfully
pause
exit /b 0
