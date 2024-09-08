@echo off
echo Changing to app directory
cd "C:\Users\aaleksan\OneDrive - Amplifon S.p.A\Documentos\python_alisa\saturation\Saturation\Satapp\agenda_app" || (echo Failed to change directory & exit /b 1)

echo Adding all modified files to Git
git add . || (echo Failed to add files & exit /b 1)

echo Committing changes
git commit -m "Automatic dataset update" || (echo Nothing to commit or commit failed & exit /b 1)

echo Pushing to GitHub
git push origin main || (echo Push failed & exit /b 1)

echo Push completed successfully
pause  # This will keep the terminal open so you can read the output
