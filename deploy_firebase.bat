@echo off
echo === BYD Search - Firebase Hosting Deployment ===
echo.
echo Step 1: Login to Firebase (if not already logged in)
firebase login
echo.
echo Step 2: Deploying to Firebase Hosting...
firebase deploy --only hosting --project project-d5a525d1-72ba-431c-80c
echo.
echo Done! Your site should now be available on Firebase Hosting.
pause
