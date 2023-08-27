@echo off
set /p Var1= "Commit String:"
@echo on
git add .
git commit -m "%Var1%"
git push
pause