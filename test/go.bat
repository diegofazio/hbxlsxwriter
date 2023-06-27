@del test.exe
hbmk2 test.prg -w0 -rebuild
if ERRORLEVEL 1 goto error
call test.exe
if ERRORLEVEL 1 goto error
start libro.xlsx
goto end
:error
echo Build failed
:end