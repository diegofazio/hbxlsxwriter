@del .\lib\x64\xlsxwriter.lib
@del .\lib\x32\xlsxwriter.lib
@del ..\hbxlsxwriter\libxlsxwriter\lib\x64\xlsxwriter.lib
@del ..\hbxlsxwriter\libxlsxwriter\lib\x32\xlsxwriter.lib
@hbmk2 libxlsxwriter.hbp -rebuild
@copy .\lib\x64\xlsxwriter.lib \harbour\lib\win\msvc64
@copy .\lib\x32\xlsxwriter.lib \harbour\lib\win\msvc